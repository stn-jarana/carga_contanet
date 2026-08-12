from pywinauto import Application, mouse
from pywinauto.keyboard import send_keys
import win32clipboard
import time
import pandas as pd
import json
import os
import csv
import subprocess
import calendar
import dotenv

from datetime import datetime
import sys

dotenv.load_dotenv()  # Cargar variables de entorno desde .env

# Redireccionar stdout y stderr a consola y a archivo .txt de logs
class LoggerTee:
    def __init__(self, filename="ejecucion_log.txt"):
        self.terminal = sys.stdout
        self.log = open(filename, "a", encoding="utf-8")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()

    def flush(self):
        self.terminal.flush()
        self.log.flush()

sys.stdout = LoggerTee("ejecucion_log.txt")
sys.stderr = sys.stdout

ARCHIVO_PROGRESO = "progreso.json"
ARCHIVO_REPORTE = "reporte_asientos.csv"
ARCHIVO_METRICAS = "reporte_metricas.csv"

# Globales de Métricas por Día
TIEMPO_INICIO_GLOBAL = datetime.now()
METRICAS_DIARIAS = {}

def _obtener_fecha_hoy():
    return datetime.now().strftime("%Y-%m-%d")

def _obtener_metricas_dia(fecha=None):
    if fecha is None:
        fecha = _obtener_fecha_hoy()
    if fecha not in METRICAS_DIARIAS:
        METRICAS_DIARIAS[fecha] = {
            "operaciones_evaluadas": 0,
            "actualizados_exito": 0,
            "no_modificados": 0,
            "voucher_no_modificable": 0,
            "asignado_a_caja": 0,
            "periodo_cerrado": 0,
            "no_encontrado_excel": 0,
            "sin_movimientos_mes": 0,
            "errores_otros": 0
        }
    return METRICAS_DIARIAS[fecha]

def inicializar_metricas():
    if not os.path.exists(ARCHIVO_METRICAS):
        with open(ARCHIVO_METRICAS, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "Fecha",
                "Total_Asientos_Evaluados",
                "Actualizados_Exito",
                "No_Modificados_Total",
                "Voucher_No_Modificable",
                "Asignado_A_Caja",
                "Periodo_Cerrado",
                "No_Encontrado_Excel",
                "Sin_Movimientos_Mes",
                "Porcentaje_Efectividad"
            ])

def generar_reporte_metricas():
    """Consolida y actualiza en reporte_metricas.csv las métricas agrupadas por DÍA."""
    inicializar_metricas()
    
    # Cargar métricas existentes en el CSV para actualizar por fecha
    registros_diarios = {}
    if os.path.exists(ARCHIVO_METRICAS):
        try:
            with open(ARCHIVO_METRICAS, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    fecha_row = row.get("Fecha") or (row.get("Fecha_Hora_Reporte", "").split(" ")[0])
                    if fecha_row:
                        registros_diarios[fecha_row] = row
        except Exception:
            pass

    # Actualizar / insertar datos desde METRICAS_DIARIAS
    for fecha, m in METRICAS_DIARIAS.items():
        total_ops = m["operaciones_evaluadas"]
        exito = m["actualizados_exito"]
        no_mod = m["no_modificados"]
        efectividad = round((exito / total_ops) * 100, 2) if total_ops > 0 else 0.0
        
        # Si ya había un registro en CSV para este día y no hemos inicializado desde él en esta sesión,
        # nos aseguramos de que el registro refleje la información diaria acumulada.
        registros_diarios[fecha] = {
            "Fecha": fecha,
            "Total_Asientos_Evaluados": str(total_ops),
            "Actualizados_Exito": str(exito),
            "No_Modificados_Total": str(no_mod),
            "Voucher_No_Modificable": str(m["voucher_no_modificable"]),
            "Asignado_A_Caja": str(m["asignado_a_caja"]),
            "Periodo_Cerrado": str(m["periodo_cerrado"]),
            "No_Encontrado_Excel": str(m["no_encontrado_excel"]),
            "Sin_Movimientos_Mes": str(m["sin_movimientos_mes"]),
            "Porcentaje_Efectividad": f"{efectividad}%"
        }

    # Escribir el archivo CSV manteniendo 1 sola fila por día
    fieldnames = [
        "Fecha",
        "Total_Asientos_Evaluados",
        "Actualizados_Exito",
        "No_Modificados_Total",
        "Voucher_No_Modificable",
        "Asignado_A_Caja",
        "Periodo_Cerrado",
        "No_Encontrado_Excel",
        "Sin_Movimientos_Mes",
        "Porcentaje_Efectividad"
    ]
    try:
        with open(ARCHIVO_METRICAS, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for fecha_key in sorted(registros_diarios.keys()):
                writer.writerow(registros_diarios[fecha_key])
        
        fecha_hoy = _obtener_fecha_hoy()
        if fecha_hoy in METRICAS_DIARIAS:
            m_hoy = METRICAS_DIARIAS[fecha_hoy]
            tot = m_hoy["operaciones_evaluadas"]
            ex = m_hoy["actualizados_exito"]
            ef = round((ex / tot) * 100, 2) if tot > 0 else 0.0
            print(f"[MÉTRICAS DIARIAS ACTUALIZADAS] Fecha: {fecha_hoy} | Total Día: {tot} ops | Éxito: {ex} | Efectividad: {ef}%\n")
    except Exception as e:
        print(f"Error escribiendo reporte de métricas por día: {e}")

def inicializar_reporte():
    if not os.path.exists(ARCHIVO_REPORTE):
        with open(ARCHIVO_REPORTE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["RUC", "Empresa", "Año", "Asiento", "Número de Operación", "Estado", "Observación"])

def registrar_reporte(ruc, nombre_empresa, anio, asiento, estado, observacion="", numero_operacion=""):
    inicializar_reporte()
    try:
        with open(ARCHIVO_REPORTE, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([ruc, nombre_empresa, anio, asiento, numero_operacion, estado, observacion])
        
        # Actualización de Métricas Diarias
        m = _obtener_metricas_dia()
        m["operaciones_evaluadas"] += 1
        if estado == "ACTUALIZADO":
            m["actualizados_exito"] += 1
        else:
            m["no_modificados"] += 1
            if observacion == "Voucher no modificable":
                m["voucher_no_modificable"] += 1
            elif "asignado a caja" in observacion.lower():
                m["asignado_a_caja"] += 1
            elif "periodo cerrado" in observacion.lower():
                m["periodo_cerrado"] += 1
            elif "no se encontró" in observacion.lower():
                m["no_encontrado_excel"] += 1
            elif "no hay registros" in observacion.lower():
                m["sin_movimientos_mes"] += 1
            else:
                m["errores_otros"] += 1

        # Actualizar reporte de métricas por día periódicamente (cada 5 operaciones)
        if m["operaciones_evaluadas"] % 5 == 0:
            generar_reporte_metricas()

    except Exception as e:
        print(f"Error registrando en el reporte: {e}")



def cargar_progreso():
    if os.path.exists(ARCHIVO_PROGRESO):
        try:
            with open(ARCHIVO_PROGRESO, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error cargando progreso: {e}")
            return {}
    return {}

def guardar_progreso(progreso):
    try:
        with open(ARCHIVO_PROGRESO, "w", encoding="utf-8") as f:
            json.dump(progreso, f, indent=4)
    except Exception as e:
        print(f"Error guardando progreso: {e}")

def _clave_mes(anio, mes):
    return f"{anio}_{mes:02d}"

def registrar_asiento_procesado(ruc, anio, mes, asiento):
    progreso = cargar_progreso()
    if ruc not in progreso:
        progreso[ruc] = {}
    
    clave = _clave_mes(anio, mes)
    if clave not in progreso[ruc]:
        progreso[ruc][clave] = {"asientos": [], "ultima_fila": 0}
    
    if isinstance(progreso[ruc][clave], list):
        progreso[ruc][clave] = {"asientos": progreso[ruc][clave], "ultima_fila": 0}
        
    if asiento not in progreso[ruc][clave]["asientos"]:
        progreso[ruc][clave]["asientos"].append(asiento)
        guardar_progreso(progreso)

def es_asiento_procesado(ruc, anio, mes, asiento):
    progreso = cargar_progreso()
    clave = _clave_mes(anio, mes)
    if ruc in progreso and clave in progreso[ruc]:
        dato = progreso[ruc][clave]
        if isinstance(dato, list):
            return asiento in dato
        return asiento in dato.get("asientos", [])
    return False

def guardar_ultima_fila(ruc, anio, mes, fila):
    progreso = cargar_progreso()
    if ruc not in progreso:
        progreso[ruc] = {}
    clave = _clave_mes(anio, mes)
    if clave not in progreso[ruc]:
        progreso[ruc][clave] = {"asientos": [], "ultima_fila": 0}
    if isinstance(progreso[ruc][clave], list):
        progreso[ruc][clave] = {"asientos": progreso[ruc][clave], "ultima_fila": 0}
    progreso[ruc][clave]["ultima_fila"] = fila
    guardar_progreso(progreso)

def obtener_ultima_fila(ruc, anio, mes):
    progreso = cargar_progreso()
    clave = _clave_mes(anio, mes)
    if ruc in progreso and clave in progreso[ruc]:
        dato = progreso[ruc][clave]
        if isinstance(dato, dict):
            return dato.get("ultima_fila", 0)
    return 0

def marcar_mes_sin_movimientos(ruc, anio, mes):
    """Marca un mes como revisado y sin movimientos en el JSON de progreso."""
    progreso = cargar_progreso()
    if ruc not in progreso:
        progreso[ruc] = {}
    clave = _clave_mes(anio, mes)
    if clave not in progreso[ruc]:
        progreso[ruc][clave] = {"asientos": [], "ultima_fila": 0}
    if isinstance(progreso[ruc][clave], list):
        progreso[ruc][clave] = {"asientos": progreso[ruc][clave], "ultima_fila": 0}
    progreso[ruc][clave]["sin_movimientos"] = True
    guardar_progreso(progreso)

def es_mes_sin_movimientos(ruc, anio, mes):
    """Verifica si un mes ya fue revisado y marcado como sin movimientos."""
    progreso = cargar_progreso()
    clave = _clave_mes(anio, mes)
    if ruc in progreso and clave in progreso[ruc]:
        dato = progreso[ruc][clave]
        if isinstance(dato, dict):
            return dato.get("sin_movimientos", False)
    return False

ruta_exe = (
    r"C:\Ejecutable ContaNet ERP 3.0.7.44 - SOUTHERN"
    r"\ContaNet.Aplicacion.exe"
)

directorio_trabajo = (
    r"C:\Ejecutable ContaNet ERP 3.0.7.44 - SOUTHERN"
)

# ── Reinicio programado ────────────────────────────────────────────────────────
# Define las 2 horas del día en que la app se reinicia sola (hora, minuto).
HORAS_REINICIO_PROGRAMADO = [
    (1, 0),   # 1:0 AM
    (23, 0),   # 11:00 PM
]
# ──────────────────────────────────────────────────────────────────────────────

EXCEL_FILE = r"\\192.168.30.36\Tesoreria\TESORERIA\CONSTANCIAS DE PAGO\constancias_de_pago.xlsx"

excel_data = pd.read_excel(
    EXCEL_FILE,
    sheet_name=None
)

hojas_excel = {
    "20506883301": "CMT",
    "20600692781": "DIONISO",
    "20514016624": "DYNAMITEX",
    "20490242407": "INFOSUR",
    "20494530865": "DINSURA",
    "20606955724": "PERU COMMERCE",
    "20376729126": "STN",    
    "20504334041": "ITS",
    "20609254778": "TECA",
    "20542813238": "IÑAPARI"
}


empresas = {
    "20506883301": "CMT",
    "20606955724": "PERU COMMERCE",
    "20494530865": "DINSURA",
    "20542813238": "IÑAPARI",
    "20609254778": "TECA",
    "20376729126": "STN"
}

FECHAS_INICIO = {
    "20376729126": {  # STN
        2026: {
            1: 30   # enero 2026: iniciar desde el día 28
        }
    }
}

# Configuración del mes inicial por RUC y año (ej: STN desde febrero 2026 en adelante)
MESES_INICIO = {
    "20506883301": {  # STN
        2026: 5       # En 2026 comenzar desde el mes 2 (febrero)
    }
}

anio = [2026, 2025, 2024, 2023, 2022, 2021]
mes = [1, 12]


def obtener_ventana(app_obj, timeout=20):
    inicio = time.time()
    ultimo_error = None

    while time.time() - inicio < timeout:
        try:
            ventanas = app_obj.windows()
            if ventanas:
                for ventana in ventanas:
                    try:
                        titulo = ventana.window_text().lower()
                    except Exception:
                        continue

                    if "login" in titulo:
                        return ventana

                for ventana in ventanas:
                    if ventana.is_visible():
                        return ventana
        except Exception as exc:
            ultimo_error = exc

        time.sleep(0.5)

    raise RuntimeError(f"No se encontró ninguna ventana de la aplicación: {ultimo_error}")


def iniciar_aplicacion():
    return Application(backend="uia").start(
        cmd_line=ruta_exe,
        work_dir=directorio_trabajo,
    )


def forzar_cierre_procesos(nombre_exe="ContaNet.Aplicacion.exe"):
    """Mata por nombre todas las instancias del proceso para evitar
    que instancias huérfanas impidan el reinicio de la aplicación."""
    try:
        resultado = subprocess.run(
            ["taskkill", "/F", "/IM", nombre_exe],
            capture_output=True,
            text=True,
        )
        if resultado.returncode == 0:
            print(f"Proceso '{nombre_exe}' terminado forzosamente.")
        else:
            # returncode 128 = proceso no encontrado (ya cerrado)
            print(f"No se encontraron instancias de '{nombre_exe}' (puede que ya estuviera cerrado).")
    except Exception as e:
        print(f"Error al intentar forzar cierre de '{nombre_exe}': {e}")


def cerrar_aplicacion(app):

    try:

        principal = app.window(
            title_re=".*ContaNet ERP.*"
        )

        principal.close()

        time.sleep(2)

        for ctrl in principal.descendants():

            try:

                texto = ctrl.window_text().strip()

                if texto == "Sí":

                    print(
                        "PULSANDO SI"
                    )

                    ctrl.click_input()

                    time.sleep(3)

                    break

            except:
                pass

        else:
            print(
                "NO SE ENCONTRO BOTON SI"
            )

    except Exception as e:

        print(
            "Error cerrando aplicación:",
            e
        )

    finally:
        # Siempre asegurarse de que no quede ninguna instancia abierta
        time.sleep(1)
        forzar_cierre_procesos()
        time.sleep(2)
    return True


# ── Helpers de reinicio programado ────────────────────────────────────────────
_horas_ya_reiniciadas: set = set()


def debe_reiniciar_por_horario() -> bool:
    """Devuelve True si la hora actual coincide (±1 min) con alguna de
    las horas programadas en HORAS_REINICIO_PROGRAMADO y todavía no se
    ha disparado el reinicio para esa hora en este día."""
    ahora = datetime.now()
    clave_dia = ahora.date()

    for hora, minuto in HORAS_REINICIO_PROGRAMADO:
        clave = (clave_dia, hora, minuto)
        if clave in _horas_ya_reiniciadas:
            continue  # ya se reinició en esta hora hoy
        diff = abs((ahora.hour * 60 + ahora.minute) - (hora * 60 + minuto))
        if diff <= 1:  # ventana de ±1 minuto
            _horas_ya_reiniciadas.add(clave)
            return True
    return False


def reiniciar_app_programado(app, ruc, anio_actual, mes_actual, dia_desde):
    """Cierra la instancia actual y vuelve a iniciarla en el mismo punto.
    Retorna el nuevo objeto `app` listo para continuar."""
    hora_str = datetime.now().strftime("%H:%M")
    print(f"[REINICIO PROGRAMADO {hora_str}] Cerrando instancia actual...")
    cerrar_aplicacion(app)
    time.sleep(3)

    print(f"[REINICIO PROGRAMADO {hora_str}] Iniciando nueva instancia...")
    app = iniciar_aplicacion()
    ventana_login = obtener_ventana(app)
    ventana_login = app.window(title_re=".*Login.*")
    iniciar_sesion(ventana_login)
    time.sleep(2)
    ventana_empresa_win = app.window(title_re=".*Seleccionar Empresa.*")
    seleccionar_empresa_y_anio(ventana_empresa_win, ruc, anio_actual)
    time.sleep(2)
    seleccionar_tesoreria_explorador(app)
    time.sleep(5)
    probar_fechas(app, anio_actual, mes_actual, dia_desde=dia_desde)
    time.sleep(7)
    send_keys("^c")
    time.sleep(1)
    print(f"[REINICIO PROGRAMADO {hora_str}] Nueva instancia lista.")
    return app
# ──────────────────────────────────────────────────────────────────────────────


def iniciar_sesion(ventana):
    user = os.getenv('user')
    password = os.getenv('password')
    ventana.child_window(auto_id="txtUsuario", control_type="Edit").set_text(user)
    ventana.child_window(auto_id="txtContrasena", control_type="Edit").set_text(password)

    boton = ventana.child_window(auto_id="bAceptar", control_type="Button")
    boton.click_input()


def es_control_visible(ctrl):
    try:
        if not getattr(ctrl, 'exists', lambda: True)():
            return False
    except Exception:
        return False
    try:
        return getattr(ctrl, 'is_visible', lambda: True)()
    except Exception:
        return False


def buscar_control_por_texto(ventana, texto, control_type=None):
    texto_buscar = texto.lower()
    controles = ventana.descendants() if control_type is None else ventana.descendants(control_type=control_type)

    for ctrl in controles:
        try:
            if not es_control_visible(ctrl):
                continue

            valores = []
            try:
                valores.append(ctrl.window_text())
            except Exception:
                pass

            try:
                valores.extend(ctrl.texts())
            except Exception:
                pass

            try:
                valores.append(ctrl.element_info.name)
            except Exception:
                pass

            try:
                valores.append(ctrl.element_info.automation_id)
            except Exception:
                pass

            texto_ctrl = " ".join([str(v) for v in valores if v]).strip().lower()
            if texto_buscar in texto_ctrl:
                return ctrl
        except Exception:
            continue

    return None


def imprimir_controles(ventana, control_type):
    # debug helper removed; kept stub for compatibility
    return


def encontrar_boton_seleccionar(ventana):
    for texto in ["Seleccionar", "Select", "Aceptar"]:
        boton = buscar_control_por_texto(ventana, texto, "Button")
        if boton is not None:
            return boton

    boton = buscar_control_por_texto(ventana, "Seleccionar")
    if boton is not None:
        return boton

    for boton in ventana.descendants(control_type="Button"):
        try:
            if boton.exists() and boton.is_visible():
                return boton
        except Exception:
            continue

    return None


def buscar_y_click_texto(ventana, texto, doble=False):
    ctrl = buscar_control_por_texto(ventana, texto, control_type="Button")
    if ctrl is None:
        ctrl = buscar_control_por_texto(ventana, texto)
    if ctrl is None:
        return False
    try:
        ctrl.set_focus()
        time.sleep(0.2)
        if doble:
            ctrl.double_click_input()
        else:
            ctrl.click_input()
        return True
    except Exception:
        try:
            rect = ctrl.rectangle()
            x = rect.left + rect.width() // 2
            y = rect.top + rect.height() // 2
            hacer_click_en_coordenadas(x, y)
            return True
        except Exception:
            return False



def encontrar_treeview(ventana):

    trees = ventana.descendants(control_type="Tree")
    for t in trees:
        try:
            return t
        except:
            pass

    return None



def seleccionar_nodo_arbol(ventana, texto_objetivo=None, indice=0, doble=False):
    tree = encontrar_treeview(ventana)
    if tree is None:
        return False

    tree.set_focus()
    time.sleep(0.3)

    # Recolectar candidatos buscando en varios tipos de control dentro del Tree
    candidates = []
    for ctrl in tree.descendants():
        try:
            # Skip invisible controls
            visible = getattr(ctrl, "is_visible", lambda: False)()
            if not visible:
                continue

            valores = []
            try:
                valores.append(ctrl.window_text())
            except Exception:
                pass
            try:
                valores.extend(ctrl.texts())
            except Exception:
                pass
            try:
                valores.append(ctrl.element_info.name)
            except Exception:
                pass
            try:
                valores.append(ctrl.element_info.automation_id)
            except Exception:
                pass

            texto_ctrl = " ".join([str(v) for v in valores if v]).strip().lower()
            if not texto_ctrl:
                continue
            # filtrar filas de cabecera/placeholder que no interesan
            if "filter" in texto_ctrl or "nodo-100000" in texto_ctrl:
                continue

            candidates.append((texto_ctrl, ctrl))
        except Exception:
            continue

    objetivo = texto_objetivo.lower() if texto_objetivo else None

    # Primero intentar coincidencia por texto completo/parcial
    if objetivo:
        for texto_ctrl, ctrl in candidates:
            if objetivo in texto_ctrl:
                try:
                    ctrl.set_focus()
                    time.sleep(0.1)
                    if hasattr(ctrl, 'type_keys'):
                        ctrl.type_keys('{ENTER}')
                    elif doble and hasattr(ctrl, 'double_click_input'):
                        ctrl.double_click_input()
                    elif hasattr(ctrl, 'click_input'):
                        ctrl.click_input()
                    else:
                        rect = ctrl.rectangle()
                        hacer_click_en_coordenadas(rect.left + rect.width()//2, rect.top + rect.height()//2, espera=0.15)
                    return True
                except Exception:
                    try:
                        rect = ctrl.rectangle()
                        hacer_click_en_coordenadas(rect.left + rect.width()//2, rect.top + rect.height()//2, espera=0.15)
                        print(f"Elemento seleccionado por coordenadas (texto): {texto_ctrl}")
                        return True
                    except Exception:
                        continue
        print(f"No se encontró el texto '{texto_objetivo}' en los candidatos. No se realiza selección por índice.")
        return False

    # Si no hay coincidencias, intentar seleccionar por índice relativo dentro de los candidatos visibles
    if candidates:
        # preferir el índice solicitado, sino probar los primeros visibles
        order = [indice] + list(range(len(candidates)))
        seen = set()
        for idx in order:
            if idx in seen:
                continue
            seen.add(idx)
            if 0 <= idx < len(candidates):
                texto_ctrl, ctrl = candidates[idx]
                try:
                    ctrl.set_focus()
                    time.sleep(0.1)
                    if hasattr(ctrl, 'type_keys'):
                        ctrl.type_keys('{ENTER}')
                    elif doble and hasattr(ctrl, 'double_click_input'):
                        ctrl.double_click_input()
                    elif hasattr(ctrl, 'click_input'):
                        ctrl.click_input()
                    else:
                        rect = ctrl.rectangle()
                        hacer_click_en_coordenadas(rect.left + rect.width()//2, rect.top + rect.height()//2, espera=0.15)
                    return True
                except Exception:
                    try:
                        rect = ctrl.rectangle()
                        hacer_click_en_coordenadas(rect.left + rect.width()//2, rect.top + rect.height()//2, espera=0.15)
                        return True
                    except Exception:
                        continue

    return False


def esperar_tree_cargado(ventana, timeout=10, min_items=4):
    inicio = time.time()
    visible_items = []
    while time.time() - inicio < timeout:
        tree = encontrar_treeview(ventana)
        if tree is None:
            time.sleep(0.5)
            continue

        visible_items = []
        for item in tree.descendants(control_type="TreeItem"):
            try:
                if not getattr(item, 'is_visible', lambda: False)():
                    continue
                visible_items.append(item)
            except Exception:
                continue

        if len(visible_items) >= min_items:
            return tree

        time.sleep(0.5)

    # timeout esperando por árbol con suficientes ítems visibles
    return tree


def click_tree_row_by_index(tree, row_index, doble=False):
    # Obtener ítems visibles y ordenarlos por posición vertical
    items = []
    for item in tree.descendants(control_type="TreeItem"):
        try:
            if not getattr(item, 'is_visible', lambda: False)():
                continue
            rect = item.rectangle()
            items.append((rect.top, rect, item))
        except Exception:
            continue

    if items:
        items.sort(key=lambda x: x[0])
        if 0 <= row_index < len(items):
            rect = items[row_index][1]
            target = items[row_index][2]
            x = rect.left + rect.width() // 2
            y = rect.top + rect.height() // 2

            try:
                target.set_focus()
                time.sleep(0.1)
                # Preferir clic directo sobre el elemento, luego presionar ENTER si está disponible.
                try:
                    target.click_input()
                except Exception:
                    mouse.move(coords=(x, y)); time.sleep(0.1); mouse.click(coords=(x, y))

                # Si el item puede recibir Enter, envíalo para activar la opción.
                try:
                    if hasattr(target, 'type_keys'):
                        target.type_keys('{ENTER}')
                except Exception:
                    pass

                # Fila clicada
                return True
            except Exception as e:
                print(f"Error al clicar fila {row_index}: {e}")
                return False

    # Si no hay TreeItem explícitos visibles, intentar calcular posición usando rect del Tree
    try:
        rect_tree = tree.rectangle()
        row_h = 36
        x = rect_tree.left + int(rect_tree.width() * 0.25)
        y = rect_tree.top + int(rect_tree.height() * 0.15) + row_index * row_h
        if doble:
            mouse.move(coords=(x, y)); time.sleep(0.1); mouse.double_click(coords=(x, y))
        else:
            hacer_click_en_coordenadas(x, y)
        # click estimado en fila por índice
        return True
    except Exception as e:
        print(f"No fue posible calcular posición para fila {row_index}: {e}")
        return False


def hacer_click_en_coordenadas(x, y, espera=0.5):
    mouse.move(coords=(x, y))
    time.sleep(espera)
    mouse.click(coords=(x, y))
    time.sleep(0.5)


def hacer_click_relativo(ventana, rel_x, rel_y, doble=False):
    rect = ventana.rectangle()
    x = rect.left + int(rect.width() * rel_x)
    y = rect.top + int(rect.height() * rel_y)
    if doble:
        mouse.move(coords=(x, y))
        time.sleep(0.2)
        mouse.double_click(coords=(x, y))
    else:
        hacer_click_en_coordenadas(x, y)
    return True


def seleccionar_empresa_y_anio(
    ventana_empresa,
    ruc,
    anio
):
    
    print(
            f"Seleccionando RUC={ruc} AÑO={anio}"
        )

    nombre_empresa = str(ruc)
    anio = str(anio)

    try:
        busqueda = ventana_empresa.child_window(auto_id="txtBusqEmp", control_type="Edit")
        busqueda.set_text("")
        busqueda.set_text(nombre_empresa)
        time.sleep(0.8)

        # Empresa seleccionada por interacción UI 
        empresa_item = buscar_control_por_texto(ventana_empresa, nombre_empresa, "ListItem")
        if empresa_item is None:
            empresa_item = buscar_control_por_texto(ventana_empresa, nombre_empresa)
        if empresa_item is not None:
            empresa_item.click_input()
        else:
            print("No se encontró el elemento de la empresa en la lista; intento clic fallback.")
            hacer_click_relativo(ventana_empresa, 0.30, 0.35)
            time.sleep(0.8)

        anio_item = buscar_control_por_texto(ventana_empresa, anio, "ListItem")
        if anio_item is None:
            anio_item = buscar_control_por_texto(ventana_empresa, anio)
        if anio_item is not None:
            anio_item.click_input()
            print(f"Año seleccionado: {anio}")
        else:
            print(f"No se encontró el elemento del año {anio} en la lista; intento clic fallback.")
            hacer_click_relativo(ventana_empresa, 0.90, 0.45)
            time.sleep(0.8)

        # Botones serán seleccionados por texto o por fallback de coordenadas
        boton_seleccionar = encontrar_boton_seleccionar(ventana_empresa)
        if boton_seleccionar is not None:
            try:
                boton_seleccionar.set_focus()
                time.sleep(0.5)
                boton_seleccionar.click()
            except Exception:
                try:
                    boton_seleccionar.click_input()
                except Exception:
                    try:
                        boton_seleccionar.invoke()
                        print("Botón Seleccionar activado con invoke().")
                    except Exception as exc:
                        print(f"No se pudo activar el botón con invoke(): {exc}")
        else:
            print("No se encontró el botón Seleccionar; intentando clic por coordenadas.")
            try:
                hacer_click_relativo(ventana_empresa, 0.06, 0.18)
                hacer_click_relativo(ventana_empresa, 0.06, 0.27)
                hacer_click_relativo(ventana_empresa, 0.06, 0.36)
            except Exception as exc:
                print(f"No se pudo hacer clic por coordenadas: {exc}")

    except Exception as exc:
        print(f"Error al seleccionar empresa y año: {exc}")

def seleccionar_tesoreria_explorador(app):
    ventana_principal = app.window(title_re=".*ContaNet ERP.*")
    time.sleep(1)

    if buscar_y_click_texto(ventana_principal, "TESORERIA"):
        print("Click en TESORERIA realizado.")
    else:
        hacer_click_relativo(ventana_principal, 0.06, 0.66)
        time.sleep(1)
        hacer_click_relativo(ventana_principal, 0.06, 0.76)

    # Esperar a que el panel lateral renderice y el Tree esté disponible
    time.sleep(1)
    ventana_tesoreria = app.window(title_re=".*ContaNet ERP.*")
    tree = esperar_tree_cargado(ventana_tesoreria, timeout=12, min_items=4)
    if tree is None:
        print("No se encontró Tree tras pulsar TESORERIA.")
    else:
        print("Tree detectado tras pulsar TESORERIA.")
    # Arbol de tesorería: intentar seleccionar "Explorador" por posición fija (índice 4) como primer intento
    tree = encontrar_treeview(ventana_tesoreria)
    if tree is not None:
        if click_tree_row_by_index(tree, 4, doble=False):
            return
        print("No se pudo seleccionar por posición fija; intentando selección por texto como fallback.")
    else:
        print("No se encontró el Tree para selección por posición fija.")

    # Intentar seleccionar por texto; si no está visible, hacer scroll en el Tree hasta encontrarlo
    intentos = 0
    encontrado = False
    max_intentos = 8
    while intentos < max_intentos and not encontrado:
        print(f"Intento de selección de 'Explorador', pasada {intentos+1}/{max_intentos}")
        if seleccionar_nodo_arbol(ventana_tesoreria, texto_objetivo="explorador", indice=0, doble=False):
            encontrado = True
            break

        # Si no se encontró, intentar desplazar la barra del Tree hacia abajo y reintentar
        try:
            tree = encontrar_treeview(ventana_tesoreria)
            if tree is None:
                break

            # Buscar la barra de desplazamiento vertical dentro del Tree
            vbar = None
            for s in tree.descendants(control_type="ScrollBar"):
                try:
                    if s.friendly_class_name().lower().startswith("scrollbar") or s.element_info.automation_id:
                        vbar = s
                        break
                except Exception:
                    continue

            if vbar is not None:
                # Intentar clic en 'Página Abajo' si existe
                pd = None
                for btn in vbar.descendants(control_type="Button"):
                    try:
                        title = str(getattr(btn, 'element_info').name or '').lower()
                        if 'página' in title or 'page' in title or 'abajo' in title:
                            pd = btn
                            break
                    except Exception:
                        continue
                if pd is not None:
                    try:
                        pd.click_input()
                    except Exception:
                        try:
                            pd.click()
                        except Exception:
                            print("No se pudo hacer click en Página Abajo del scrollbar.")
                else:
                    # Fallback: hacer scroll con la rueda del ratón sobre el centro del Tree
                    try:
                        rect = tree.rectangle()
                        hacer_click_en_coordenadas(rect.left + rect.width()//2, rect.top + rect.height()//2)
                        mouse.scroll(coords=(rect.left + rect.width()//2, rect.top + rect.height()//2), wheel_dist=-3)
                    except Exception:
                        print("No se pudo desplazar el Tree con la rueda.")
            else:
                hacer_click_relativo(ventana_tesoreria, 0.08, 0.30)

        except Exception as exc:
            print(f"Error al intentar desplazar el Tree: {exc}")

        time.sleep(0.7)
        intentos += 1

    if not encontrado:
        print("No se pudo seleccionar 'Explorador' tras intentos de texto; intentando clic directo en la fila esperada (índice 3).")
        tree = encontrar_treeview(ventana_tesoreria)
        if tree is not None:
            if click_tree_row_by_index(tree, 3, doble=False):
                return
        print("Fallo clic por fila; realizando clic por coordenadas del panel lateral como último recurso.")
        hacer_click_relativo(ventana_tesoreria, 0.08, 0.20, doble=True)
        time.sleep(0.5)
        hacer_click_relativo(ventana_tesoreria, 0.08, 0.25, doble=True)
        time.sleep(0.5)
        hacer_click_relativo(ventana_tesoreria, 0.08, 0.30, doble=True)


def probar_fechas(app, anio_actual, mes_actual, dia_desde=1):

    principal = app.window(title_re=".*ContaNet ERP.*")

    fec_desde = None
    fec_hasta = None
    btn_actualizar = None

    for ctrl in principal.descendants():

        try:
            texto = ctrl.window_text().strip()

            if texto == "Fec. Desde :":
                fec_desde = ctrl

            elif texto == "Fec. Hasta :":
                fec_hasta = ctrl

            elif texto == "Actualizar":
                btn_actualizar = ctrl

        except Exception:
            pass

    ultimo_dia = calendar.monthrange(int(anio_actual), int(mes_actual))[1]
    str_fec_desde = f"{int(dia_desde):02d}/{mes_actual:02d}/{anio_actual}"
    str_fec_hasta = f"{ultimo_dia:02d}/{mes_actual:02d}/{anio_actual}"

    # FECHA DESDE
    if fec_desde:

        rect = fec_desde.rectangle()

        # Clic dentro de la caja de texto, NO sobre la flecha
        x = rect.left + 40
        y = rect.top + rect.height() // 2

        print(f"CLICK DENTRO FECHA DESDE ({str_fec_desde}):", x, y)

        mouse.click(coords=(x, y))

        time.sleep(1)

        send_keys("^a")
        send_keys("{BACKSPACE}")

        time.sleep(0.5)

        send_keys(str_fec_desde)

    time.sleep(1)

    # FECHA HASTA
    if fec_hasta:

        rect = fec_hasta.rectangle()

        x = rect.left + 40
        y = rect.top + rect.height() // 2

        print(f"CLICK DENTRO FECHA HASTA ({str_fec_hasta}):", x, y)

        mouse.click(coords=(x, y))

        time.sleep(1)

        send_keys("^a")
        send_keys("{BACKSPACE}")

        time.sleep(0.5)

        send_keys(str_fec_hasta)

    time.sleep(1)

    # ACTUALIZAR
    if btn_actualizar:
        btn_actualizar.click_input()


def seleccionar_registro_y_modificar(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    btn_modificar = None

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()

            if texto == "Modificar Asiento":

                btn_modificar = ctrl
                break

        except Exception:
            pass

    if btn_modificar:
        btn_modificar.click_input()

        time.sleep(5)

    else:

        print("NO SE ENCONTRO MODIFICAR ASIENTO")


def copiar_todo_asiento(app):

    principal = app.window(
        title_re=".*ContaNet ERP.*"
    )

    boton_copiar_todo = None

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()

            if "Copiar" in texto and "Todo" in texto:

                boton_copiar_todo = ctrl
                break

        except Exception:
            pass

    if boton_copiar_todo:
        boton_copiar_todo.click_input()
        time.sleep(5)

    else:

        print(
            "NO SE ENCONTRO EL BOTON COPIAR TODO"
        )

def inspeccionar_ventana_actual(app):

    principal = app.window(
        title_re=".*ContaNet ERP.*"
    )

    print("\n========== CONTROLES ==========\n")

    for ctrl in principal.descendants():

        try:

            print(
                ctrl.friendly_class_name(),
                "=>",
                ctrl.window_text()
            )

        except:
            pass

    print("\n========== FIN CONTROLES ==========\n")


def filtrar_cuenta_10(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    filtro = None

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()

            #
            # NOS QUEDAMOS CON EL FILTRO DE LA GRILLA INFERIOR
            #
            if texto == "N° Cuenta fila del filtro":
                rect = ctrl.rectangle()
                if rect.top > 600:
                    filtro = ctrl
                    break

        except Exception:
            pass

    if filtro is None:
        return

    rect = filtro.rectangle()

    x = rect.left + 40
    y = rect.top + rect.height() // 2

    mouse.click(coords=(x, y))

    time.sleep(1)

    send_keys("^a")
    send_keys("{BACKSPACE}")
    send_keys("10.")

    time.sleep(3)


def abrir_librito_cuenta_10(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    fila = None

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()

            if texto == "N° Cuenta fila 1":

                rect = ctrl.rectangle()

                #
                # SOLO LA GRILLA INFERIOR
                #
                if rect.top > 650:
                    fila = ctrl
                    break

        except Exception:
            pass

    if fila is None:

        print("NO SE ENCONTRO FILA INFERIOR")

        return

    rect = fila.rectangle()

    #
    # X ROJA | LIBRITO | CUENTA
    #
    x = rect.left - 25
    y = rect.top + rect.height() // 2
    mouse.double_click(coords=(x, y))
    time.sleep(5)


def completar_agregar_cuenta(app, numero_operacion):

    principal = app.window(title_re=".*ContaNet ERP.*")

    checkbox = None
    edit_operacion = None
    boton_aceptar = None

    for ctrl in principal.descendants():

        try:

            clase = ctrl.friendly_class_name()
            rect = ctrl.rectangle()
            texto = ctrl.window_text().strip()

            #
            # Check Tipo Operación
            #
            if (
                clase == "CheckBox"
                and rect.top > 800
                and rect.top < 900
            ):
                checkbox = ctrl

            #
            # Campo Nro Ope./Cheque
            #
            elif (
                clase == "Edit"
                and rect.left > 740
                and rect.left < 950
                and rect.top > 870
                and rect.top < 930
            ):
                edit_operacion = ctrl

            #
            # Botón Aceptar
            #
            elif texto == "Aceptar":
                boton_aceptar = ctrl

        except Exception:
            pass

    #
    # Marcar Tipo Operación
    #
    if checkbox:

        try:
        
            estado = checkbox.get_toggle_state()
            if estado == 0:
                checkbox.click_input()
                time.sleep(1)
    
        except Exception as e:
            print("NO SE PUDO LEER EL ESTADO:", e)
            checkbox.click_input()
    
            time.sleep(1)

    #
    # Escribir número movimiento
    #
    if edit_operacion:

        edit_operacion.click_input()

        time.sleep(0.5)

        send_keys("^a")
        send_keys("{BACKSPACE}")
        send_keys(str(numero_operacion))

        time.sleep(1)

    #
    # Aceptar
    #
    if boton_aceptar:

        boton_aceptar.click_input()
        time.sleep(3)

        # Verificar mensaje de error de cliente
        error_cliente = False
        for ctrl in principal.descendants():
            try:
                texto_raw = ctrl.window_text()
                if texto_raw and "cliente" in texto_raw.lower() and "debe completar" in texto_raw.lower():
                    error_cliente = True
                    break
            except:
                pass
                
        if error_cliente:
            print("ERROR CLIENTE DETECTADO")
            # Cerrar el mensaje de error (Aceptar)
            for ctrl in principal.descendants():
                try:
                    if ctrl.window_text().strip() == "Aceptar":
                        rect = ctrl.rectangle()
                        if rect.top > 400 and rect.top < 700:
                            ctrl.click_input()
                            time.sleep(1)
                            break
                except:
                    pass
                    
            # Click Cancelar en Agregar Cuenta
            for ctrl in principal.descendants():
                try:
                    if ctrl.window_text().strip() == "Cancelar":
                        ctrl.click_input()
                        time.sleep(1)
                        break
                except:
                    pass
                    
            # Click Cancelar en Modificar Asiento
            for ctrl in principal.descendants():
                try:
                    if ctrl.window_text().strip() == "Cancelar":
                        ctrl.click_input()
                        time.sleep(1)
                        break
                except:
                    pass
            
            return "FALTA CLIENTE"

    return "OK"


def guardar_asiento(app):

    principal = app.window(title_re=".*ContaNet ERP.*")
    boton_guardar = None
    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()
            if texto == "Guardar":
                boton_guardar = ctrl
                break

        except Exception:
            pass

    if boton_guardar:
        boton_guardar.click_input()

        time.sleep(5)


def limpiar_filtro_cuenta(app):

    principal = app.window(title_re=".*ContaNet ERP.*")
    try:
        principal.set_focus()
        time.sleep(1)
    except Exception:
        pass

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()

            if texto == "N° Cuenta fila del filtro":

                rect = ctrl.rectangle()

                #
                # Solo filtro inferior
                #
                if rect.top > 600:

                    x = rect.left + 40
                    y = rect.top + rect.height() // 2

                    mouse.click(coords=(x, y))
                    time.sleep(1)

                    send_keys("^a")
                    send_keys("{BACKSPACE}")
                    time.sleep(2)

                    return

        except Exception:
            pass

def copiar_todo_nuevamente(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()

            if "Copiar" in texto and "Todo" in texto:
                ctrl.click_input()
                time.sleep(5)

                return

        except Exception:
            pass

    print("NO SE ENCONTRO BOTON COPIAR TODO")


def aceptar_mensaje_sistema(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    boton_aceptar = None

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()
            if texto == "Aceptar":
                rect = ctrl.rectangle()

                #
                # El botón del mensaje está aproximadamente
                # entre 500 y 900 de alto.
                #
                if rect.top > 450 and rect.top < 700:
                    boton_aceptar = ctrl
                    break

        except Exception:
            pass

    if boton_aceptar:
        boton_aceptar.click_input()

        for _ in range(10):
            if cerrar_mensaje_tipo_cambio(app):
                break
        time.sleep(1)

        print("MENSAJE CERRADO")
        time.sleep(2)
    else:
        print("NO SE ENCONTRO EL MENSAJE")


def seleccionar_fila_reg_ctb(app, indice):
    principal = app.window(title_re=".*ContaNet ERP.*")
    for ctrl in principal.descendants():
        try:
            texto = ctrl.window_text().strip()
            if texto == "Reg. Ctb.":
                rect = ctrl.rectangle()
                x = rect.left + rect.width() // 2
                y_top = rect.bottom + 55

                try:
                    principal.set_focus()
                    time.sleep(0.5)
                except:
                    pass

                # Calibración exacta basada en la pantalla del usuario
                max_visible = 14
                
                if indice < max_visible:
                    y = y_top + (indice * 34)
                    mouse.click(coords=(x, y))
                    mouse.move(coords=(x, y))
                    time.sleep(1)
                    mouse.double_click(coords=(x, y))
                    time.sleep(1)
                    return y
                else:
                    y_first = y_top  # Primera fila visible
                    y_last_visible = y_top + ((max_visible - 1) * 34)
                                        
                    # 1. Hacer click en la primera fila para asegurar foco en la grilla
                    mouse.click(coords=(x, y_first))
                    time.sleep(0.5)
                    
                    # 2. Presionar Ctrl+Home para ir al inicio absoluto de la grilla
                    send_keys("^{HOME}")
                    time.sleep(0.5)
                    
                    # 3. Hacer click en la primera fila de nuevo para confirmar posición
                    mouse.click(coords=(x, y_first))
                    time.sleep(0.3)
                    
                    # 4. Bajar con DOWN desde la primera fila hasta la fila destino
                    for i in range(indice):
                        send_keys("{DOWN}")
                        time.sleep(0.1)
                    
                    time.sleep(0.5)
                    
                    # 5. La fila destino queda en la última posición visible tras el scroll
                    mouse.click(coords=(x, y_last_visible))
                    mouse.move(coords=(x, y_last_visible))
                    time.sleep(1)
                    mouse.double_click(coords=(x, y_last_visible))
                    time.sleep(1)
                    return y_last_visible
                
        except:
            pass
    return None

def seleccionar_siguiente_fila(app):
    # Ya no es necesario porque seleccionar_fila_reg_ctb navega desde el inicio
    pass



def obtener_asiento_contable(app, y):
    principal = app.window(title_re=".*ContaNet ERP.*")
    x = None
    for ctrl in principal.descendants():
        try:
            texto = ctrl.window_text().strip()
            if texto == "Reg. Ctb.":
                rect = ctrl.rectangle()
                x = rect.left + rect.width() // 2
                break
        except Exception:
            pass

    if x is None:
        print("NO SE ENCONTRO LA COLUMNA Reg. Ctb.")
        return None

    # Click en la celda de la columna "Reg. Ctb." para obtener el número de asiento
    mouse.click(coords=(x, y))
    time.sleep(0.3)
    send_keys("^c")
    time.sleep(0.3)

    try:
        win32clipboard.OpenClipboard()
        asiento = (
                win32clipboard
                .GetClipboardData()
                .strip()
            )
        win32clipboard.CloseClipboard()

        return asiento

    except Exception:
        return None
    

def obtener_cuenta_contable_grilla(app, y):
    principal = app.window(title_re=".*ContaNet ERP.*")
    x = None
    for ctrl in principal.descendants():
        try:
            texto = ctrl.window_text().strip()
            if texto == "Nro. Cta.":
                rect = ctrl.rectangle()
                x = rect.left + rect.width() // 2
                break
        except Exception:
            pass

    if x is None:
        return None

    # Click directo en la celda de la columna "Nro. Cta." en la fila actual
    mouse.click(coords=(x, y))
    time.sleep(0.3)

    send_keys("^c")
    time.sleep(0.3)

    try:
        win32clipboard.OpenClipboard()
        cuenta = win32clipboard.GetClipboardData().strip()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
    except Exception:
        cuenta = None

    return cuenta


def buscar_numero_operacion(
    hoja,
    asiento,
    anio
):
    """Busca el número de operación en el Excel.
    
    Retorna una tupla (numero_operacion, error):
    - (str, None) si se encontró y es un número válido
    - (None, "NO_ENCONTRADO") si no se encontró el asiento en el Excel
    - (None, "NO_ES_NUMERO") si el valor encontrado no es numérico
    """
    try:

        df = excel_data[hoja]

        df["Asiento Contable"] = (
            df["Asiento Contable"]
            .astype(str)
            .str.strip()
        )

        df["Año"] = (
            pd.to_numeric(
                df["Año"],
                errors="coerce"
            )
        )

        resultado = df[
            (df["Asiento Contable"] == asiento)
            &
            (df["Año"] == int(anio))
        ]

        if resultado.empty:
            return None, "NO_ENCONTRADO"

        valor_operacion = resultado.iloc[0]["Número de Operación"]
        valor_str = str(valor_operacion).strip()

        # Validar que el valor sea un número entero
        try:
            int(valor_str)
        except (ValueError, TypeError):
            print(
                f"El número de operación '{valor_str}' para el asiento {asiento} "
                f"no es un valor numérico válido."
            )
            return None, "NO_ES_NUMERO"

        return valor_str, None

    except Exception as e:
        print(e)

        print(
            f"Asiento {asiento} no encontrado para el año {anio}"
        )

        return None, "NO_ENCONTRADO"


def procesar_fila(app, numero_operacion):

    seleccionar_registro_y_modificar(app)
    time.sleep(3)

    tipo_error = tipo_mensaje_error(app)
    if tipo_error:
        cerrar_mensaje_no_modificable(app)

        seleccionar_siguiente_fila(app)

        print(f"SE OMITE ESTA FILA: {tipo_error}")
        return tipo_error


    copiar_todo_asiento(app)
    time.sleep(3)

    filtrar_cuenta_10(app)
    time.sleep(2)

    abrir_librito_cuenta_10(app)
    time.sleep(2)

    # Si aparece el aviso de "No se ha encontrado el tipo de cambio", solicitamos reinicio
    if cerrar_mensaje_tipo_cambio(app) == "REINICIAR":
        print("Mensaje de tipo de cambio detectado. Se retorna REINICIAR.")
        return "REINICIAR"

    resultado_agregar = completar_agregar_cuenta(app, numero_operacion)
    time.sleep(2)
    if resultado_agregar == "FALTA CLIENTE":
        seleccionar_siguiente_fila(app)
        return "FALTA CLIENTE"

    limpiar_filtro_cuenta(app)
    time.sleep(2)

    copiar_todo_nuevamente(app)
    time.sleep(3)

    guardar_asiento(app)
    time.sleep(2)

    # Verificar si al guardar aparece mensaje de error (ej. Periodo cerrado o No se puede...)
    tipo_err_guardar = tipo_mensaje_error(app)
    if tipo_err_guardar:
        print(f"ERROR DETECTADO AL GUARDAR: {tipo_err_guardar}")
        cerrar_mensaje_no_modificable(app)
        time.sleep(1)

        # Click en Cancelar en la ventana Modificar Asiento para salir limpia/cerradamente
        principal = app.window(title_re=".*ContaNet ERP.*")
        for ctrl in principal.descendants():
            try:
                if ctrl.window_text().strip() == "Cancelar":
                    ctrl.click_input()
                    time.sleep(1)
                    break
            except:
                pass

        seleccionar_siguiente_fila(app)
        return tipo_err_guardar

    aceptar_mensaje_sistema(app)
    time.sleep(2)
    seleccionar_siguiente_fila(app)
    return "EXITO"

def tipo_mensaje_error(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    for ctrl in principal.descendants():
        try:
            texto_raw = ctrl.window_text()

            if not texto_raw:
                continue

            texto = texto_raw.strip().lower()

            if "no se puede modificar un voucher" in texto:
                return "VOUCHER NO MODIFICABLE"

            if "asignado a caja" in texto:
                return "ASIGNADO A CAJA"

            if "no puede culminar el proceso" in texto or "periodo cerrado" in texto or "período cerrado" in texto:
                return "PERIODO CERRADO"

        except:
            pass

    return None


def existe_mensaje_no_modificable(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    for ctrl in principal.descendants():
        try:
            texto = ctrl.window_text().strip()
            if "No se puede Modificar un Voucher" in texto:
                print("VOUCHER NO MODIFICABLE")
                return True
        except:
            pass
    return False


def cerrar_mensaje_no_modificable(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    for ctrl in principal.descendants():
        try:
            if ctrl.window_text().strip() == "Aceptar":
                rect = ctrl.rectangle()
                #
                # mensaje emergente
                #
                if rect.top > 400 and rect.top < 700:
                    print("CERRANDO MENSAJE")
                    ctrl.click_input()
                    time.sleep(2)
                    return True
        except:
            pass
    return False


def tiene_movimientos(app, timeout=8):
    """Verifica si existen movimientos en la grilla tras presionar Actualizar.
    Si tras 'timeout' segundos sigue indicando activamente '0 filas' (o '0 fila'), retorna False.
    En cualquier otro caso (si detecta filas o si hay duda/retraso), retorna True.
    """
    principal = app.window(title_re=".*ContaNet ERP.*")
    inicio = time.time()

    # Primero intentamos detectar si explícitamente hay filas cargadas (> 0 filas)
    while time.time() - inicio < timeout:
        for ctrl in principal.descendants():
            try:
                texto = ctrl.window_text().strip()

                if "Resultado :" in texto:
                    texto_lower = texto.lower()

                    # Si NO tiene '0 fila' ni '0 filas', pero sí menciona 'fila', definitivamente hay movimientos
                    if ("fila" in texto_lower) and ("0 fila" not in texto_lower) and ("0 filas" not in texto_lower):
                        return True
            except Exception:
                pass

        time.sleep(1)

    # Verificación final: ÚNICAMENTE si la etiqueta dice explícitamente "0 fila" o "0 filas" al terminar el tiempo
    for ctrl in principal.descendants():
        try:
            texto = ctrl.window_text().strip()
            if "Resultado :" in texto:
                texto_lower = texto.lower()
                if "0 fila" in texto_lower or "0 filas" in texto_lower:
                    return False
        except Exception:
            pass

    # Si por cualquier razón la interfaz no dio lectura clara, asumimos que SÍ hay movimientos
    return True


def procesar_empresa_anio(
    ruc,
    anio_actual
):

    app = iniciar_aplicacion()

    ventana_login = obtener_ventana(app)

    ventana_login = app.window(
        title_re=".*Login.*"
    )

    iniciar_sesion(
        ventana_login
    )

    time.sleep(2)

    ventana_empresa = app.window(
        title_re=".*Seleccionar Empresa.*"
    )

    seleccionar_empresa_y_anio(
        ventana_empresa,
        ruc,
        anio_actual
    )

    time.sleep(2)

    seleccionar_tesoreria_explorador(app)

    time.sleep(5)

    probar_fechas(
        app,
        anio_actual
    )

    time.sleep(5)

    if not tiene_movimientos(app):
        cerrar_aplicacion(app)
        return False

    print(
        f"MOVIMIENTOS EN {anio_actual}"
    )

    return True


def tipo_mensaje_error(app):
    principal = app.window(title_re=".*ContaNet ERP.*")

    for ctrl in principal.descendants():
        try:
            texto_raw = ctrl.window_text()
            if not texto_raw:
                continue

            texto = texto_raw.strip().lower()

            if "no se puede modificar un voucher" in texto:
                return "VOUCHER NO MODIFICABLE"

            if "asignado a caja" in texto:
                return "ASIGNADO A CAJA"

            if "periodo" in texto and "cerrado" in texto:
                return "PERIODO CERRADO"

        except:
            pass

    return None


def cerrar_mensaje_tipo_cambio(app):
    try:
        for ventana in app.windows():
            try:
                for ctrl in ventana.descendants():
                    texto = ctrl.window_text().strip()

                    if "No se ha encontrado el tipo de cambio" in texto:
                        for btn in ventana.descendants(control_type="Button"):
                            if btn.window_text().strip() == "Aceptar":
                                btn.click_input()
                                time.sleep(2)
                                return "REINICIAR"
            except Exception:
                pass

    except Exception as e:
        print("Error detectando mensaje:", e)

    return False

def main():
    anios = [2026, 2025, 2024, 2023, 2022, 2021]

    for ruc, nombre_empresa in empresas.items():
        hoja_excel = nombre_empresa
        if hoja_excel not in excel_data:
            print(f"RUC {ruc} ({nombre_empresa}) no tiene pestaña con su nombre en el Excel, se omite.")
            continue

        print(f"\n{'='*60}")
        print(f"EMPRESA: {nombre_empresa} (RUC: {ruc})")
        print(f"{'='*60}")

        for anio_actual in anios:
            print(f"\n--- Procesando {nombre_empresa} | Año {anio_actual} ---")

            for mes_actual in range(1, 13):
                mes_inicio = MESES_INICIO.get(ruc, {}).get(anio_actual, 1)
                if mes_actual < mes_inicio:
                    print(f"Saltando mes {mes_actual:02d}/{anio_actual} para {nombre_empresa} por configuración (inicia en mes {mes_inicio:02d}).")
                    continue

                print(f"\n--- Procesando Mes {mes_actual:02d}/{anio_actual} ---")

                if es_mes_sin_movimientos(ruc, anio_actual, mes_actual):
                    print(f"Mes {mes_actual:02d}/{anio_actual} ya fue revisado y no tiene movimientos, saltando.")
                    continue

                app = iniciar_aplicacion()

                ventana_login = obtener_ventana(app)
                print("Ventana detectada:", ventana_login.window_text())
                time.sleep(1)

                ventana_login = app.window(title_re=".*Login.*")
                iniciar_sesion(ventana_login)
                time.sleep(2)

                try:
                    ventana_empresa_win = app.window(title_re=".*Seleccionar Empresa.*")
                except Exception as exc:
                    print(f"No se encontró la ventana de selección de empresa: {exc}")
                    try:
                        cerrar_aplicacion(app)
                    except Exception:
                        pass
                    continue

                seleccionar_empresa_y_anio(
                    ventana_empresa_win,
                    ruc,
                    anio_actual
                )
                time.sleep(2)

                seleccionar_tesoreria_explorador(app)
                time.sleep(5)

                dia_desde = FECHAS_INICIO.get(ruc, {}).get(anio_actual, {}).get(mes_actual, 1)

                probar_fechas(app, anio_actual, mes_actual, dia_desde=dia_desde)
                time.sleep(7)

                # Si es un rango de días especial (como 28/01 a 31/01 en STN), omitimos el descarte para garantizar que procese
                if dia_desde == 1 and not tiene_movimientos(app):
                    print(f"SIN MOVIMIENTOS para {nombre_empresa} en {mes_actual:02d}/{anio_actual}, pasando al siguiente mes.")
                    registrar_reporte(ruc, nombre_empresa, anio_actual, "N/A", "SIN MOVIMIENTOS", f"No hay registros en el mes {mes_actual:02d}", numero_operacion="")
                    marcar_mes_sin_movimientos(ruc, anio_actual, mes_actual)
                    cerrar_aplicacion(app)
                    time.sleep(3)
                    continue

                send_keys("^c")
                time.sleep(1)

                print(f"TRABAJANDO CON {nombre_empresa} AÑO {anio_actual} MES {mes_actual:02d}")

                ultimo_asiento = None
                ultima_cuenta = None
                ultimo_asiento_procesado = None
                repetidos = 0
                filas_estancadas = 0
                fila = obtener_ultima_fila(ruc, anio_actual, mes_actual)
                if fila > 0:
                    print(f"REANUDANDO DESDE FILA {fila + 1} (progreso guardado)")
                while True:
                    # ── Reinicio programado ──────────────────────────────────
                    if debe_reiniciar_por_horario():
                        app = reiniciar_app_programado(
                            app, ruc, anio_actual, mes_actual, dia_desde
                        )
                    # ────────────────────────────────────────────────────────

                    y_coord = seleccionar_fila_reg_ctb(app, fila)
                    if not y_coord:
                        print(f"No se pudo seleccionar fila {fila}, terminando este mes.")
                        break

                    time.sleep(0.5)

                    cuenta_raw = obtener_cuenta_contable_grilla(app, y_coord)
                    
                    cuenta = None
                    is_10_4 = False
                    if cuenta_raw:
                        palabras = str(cuenta_raw).split()
                        for p in palabras:
                            if p.startswith("10."):
                                cuenta = p
                                if p.startswith("10.4."):
                                    is_10_4 = True
                                break
                        if not cuenta:
                            cuenta = str(cuenta_raw)[:20]

                    asiento = obtener_asiento_contable(app, y_coord)

                    if cuenta == ultima_cuenta and asiento is not None and asiento == ultimo_asiento:
                        filas_estancadas += 1
                        if filas_estancadas >= 3:
                            print("Se alcanzó el final de la grilla (la fila no avanza).")
                            break
                    else:
                        filas_estancadas = 0

                    ultima_cuenta = cuenta
                    if asiento is not None:
                        ultimo_asiento = asiento

                    if not is_10_4:
                        print(f"FILA {fila + 1} NO PROCESADO: La cuenta {cuenta} no empieza con 10.4.xxx")
                        fila += 1
                        guardar_ultima_fila(ruc, anio_actual, mes_actual, fila)
                        continue
                    
                    print(f"EVALUANDO: Cuenta {cuenta} | Asiento {asiento}")
                    
                    if asiento == ultimo_asiento_procesado:
                        repetidos += 1
                        if repetidos >= 3:
                            print("Se alcanzó el límite (mismo asiento 3 veces consecutivas para cuentas 10.4).")
                            break
                        else:
                            print(f"Asiento repetido {repetidos} vez/veces, saltando para ver si la grilla avanza...")
                    else:
                        repetidos = 0
                    
                    ultimo_asiento_procesado = asiento

                    if es_asiento_procesado(ruc, anio_actual, mes_actual, asiento):
                        print(f"FILA {fila + 1} ASIENTO {asiento} ya procesado anteriormente, saltando.")
                        fila += 1
                        guardar_ultima_fila(ruc, anio_actual, mes_actual, fila)
                        continue

                    numero_operacion, error_operacion = buscar_numero_operacion(
                        hoja_excel,
                        asiento,
                        anio_actual
                    )
                    print("NUMERO OPERACION:", numero_operacion)

                    if numero_operacion is None:
                        if error_operacion == "NO_ES_NUMERO":
                            print(f"NÚMERO DE OPERACIÓN NO VÁLIDO (no es numérico) para asiento {asiento}, saltando fila.")
                            registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "NO MODIFICADO", "El número de operación no es un valor numérico válido", numero_operacion="")
                        else:
                            print("NO ENCONTRADO EN EXCEL, saltando fila.")
                            registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "NO MODIFICADO", "No se encontró el número de operación en el Excel", numero_operacion="")
                        registrar_asiento_procesado(ruc, anio_actual, mes_actual, asiento)
                        fila += 1
                        guardar_ultima_fila(ruc, anio_actual, mes_actual, fila)
                        continue

                    resultado = procesar_fila(app, numero_operacion)

                    if resultado == "PERIODO CERRADO":
                        registrar_reporte(
                            ruc,
                            nombre_empresa,
                            anio_actual,
                            asiento,
                            "NO MODIFICADO",
                            "Periodo cerrado",
                            numero_operacion=numero_operacion
                        )
                    
                    if resultado == "REINICIAR":
                        print(f"REINICIANDO APLICACIÓN tras mensaje de tipo de cambio en asiento {asiento}")
                        registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "ACTUALIZADO", "Se detectó tipo de cambio (requirió reinicio)", numero_operacion=numero_operacion)
                        registrar_asiento_procesado(ruc, anio_actual, mes_actual, asiento)
                        fila += 1
                        guardar_ultima_fila(ruc, anio_actual, mes_actual, fila)
                        
                        cerrar_aplicacion(app)
                        time.sleep(3)
                        
                        app = iniciar_aplicacion()
                        ventana_login = obtener_ventana(app)
                        ventana_login = app.window(title_re=".*Login.*")
                        iniciar_sesion(ventana_login)
                        time.sleep(2)
                        ventana_empresa_win = app.window(title_re=".*Seleccionar Empresa.*")
                        seleccionar_empresa_y_anio(ventana_empresa_win, ruc, anio_actual)
                        time.sleep(2)
                        seleccionar_tesoreria_explorador(app)
                        time.sleep(5)
                        probar_fechas(app, anio_actual, mes_actual, dia_desde=dia_desde)
                        time.sleep(7)
                        send_keys("^c")
                        time.sleep(1)
                        continue

                    if resultado == "VOUCHER NO MODIFICABLE":
                        registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "NO MODIFICADO", "Voucher no modificable", numero_operacion=numero_operacion)
                    elif resultado == "ASIGNADO A CAJA":
                        registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "NO MODIFICADO", "El asiento contable ha sido asignado a caja", numero_operacion=numero_operacion)
                    else:
                        registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "ACTUALIZADO", "Se completó la actualización", numero_operacion=numero_operacion)
                        
                    registrar_asiento_procesado(ruc, anio_actual, mes_actual, asiento)
                    time.sleep(3)
                    fila += 1
                    guardar_ultima_fila(ruc, anio_actual, mes_actual, fila)

                print(f"Cerrando aplicación para {nombre_empresa} año {anio_actual} mes {mes_actual:02d}")
                cerrar_aplicacion(app)
                time.sleep(3)

    print("\n" + "="*60)
    print("PROCESO COMPLETADO PARA TODAS LAS EMPRESAS Y AÑOS")
    print("="*60)
    generar_reporte_metricas()


def cargar_data_excel():
    pass


def open_contanet():
    pass


def ingresar_datos():
    pass


if __name__ == "__main__":
    ESPERA_ENTRE_REINICIOS = 30  # segundos a esperar antes de reintentar tras un fallo

    intento = 0
    while True:
        intento += 1
        print(f"\n{'='*60}")
        print(f"[WATCHDOG] Iniciando ejecución (intento #{intento})")
        print(f"[WATCHDOG] Hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*60}\n")

        try:
            main()
            # Si main() terminó sin excepción, el proceso está completo
            print("[WATCHDOG] main() finalizó correctamente. Cerrando watchdog.")
            break

        except KeyboardInterrupt:
            print("\n[WATCHDOG] Interrupción manual (Ctrl+C). Cerrando.")
            forzar_cierre_procesos()
            break

        except Exception as e:
            print(f"\n[WATCHDOG] ERROR INESPERADO en intento #{intento}: {e}")
            print(f"[WATCHDOG] Limpiando procesos huérfanos...")
            forzar_cierre_procesos()
            print(f"[WATCHDOG] Reintentando en {ESPERA_ENTRE_REINICIOS} segundos...")
            print(f"[WATCHDOG] (El progreso guardado en progreso.json se respetará)")
            time.sleep(ESPERA_ENTRE_REINICIOS)
