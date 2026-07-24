from pywinauto import Application, mouse
from pywinauto.keyboard import send_keys
import win32clipboard
import time
import pandas as pd
import json
import os
import csv

ARCHIVO_PROGRESO = "progreso.json"
ARCHIVO_REPORTE = "reporte_asientos.csv"

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

def registrar_asiento_procesado(ruc, anio, asiento):
    progreso = cargar_progreso()
    if ruc not in progreso:
        progreso[ruc] = {}
    
    anio_str = str(anio)
    if anio_str not in progreso[ruc]:
        progreso[ruc][anio_str] = {"asientos": [], "ultima_fila": 0}
    
    # Migración: si el formato antiguo era una lista, convertirlo
    if isinstance(progreso[ruc][anio_str], list):
        progreso[ruc][anio_str] = {"asientos": progreso[ruc][anio_str], "ultima_fila": 0}
        
    if asiento not in progreso[ruc][anio_str]["asientos"]:
        progreso[ruc][anio_str]["asientos"].append(asiento)
        guardar_progreso(progreso)

def es_asiento_procesado(ruc, anio, asiento):
    progreso = cargar_progreso()
    anio_str = str(anio)
    if ruc in progreso and anio_str in progreso[ruc]:
        dato = progreso[ruc][anio_str]
        # Migración: formato antiguo era lista
        if isinstance(dato, list):
            return asiento in dato
        return asiento in dato.get("asientos", [])
    return False

def guardar_ultima_fila(ruc, anio, fila):
    progreso = cargar_progreso()
    if ruc not in progreso:
        progreso[ruc] = {}
    anio_str = str(anio)
    if anio_str not in progreso[ruc]:
        progreso[ruc][anio_str] = {"asientos": [], "ultima_fila": 0}
    if isinstance(progreso[ruc][anio_str], list):
        progreso[ruc][anio_str] = {"asientos": progreso[ruc][anio_str], "ultima_fila": 0}
    progreso[ruc][anio_str]["ultima_fila"] = fila
    guardar_progreso(progreso)

def obtener_ultima_fila(ruc, anio):
    progreso = cargar_progreso()
    anio_str = str(anio)
    if ruc in progreso and anio_str in progreso[ruc]:
        dato = progreso[ruc][anio_str]
        if isinstance(dato, dict):
            return dato.get("ultima_fila", 0)
    return 0

ruta_exe = (
    r"C:\Ejecutable ContaNet ERP 3.0.7.44 - SOUTHERN"
    r"\ContaNet.Aplicacion.exe"
)

directorio_trabajo = (
    r"C:\Ejecutable ContaNet ERP 3.0.7.44 - SOUTHERN"
)

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
    "20490242407": "INFOSUR",
    "20504334041": "ITS",
    "20376729126": "STN",
    "20506883301": "CMT",
    "20600692781": "DIONISO",
    "20514016624": "DYNAMITEX",
    "20606955724": "PERU COMMERCE",
    "20494530865": "DINSURA",
    "20542813238": "IÑAPARI",
    "20609254778": "TECA"
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

                    return True

            except:
                pass

        print(
            "NO SE ENCONTRO BOTON SI"
        )

        return False

    except Exception as e:

        print(
            "Error cerrando aplicación:",
            e
        )

        return False


def iniciar_sesion(ventana):
    ventana.child_window(auto_id="txtUsuario", control_type="Edit").set_text("COSTOS")
    ventana.child_window(auto_id="txtContrasena", control_type="Edit").set_text("1234")

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

    print(f"Trees encontrados: {len(trees)}")

    for t in trees:
        try:
            print(
                "TREE:",
                t.element_info.name,
                t.element_info.automation_id,
                t.rectangle()
            )
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

    # Mostrar los textos detectados para depuración
    print("Candidatos detectados en árbol:", [t for t, _ in candidates])

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
                    print(f"Elemento seleccionado por texto: {texto_ctrl}")
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

    # timeout waiting for tree to populate
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
            # selecting target row
            
            print("Nombre:", target.window_text())
            print("Element name:", target.element_info.name)
            print("AutomationId:", target.element_info.automation_id)
            print("Rect:", target.rectangle())

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

                # clicked target row
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
        # click estimated on tree row
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
    print(f"Fallback: clic relativo en ({x}, {y})")
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

        # Empresa selection UI interaction

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

        # Buttons will be located programmatically

        boton_seleccionar = encontrar_boton_seleccionar(ventana_empresa)
        if boton_seleccionar is not None:
            try:
                boton_seleccionar.set_focus()
                time.sleep(0.5)
                boton_seleccionar.click()
                print("Botón Seleccionar presionado con click().")
            except Exception:
                try:
                    boton_seleccionar.click_input()
                    print("Botón Seleccionar presionado con click_input().")
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
                print("Clic por coordenadas intentado en posiciones de boton.")
            except Exception as exc:
                print(f"No se pudo hacer clic por coordenadas: {exc}")

    except Exception as exc:
        print(f"Error al seleccionar empresa y año: {exc}")

def seleccionar_tesoreria_explorador(app):
    ventana_principal = app.window(title_re=".*ContaNet ERP.*")
    print("Ventana principal detectada:", ventana_principal.window_text())
    print("Rect principal:", ventana_principal.rectangle())
    time.sleep(1)

    if buscar_y_click_texto(ventana_principal, "TESORERIA"):
        print("Click en TESORERIA realizado.")
    else:
        print("No se encontró TESORERIA en la ventana principal; intento por coordenadas.")
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
    # Tree and related controls will be inspected programmatically when needed

    tree = encontrar_treeview(ventana_tesoreria)
    if tree is not None:
        print("Intentando seleccionar por posición fija: quinta fila (índice 4).")
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
                        print("Scroll: página abajo pulsada.")
                    except Exception:
                        try:
                            pd.click()
                            print("Scroll: página abajo pulsada (click()).")
                        except Exception:
                            print("No se pudo hacer click en Página Abajo del scrollbar.")
                else:
                    # Fallback: hacer scroll con la rueda del ratón sobre el centro del Tree
                    try:
                        rect = tree.rectangle()
                        hacer_click_en_coordenadas(rect.left + rect.width()//2, rect.top + rect.height()//2)
                        mouse.scroll(coords=(rect.left + rect.width()//2, rect.top + rect.height()//2), wheel_dist=-3)
                        print("Scroll: rueda enviada al Tree.")
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


def probar_fechas(app, anio_actual):

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

    # FECHA DESDE
    if fec_desde:

        rect = fec_desde.rectangle()

        # Clic dentro de la caja de texto, NO sobre la flecha
        x = rect.left + 40
        y = rect.top + rect.height() // 2

        print("CLICK DENTRO FECHA DESDE:", x, y)

        mouse.click(coords=(x, y))

        time.sleep(1)

        send_keys("^a")
        send_keys("{BACKSPACE}")

        time.sleep(0.5)

        send_keys(f"01/01/{anio_actual}")

    time.sleep(1)

    # FECHA HASTA
    if fec_hasta:

        rect = fec_hasta.rectangle()

        x = rect.left + 40
        y = rect.top + rect.height() // 2

        print("CLICK DENTRO FECHA HASTA:", x, y)

        mouse.click(coords=(x, y))

        time.sleep(1)

        send_keys("^a")
        send_keys("{BACKSPACE}")

        time.sleep(0.5)

        send_keys(f"31/12/{anio_actual}")

    time.sleep(1)

    # ACTUALIZAR
    if btn_actualizar:

        print("CLICK ACTUALIZAR")

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

        print("CLICK MODIFICAR ASIENTO")

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

                    print(
                        "FILTRO INFERIOR:",
                        rect
                    )

                    break

        except Exception:
            pass

    if filtro is None:

        print("NO SE ENCONTRO FILTRO INFERIOR")

        return

    rect = filtro.rectangle()

    x = rect.left + 40
    y = rect.top + rect.height() // 2

    print("CLICK FILTRO:", x, y)

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

                    print("FILA INFERIOR:", rect)

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

    print("CLICK LIBRITO:", x, y)

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

    print("CHECK:", checkbox)
    print("EDIT:", edit_operacion)
    print("ACEPTAR:", boton_aceptar)

    #
    # Marcar Tipo Operación
    #
    if checkbox:

        try:
        
            estado = checkbox.get_toggle_state()
    
            print(
                "ESTADO CHECK:",
                estado
            )
    
            if estado == 0:
            
                print(
                    "MARCANDO CHECK"
                )
    
                checkbox.click_input()
    
                time.sleep(1)
    
            else:
            
                print(
                    "CHECK YA MARCADO"
                )
    
        except Exception as e:
        
            print(
                "NO SE PUDO LEER EL ESTADO:",
                e
            )
    
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

        print("ACEPTAR PRESIONADO")

        time.sleep(3)


def guardar_asiento(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    boton_guardar = None

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()

            if texto == "Guardar":

                boton_guardar = ctrl

                print(
                    "BOTON GUARDAR:",
                    ctrl.rectangle()
                )

                break

        except Exception:
            pass

    if boton_guardar:

        print("CLICK GUARDAR")

        boton_guardar.click_input()

        time.sleep(5)

    else:

        print("NO SE ENCONTRO BOTON GUARDAR")


def limpiar_filtro_cuenta(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

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

                    print("LIMPIANDO FILTRO:", x, y)

                    mouse.click(coords=(x, y))

                    time.sleep(1)

                    send_keys("^a")
                    send_keys("{BACKSPACE}")

                    time.sleep(2)

                    return

        except Exception:
            pass

    print("NO SE ENCONTRO FILTRO INFERIOR")


def copiar_todo_nuevamente(app):

    principal = app.window(title_re=".*ContaNet ERP.*")

    for ctrl in principal.descendants():

        try:

            texto = ctrl.window_text().strip()

            if "Copiar" in texto and "Todo" in texto:

                print(
                    "BOTON COPIAR TODO:",
                    ctrl.rectangle()
                )

                ctrl.click_input()

                print("CLICK COPIAR TODO NUEVAMENTE")

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
                    print(
                        "ACEPTAR MENSAJE:",
                        rect
                    )
                    break

        except Exception:
            pass

    if boton_aceptar:
        boton_aceptar.click_input()
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
                    print(f"SELECCIONANDO FILA {indice + 1} (Coordenada directa: {x}, {y})")
                    mouse.click(coords=(x, y))
                    mouse.move(coords=(x, y))
                    time.sleep(1)
                    mouse.double_click(coords=(x, y))
                    time.sleep(1)
                    return y
                else:
                    y_first = y_top  # Primera fila visible
                    y_last_visible = y_top + ((max_visible - 1) * 34)
                    
                    print(f"SELECCIONANDO FILA {indice + 1} (Scrolleando desde el inicio, {indice - max_visible + 1} veces)")
                    
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



def obtener_asiento_contable():
    # Clic para quitar foco de la grilla
    mouse.click(coords=(650, 171))
    time.sleep(0.3)
    # Clic para enfocar el textbox
    mouse.click(coords=(650, 171))
    time.sleep(0.3)
    # Doble clic para seleccionar el texto
    mouse.double_click(coords=(650, 171))

    time.sleep(0.5)

    send_keys("^a")
    time.sleep(0.2)

    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
    except Exception:
        pass

    send_keys("^c")
    time.sleep(0.5)

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
        print("NO SE ENCONTRO LA COLUMNA Nro. Cta.")
        return None

    # Click en la celda de la columna "Nro. Cta." en la fila actual
    mouse.click(coords=(x, y))
    time.sleep(0.5)

    send_keys("^c")
    time.sleep(0.5)

    try:
        win32clipboard.OpenClipboard()
        cuenta = win32clipboard.GetClipboardData().strip()
        win32clipboard.EmptyClipboard() # Limpiar portapapeles para que no se filtre al asiento
        win32clipboard.CloseClipboard()
    except Exception:
        cuenta = None

    # Hacemos clic en la columna "Reg. Ctb." de esta misma fila para asegurar que salimos del modo edición de la celda
    for ctrl in principal.descendants():
        try:
            texto = ctrl.window_text().strip()
            if texto == "Reg. Ctb.":
                rect = ctrl.rectangle()
                x_reg = rect.left + rect.width() // 2
                mouse.click(coords=(x_reg, y))
                time.sleep(0.3)
                break
        except Exception:
            pass

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

    completar_agregar_cuenta(app, numero_operacion)
    time.sleep(2)

    limpiar_filtro_cuenta(app)
    time.sleep(2)

    copiar_todo_nuevamente(app)
    time.sleep(3)

    guardar_asiento(app)
    time.sleep(2)

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

            if "no puede culminar el proceso" in texto:
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


def tiene_movimientos(app):

    principal = app.window(
        title_re=".*ContaNet ERP.*"
    )

    for ctrl in principal.descendants():
        try:
            texto = ctrl.window_text().strip()

            if "Resultado :" in texto:
                print(
                        "RESULTADO DETECTADO:",
                        texto
                    )


                if "0 fila" in texto:
                    return False

                return True

        except:
            pass

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

        print(
            f"SIN MOVIMIENTOS EN {anio_actual}"
        )

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


def main():
    anios = [2026, 2025, 2024, 2023, 2022, 2021]

    for ruc, nombre_empresa in empresas.items():
        # Validar que el nombre de la empresa coincida con el nombre de la pestaña de la página
        hoja_excel = nombre_empresa
        if hoja_excel not in excel_data:
            print(f"RUC {ruc} ({nombre_empresa}) no tiene pestaña con su nombre en el Excel, se omite.")
            continue

        print(f"\n{'='*60}")
        print(f"EMPRESA: {nombre_empresa} (RUC: {ruc})")
        print(f"{'='*60}")

        for anio_actual in anios:
            print(f"\n--- Procesando {nombre_empresa} | Año {anio_actual} ---")

            # Abrir la aplicación para cada empresa/año
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

            # Cargar fechas: 01/01/YYYY hasta 31/12/YYYY
            probar_fechas(app, anio_actual)
            time.sleep(5)

            # Si no hay movimientos, cerrar y pasar al siguiente año
            if not tiene_movimientos(app):
                print(f"SIN MOVIMIENTOS para {nombre_empresa} en {anio_actual}, pasando al siguiente año.")
                registrar_reporte(ruc, nombre_empresa, anio_actual, "N/A", "SIN MOVIMIENTOS", "No hay registros en este año", numero_operacion="")
                cerrar_aplicacion(app)
                time.sleep(3)
                continue

            print(f"MOVIMIENTOS ENCONTRADOS para {nombre_empresa} en {anio_actual}")

            send_keys("^c")
            time.sleep(1)

            print(f"TRABAJANDO CON {nombre_empresa} AÑO {anio_actual}")

            ultimo_asiento = None
            ultima_cuenta = None
            ultimo_asiento_procesado = None
            repetidos = 0
            filas_estancadas = 0
            fila = obtener_ultima_fila(ruc, anio_actual)
            if fila > 0:
                print(f"REANUDANDO DESDE FILA {fila + 1} (progreso guardado)")
            while True:
                y_coord = seleccionar_fila_reg_ctb(app, fila)
                if not y_coord:
                    print(f"No se pudo seleccionar fila {fila}, terminando este año.")
                    break

                time.sleep(2)

                cuenta_raw = obtener_cuenta_contable_grilla(app, y_coord)
                
                # A veces ContaNet copia toda la fila en lugar de solo la celda. Extraer la cuenta.
                cuenta = None
                is_10_4 = False
                if cuenta_raw:
                    # Buscamos cualquier palabra que parezca una cuenta (ej: empieza con 10.)
                    palabras = str(cuenta_raw).split()
                    for p in palabras:
                        if p.startswith("10."):
                            cuenta = p
                            if p.startswith("10.4."):
                                is_10_4 = True
                            break
                    if not cuenta:
                        cuenta = str(cuenta_raw)[:20] # fallback para logs

                # Solo obtenemos el asiento si es 10.4, o si la cuenta se repite para detectar si la grilla se atascó
                if is_10_4 or cuenta == ultima_cuenta:
                    asiento = obtener_asiento_contable()
                else:
                    asiento = None

                # Validar si estamos atascados al final de la grilla (mismos datos visuales exactos)
                if cuenta == ultima_cuenta and asiento is not None and asiento == ultimo_asiento:
                    filas_estancadas += 1
                    if filas_estancadas >= 4:
                        print("Se alcanzó el final de la grilla (la fila no avanza).")
                        break
                else:
                    filas_estancadas = 0

                ultima_cuenta = cuenta
                if asiento is not None:
                    ultimo_asiento = asiento

                # 1) Validar que empiece con 10.4.
                if not is_10_4:
                    print(f"FILA {fila + 1} NO PROCESADO: La cuenta {cuenta} no empieza con 10.4.xxx")
                    fila += 1
                    guardar_ultima_fila(ruc, anio_actual, fila)
                    continue
                
                print(f"EVALUANDO: Cuenta {cuenta} | Asiento {asiento}")
                
                # 2) Solo si empieza con 10.4, aplicamos la regla de cierre por asiento contable repetido
                if asiento == ultimo_asiento_procesado:
                    repetidos += 1
                    if repetidos >= 10:
                        print("Se alcanzó el límite (mismo asiento 10 veces consecutivas para cuentas 10.4).")
                        break
                    else:
                        print(f"Asiento repetido {repetidos} vez/veces, saltando para ver si la grilla avanza...")
                else:
                    repetidos = 0
                
                ultimo_asiento_procesado = asiento

                if es_asiento_procesado(ruc, anio_actual, asiento):
                    print(f"FILA {fila + 1} ASIENTO {asiento} ya procesado anteriormente, saltando.")
                    fila += 1
                    guardar_ultima_fila(ruc, anio_actual, fila)
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
                    # Registrar progreso y pasar a la siguiente fila
                    registrar_asiento_procesado(ruc, anio_actual, asiento)
                    fila += 1
                    guardar_ultima_fila(ruc, anio_actual, fila)
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
                
                if resultado == "VOUCHER NO MODIFICABLE":
                    registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "NO MODIFICADO", "Voucher no modificable", numero_operacion=numero_operacion)
                elif resultado == "ASIGNADO A CAJA":
                    registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "NO MODIFICADO", "El asiento contable ha sido asignado a caja", numero_operacion=numero_operacion)
                else:
                    registrar_reporte(ruc, nombre_empresa, anio_actual, asiento, "ACTUALIZADO", "Se completó la actualización", numero_operacion=numero_operacion)
                    
                registrar_asiento_procesado(ruc, anio_actual, asiento)
                time.sleep(3)
                fila += 1
                guardar_ultima_fila(ruc, anio_actual, fila)

            # Cerrar la aplicación al terminar el año
            print(f"Cerrando aplicación para {nombre_empresa} año {anio_actual}")
            cerrar_aplicacion(app)
            time.sleep(3)

    print("\n" + "="*60)
    print("PROCESO COMPLETADO PARA TODAS LAS EMPRESAS Y AÑOS")
    print("="*60)


def cargar_data_excel():
    pass


def open_contanet():
    pass


def ingresar_datos():
    pass


if __name__ == "__main__":
    main()

