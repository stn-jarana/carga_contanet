import csv
import os
from collections import defaultdict

def consolidar_reporte_metricas(csv_path):
    if not os.path.exists(csv_path):
        print("El archivo no existe.")
        return

    registros_por_dia = defaultdict(list)
    
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            fecha_hora = row.get("Fecha") or row.get("Fecha_Hora_Reporte")
            if not fecha_hora:
                continue
            fecha = fecha_hora.split(" ")[0]
            registros_por_dia[fecha].append(row)

    resumen_diario = {}
    
    for fecha, filas in registros_por_dia.items():
        # Si ya está consolidado (1 sola fila por fecha y campo 'Fecha' existente con formato simple)
        if len(filas) == 1 and "Total_Asientos_Evaluados" in filas[0]:
            f = filas[0]
            resumen_diario[fecha] = {
                "Fecha": fecha,
                "Total_Asientos_Evaluados": str(f.get("Total_Asientos_Evaluados", 0)),
                "Actualizados_Exito": str(f.get("Actualizados_Exito", 0)),
                "No_Modificados_Total": str(f.get("No_Modificados_Total", 0)),
                "Voucher_No_Modificable": str(f.get("Voucher_No_Modificable", 0)),
                "Asignado_A_Caja": str(f.get("Asignado_A_Caja", 0)),
                "Periodo_Cerrado": str(f.get("Periodo_Cerrado", 0)),
                "No_Encontrado_Excel": str(f.get("No_Encontrado_Excel", 0)),
                "Sin_Movimientos_Mes": str(f.get("Sin_Movimientos_Mes", 0)),
                "Porcentaje_Efectividad": str(f.get("Porcentaje_Efectividad", "0%"))
            }
            continue

        total_evaluadas = 0
        actualizados = 0
        no_modificados = 0
        voucher_no_mod = 0
        caja = 0
        periodo_cerrado = 0
        no_encontrado = 0
        sin_movimientos = 0
        
        sesiones = []
        sesion_actual = []
        last_val = -1
        
        for f in filas:
            curr_val = int(f.get("Total_Operaciones_Evaluadas", f.get("Total_Asientos_Evaluados", 0)))
            if curr_val <= last_val:
                if sesion_actual:
                    sesiones.append(sesion_actual)
                sesion_actual = []
            sesion_actual.append(f)
            last_val = curr_val
        if sesion_actual:
            sesiones.append(sesion_actual)
            
        for s in sesiones:
            ultimo = s[-1]
            total_evaluadas += int(ultimo.get("Total_Operaciones_Evaluadas", ultimo.get("Total_Asientos_Evaluados", 0)))
            actualizados += int(ultimo.get("Actualizados_Exito", 0))
            no_modificados += int(ultimo.get("No_Modificados_Total", 0))
            voucher_no_mod += int(ultimo.get("Voucher_No_Modificable", 0))
            caja += int(ultimo.get("Asignado_A_Caja", 0))
            periodo_cerrado += int(ultimo.get("Periodo_Cerrado", 0))
            no_encontrado += int(ultimo.get("No_Encontrado_Excel", 0))
            sin_movimientos += int(ultimo.get("Sin_Movimientos_Mes", 0))
            
        efectividad = round((actualizados / total_evaluadas) * 100, 2) if total_evaluadas > 0 else 0.0
        
        resumen_diario[fecha] = {
            "Fecha": fecha,
            "Total_Asientos_Evaluados": str(total_evaluadas),
            "Actualizados_Exito": str(actualizados),
            "No_Modificados_Total": str(no_modificados),
            "Voucher_No_Modificable": str(voucher_no_mod),
            "Asignado_A_Caja": str(caja),
            "Periodo_Cerrado": str(periodo_cerrado),
            "No_Encontrado_Excel": str(no_encontrado),
            "Sin_Movimientos_Mes": str(sin_movimientos),
            "Porcentaje_Efectividad": f"{efectividad}%"
        }

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

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for fecha_key in sorted(resumen_diario.keys()):
            writer.writerow(resumen_diario[fecha_key])
            
    print(f"¡Consolidación exitosa de {csv_path}!")

consolidar_reporte_metricas(r"C:\Users\jarana\repositorio\proyect_conta\carga_contanet\reporte_metricas.csv")
