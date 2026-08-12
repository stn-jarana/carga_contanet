import csv

with open("reporte_asientos.csv", "r", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    filas = list(reader)

print(f"Total registros en reporte_asientos.csv: {len(filas)}")
estados = {}
for r in filas:
    st = r.get("Estado")
    estados[st] = estados.get(st, 0) + 1
print("Conteo por Estado:", estados)
