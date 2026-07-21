import sys
import requests
sys.path.insert(0, "backend")
from lector_excel import leer_inventario_completo

URL  = "https://displays-app.onrender.com/api/sync"
KEY  = "cpfr2024"

print("Leyendo archivos Excel...")
inventario = leer_inventario_completo()
tiendas = len(inventario)
productos = sum(len(v) for v in inventario.values())
print(f"  {tiendas} tiendas, {productos} productos encontrados")

print("Enviando a Render...")
r = requests.post(URL, json={"key": KEY, "inventario": inventario}, timeout=60)
data = r.json()
if data.get("ok"):
    print(f"Sincronización exitosa: {data['productos']} productos subidos")
else:
    print("Error:", data)
