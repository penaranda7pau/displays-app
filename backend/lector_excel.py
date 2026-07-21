import os
import json
import glob
import pandas as pd

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "..", "config", "config.json")

_DEFAULTS = {
    "ruta_wm":   r"C:\Users\HP 645\OneDrive - CPFR\CPFR Advisor - General\reportes generales\Mundo Vegano\DISPLAYS\BD INV\WM",
    "ruta_auto": r"C:\Users\HP 645\OneDrive - CPFR\CPFR Advisor - General\reportes generales\Mundo Vegano\DISPLAYS\BD INV\AUTO",
}

def _cargar_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, ValueError):
        cfg = {}
    return {**_DEFAULTS, **cfg}

_config = _cargar_config()
RUTA_WM   = _config["ruta_wm"]
RUTA_AUTO = _config["ruta_auto"]

def get_primer_excel(ruta):
    archivos = glob.glob(os.path.join(ruta, "*.xlsx"))
    if not archivos:
        archivos = glob.glob(os.path.join(ruta, "*.xls"))
    return archivos[0] if archivos else None

def leer_inventario_wm():
    archivo = get_primer_excel(RUTA_WM)
    if not archivo:
        return {}
    df = pd.read_excel(archivo, header=20, dtype=str)
    col_tienda     = df.columns[3]   # D - Nombre de la Tienda
    col_producto   = df.columns[5]   # F - Descripción de Señalización
    col_estatus    = df.columns[6]   # G - Estatus del Artículo
    col_inventario = df.columns[68]  # BQ - Cantidad en Existencia actual
    df = df[df[col_estatus].str.strip() == "A"]
    resultado = {}
    for tienda, grupo in df.groupby(col_tienda):
        tienda = str(tienda).strip()
        productos = []
        for _, row in grupo.iterrows():
            nombre = str(row[col_producto]).strip() if pd.notna(row[col_producto]) else ""
            if not nombre:
                continue
            try:
                cantidad = int(float(str(row[col_inventario]).replace(",", ".")))
            except:
                cantidad = 0
            productos.append({"nombre": nombre, "cantidad": cantidad})
        nombres_vistos = set()
        unicos = []
        for p in productos:
            if p["nombre"] not in nombres_vistos:
                nombres_vistos.add(p["nombre"])
                unicos.append(p)
        if unicos:
            resultado[tienda] = sorted(unicos, key=lambda x: x["cantidad"], reverse=True)
    return resultado

def leer_inventario_auto():
    archivo = get_primer_excel(RUTA_AUTO)
    if not archivo:
        return {}
    df = pd.read_excel(archivo, dtype=str)
    col_producto   = df.columns[3]   # D - NombreProducto
    col_tienda     = df.columns[31]  # AG - NombrePDX
    col_inventario = df.columns[32]  # AH - SaldoDisponible
    resultado = {}
    for tienda, grupo in df.groupby(col_tienda):
        tienda = str(tienda).strip()
        productos = []
        for _, row in grupo.iterrows():
            nombre = str(row[col_producto]).strip() if pd.notna(row[col_producto]) else ""
            if not nombre:
                continue
            try:
                cantidad = int(float(str(row[col_inventario]).replace(",", ".")))
            except:
                cantidad = 0
            productos.append({"nombre": nombre, "cantidad": cantidad})
        nombres_vistos = set()
        unicos = []
        for p in productos:
            if p["nombre"] not in nombres_vistos:
                nombres_vistos.add(p["nombre"])
                unicos.append(p)
        if unicos:
            resultado[tienda] = sorted(unicos, key=lambda x: x["cantidad"], reverse=True)
    return resultado

def leer_inventario_completo():
    inventario = {}
    inventario.update(leer_inventario_wm())
    inventario.update(leer_inventario_auto())
    return inventario
