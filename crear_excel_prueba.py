"""
Script para generar un archivo Excel de prueba con datos para el mapa de calor.
"""
from pathlib import Path

import pandas as pd
import numpy as np

# Semilla para reproducibilidad
np.random.seed(42)

meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

productos = [
    "Producto A",
    "Producto B",
    "Producto C",
    "Producto D",
    "Producto E",
    "Producto F",
]

# Generar datos de ventas aleatorios (unidades vendidas por mes)
datos = np.random.randint(50, 500, size=(len(productos), len(meses)))

df = pd.DataFrame(datos, index=productos, columns=meses)

# Guardar en Excel
carpeta = Path("input")
carpeta.mkdir(exist_ok=True)
ruta_excel = carpeta / "datos_prueba.xlsx"
df.to_excel(ruta_excel, index=True, sheet_name="Ventas")

print(f"Archivo Excel creado: {ruta_excel}")
print("\nVista previa de los datos:")
print(df.to_string())
