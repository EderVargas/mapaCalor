# Mapa de Calor con Seaborn

Herramienta de línea de comandos para generar mapas de calor a partir de archivos Excel, usando **Seaborn** y **Matplotlib**.

**Autor:** EderVargas

---

## Requisitos previos

- Python 3.10+
- Entorno virtual en `.venv/` (local, no va al repo)

### Crear y activar el entorno virtual

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Paquetes instalados:

| Paquete      | Uso                                  |
|-------------|---------------------------------------|
| `pandas`    | Lectura y manejo del Excel            |
| `openpyxl`  | Motor de lectura de archivos `.xlsx`  |
| `seaborn`   | Generación del mapa de calor          |
| `matplotlib`| Renderizado y exportación del gráfico |

---

## Archivos del proyecto

```
mapaCalor/
├── mapa_calor.py           ← Herramienta principal
├── generar_mapas.ps1       ← Genera un mapa por cada hoja del Excel
├── crear_excel_prueba.py   ← Genera el Excel de ejemplo en input/
├── requirements.txt        ← Dependencias
├── README.md               ← Este archivo
├── input/                  ← Excel de entrada (no va a git)
├── output/                 ← Mapas generados (no va a git)
└── tmp/                    ← Zips y copias para revisar (no va a git)
```

Coloca el Excel en `input/` y corre `generar_mapas.ps1`. Por defecto los PNG salen en `output/mapas_YYYYMMDD_HHMMSS/`.

```powershell
.\generar_mapas.ps1                  # PNG (por defecto)
.\generar_mapas.ps1 -Formato png     # igual
.\generar_mapas.ps1 -Formato pkl     # figura Matplotlib para zoom/pan
.\generar_mapas.ps1 -Formato pdf
.\generar_mapas.ps1 -Formato svg
.\generar_mapas.ps1 -Formato jpg
.\generar_mapas.ps1 -Formato show    # ventana al generar; no guarda
```

`-Formato pkl` guarda `.pkl`. Abrir después:

```powershell
python mapa_calor.py --mostrar output\mapas_YYYYMMDD_HHMMSS\mapa1.pkl
```

---
## SOLO EN CASO DE NO EXISTIR EXCEL Y REQUERIR UN ARCHIVO DE PRUEBA PARA EJECUTAR `mapa_calor.py` (opcional)

Si no existe aún `input/datos_prueba.xlsx`, ejecútalo una sola vez:

```powershell
python crear_excel_prueba.py
```

Crea una hoja **Ventas** con 6 productos × 12 meses con valores aleatorios de entre 50 y 500 unidades.

---

## Uso de `mapa_calor.py`

### Sintaxis

```
usage: mapa_calor.py [-h] [--hoja HOJA] [--titulo TITULO] [--paleta PALETA]
                     [--decimales N] [--sin-anotaciones] [--salida SALIDA]
                     [--formato {png,pdf,svg,jpg,pkl,show}] [--mostrar PKL]
                     [excel]
```

### Argumentos

| Argumento           | Tipo       | Valor por defecto                           | Descripción                                                              |
|--------------------|------------|----------------------------------------------|--------------------------------------------------------------------------|
| `excel`            | posicional | `input/datos_prueba.xlsx`                    | Ruta al archivo Excel a leer.                                            |
| `--hoja HOJA`      | opcional   | Primera hoja del libro                       | Nombre de la hoja a leer.                                                |
| `--titulo TITULO`  | opcional   | `"Sin Título"` | Título que aparece en la parte superior del gráfico.                     |
| `--paleta PALETA`  | opcional   | `beige_vino`                                | Paleta de colores (ver tabla de paletas más abajo).                      |
| `--decimales N`    | opcional   | `0`                                         | Decimales en celdas. Miles y millones con coma (`1,234,567`).            |
| `--sin-anotaciones`| flag       | _(desactivado)_                             | Oculta los valores numéricos dentro de cada celda del mapa.              |
| `--salida SALIDA`  | opcional   | _(muestra en pantalla)_                     | Guarda el gráfico en el archivo indicado.                                |
| `--formato`        | opcional   | se infiere de `--salida` o `show`           | `png`, `pdf`, `svg`, `jpg`, `pkl` (reabrir con `--mostrar`) o `show`.    |
| `--mostrar PKL`    | opcional   | —                                           | Abre un `.pkl` en la ventana interactiva de Matplotlib.                  |
| `-h`, `--help`     | flag       | —                                           | Muestra la ayuda y termina.                                              |

---

## Ejemplos de uso

### 1. Uso básico — muestra el gráfico en pantalla

```powershell
python mapa_calor.py
```

### 2. Especificar un archivo Excel diferente

```powershell
python mapa_calor.py input/ventas_2025.xlsx
```

### 3. Leer una hoja específica del Excel

```powershell
python mapa_calor.py input/ventas_2025.xlsx --hoja "Enero-Junio"
```

### 4. Personalizar el título

```powershell
python mapa_calor.py --titulo "Ventas regionales 2025"
```

### 5. Cambiar la paleta de colores

```powershell
python mapa_calor.py --paleta viridis
```

### 6. Ocultar los valores numéricos en las celdas

```powershell
python mapa_calor.py --sin-anotaciones
```

### 7. Guardar el gráfico como imagen PNG

```powershell
python mapa_calor.py --salida output/mapa_calor.png
```

### 7b. Guardar figura interactiva (`.pkl`) y reabrirla

```powershell
python mapa_calor.py --formato pkl --salida output/mapa_calor.pkl
python mapa_calor.py --mostrar output/mapa_calor.pkl
```

### 8. Combinar varias opciones

```powershell
python mapa_calor.py input/ventas_2025.xlsx --hoja "Resumen" --titulo "Ventas 2025" --paleta coolwarm --salida output/resultado.png
```

### 9. Ver la ayuda del comando

```powershell
python mapa_calor.py --help
```

---

## Paletas de colores disponibles

| Paleta        | Descripción                                      |
|---------------|--------------------------------------------------|
| `beige_vino`  | Beige → camel → vino / borgoña _(por defecto)_   |
| `YlOrRd`      | Amarillo → Naranja → Rojo                        |
| `Blues`       | Degradado de azules                              |
| `viridis`     | Morado → Verde → Amarillo (accesible para daltonismo) |
| `coolwarm`    | Azul (frío) → Rojo (caliente)                    |
| `RdYlGn`      | Rojo → Amarillo → Verde (tipo semáforo)          |
| `magma`       | Negro → Violeta → Amarillo                       |

Para una lista completa de paletas disponibles consulta la [documentación de Matplotlib](https://matplotlib.org/stable/gallery/color/colormap_reference.html).

---

## Formato esperado del Excel

El archivo Excel debe tener la siguiente estructura:

|               | Col 1 | Col 2 | Col 3 | ... |
|---------------|-------|-------|-------|-----|
| **Fila 1**    | valor | valor | valor | ... |
| **Fila 2**    | valor | valor | valor | ... |
| ...           | ...   | ...   | ...   | ... |

- La **primera columna** es el índice de filas (ej: nombres de productos).
- Las **columnas restantes** son categorías del eje X (ej: meses).
- Los **valores** deben ser numéricos.

---

## Notas

- Si no se usa `--salida` (o se usa `--formato show`), el gráfico se muestra en una ventana interactiva.
- `--formato pkl` guarda la figura de Matplotlib. Se reabre con `--mostrar` (misma versión de Matplotlib recomendada).
- Los formatos de imagen soportados: `.png`, `.jpg`, `.pdf`, `.svg`.
- Cualquier valor no numérico en el Excel se convierte a `NaN` y se muestra como celda vacía en el mapa.
- Las cifras en celdas y en la barra de color usan coma de miles: `1,234` / `1,234,567.00`.
