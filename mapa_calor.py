"""
Herramienta para generar mapas de calor a partir de un archivo Excel usando Seaborn.
"""
import argparse
import pickle
import sys
import textwrap
from pathlib import Path

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.ticker import StrMethodFormatter

FORMATOS_ARCHIVO = ("png", "pdf", "svg", "jpg", "pkl")
FORMATOS = FORMATOS_ARCHIVO + ("show",)

NOMBRE_PALETA_DEFAULT = "beige_vino"
COLORES = [
    "#F3E6C8",  # Beige muy claro
    "#E8D3A8",  # Beige
    "#D7B77A",  # Beige dorado
    "#B98A58",  # Camel / arena
    "#B38E5D",  # Marrón rojizo
    "#6E3038",  # Vino medio
    "#5A1F2A",  # Vino
    "#9D2449",  # Borgoña profundo
    "#48131F",  # Vino oscuro
]
_CMAP_DEFAULT = LinearSegmentedColormap.from_list(NOMBRE_PALETA_DEFAULT, COLORES)


def _resolver_paleta(paleta: str):
    """Devuelve el colormap custom o un nombre de Matplotlib/Seaborn."""
    if not paleta or paleta.lower() in (NOMBRE_PALETA_DEFAULT, "default"):
        return _CMAP_DEFAULT
    return paleta


def _esta_vacio(valor) -> bool:
    if valor is None:
        return True
    if isinstance(valor, float) and pd.isna(valor):
        return True
    texto = str(valor).strip()
    return texto == "" or texto.lower() == "nan"


def _texto_original(valor) -> str:
    """Conserva el texto entregado. Solo recorta espacios sobrantes."""
    if _esta_vacio(valor):
        return ""
    return " ".join(str(valor).split())


def _envolver(texto: str, ancho: int) -> list[str]:
    if not texto:
        return []
    if len(texto) <= ancho:
        return [texto]
    return textwrap.wrap(
        texto,
        width=ancho,
        break_long_words=False,
        break_on_hyphens=False,
    )


def partir_etiqueta(valor, ancho: int) -> str:
    """Parte etiquetas largas en renglones. No acorta ni cambia palabras."""
    if _esta_vacio(valor):
        return ""
    crudo = str(valor).strip()
    if "\n" in crudo:
        partes = [partir_etiqueta(linea, ancho) for linea in crudo.split("\n")]
        return "\n".join(p for p in partes if p)

    texto = " ".join(crudo.split())
    if not texto:
        return ""

    idx = texto.find(" (")
    if idx > 0:
        cabeza = texto[:idx]
        cola = texto[idx + 1 :]
        return "\n".join([cabeza, *_envolver(cola, ancho)])

    lineas = _envolver(texto, ancho)
    return "\n".join(lineas) if lineas else texto


def _max_lineas(etiquetas) -> int:
    return max((str(e).count("\n") + 1 for e in etiquetas), default=1)


def _max_ancho_linea(etiquetas) -> int:
    maximo = 1
    for etiqueta in etiquetas:
        for linea in str(etiqueta).split("\n"):
            maximo = max(maximo, len(linea))
    return maximo


def _anchos_envoltura(n_filas: int, n_cols: int, max_len_y: int) -> tuple[int, int]:
    if n_filas >= 36:
        ancho_y = 36
    elif max_len_y > 80:
        ancho_y = 34
    elif n_filas >= 18:
        ancho_y = 26
    else:
        ancho_y = 26

    if n_cols == 1:
        ancho_x = 24
    elif n_cols >= 8:
        ancho_x = 80
    else:
        ancho_x = 20
    return ancho_y, ancho_x


def preparar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Quita filas vacías y une renglones de monto que Excel partió."""
    df = df.copy()
    df = df.loc[[not _esta_vacio(i) for i in df.index]]

    indices = [_texto_original(i) for i in df.index]
    eliminar = []
    for i, etiqueta in enumerate(indices):
        if etiqueta.startswith("(Monto") and i > 0:
            indices[i - 1] = f"{indices[i - 1]}\n{etiqueta}"
            eliminar.append(i)
    df.index = indices
    if eliminar:
        df = df.drop(df.index[eliminar])

    n_filas, n_cols = df.shape
    n_cols = max(n_cols, 1)
    max_len_y = max((len(i) for i in df.index), default=1)
    ancho_y, ancho_x = _anchos_envoltura(n_filas, n_cols, max_len_y)
    df.index = [partir_etiqueta(i, ancho_y) for i in df.index]
    df.columns = [partir_etiqueta(c, ancho_x) for c in df.columns]
    return df


def _tamanos_fuente(n_filas: int, n_cols: int, lineas_y: int) -> dict:
    if n_filas > 36:
        ytick = 7
    elif n_filas > 28:
        ytick = 8
    elif n_filas > 14 or lineas_y >= 4:
        ytick = 9
    else:
        ytick = 10

    if n_cols == 1:
        xtick = 10
        annot = 11
    elif n_cols >= 8:
        xtick = 8
        annot = 8
    else:
        xtick = 8
        annot = 9

    return {"ytick": ytick, "xtick": xtick, "annot": annot, "titulo": 14}


def _alto_celda(n_filas: int, lineas_y: int, ytick: int) -> float:
    por_linea = (ytick / 72.0) * 1.35 + 0.10
    alto = max(0.48, lineas_y * por_linea)
    if n_filas >= 40:
        return min(alto, 0.40)
    if n_filas >= 22:
        return min(alto, 0.52)
    return min(alto, 1.55)


def cargar_datos(ruta_excel: str, hoja: str = None) -> pd.DataFrame:
    """Carga los datos desde un archivo Excel."""
    try:
        sheet = hoja if hoja is not None else 0
        df = pd.read_excel(ruta_excel, sheet_name=sheet, index_col=0)
        print(f"Datos cargados correctamente desde '{ruta_excel}'.")
        print(f"  Forma: {df.shape[0]} filas x {df.shape[1]} columnas")
        return df
    except FileNotFoundError:
        print(f"[ERROR] No se encontró el archivo: {ruta_excel}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] No se pudo cargar el archivo: {e}")
        sys.exit(1)


def _normalizar_formato(formato: str | None, ruta_salida: str | None) -> str:
    if formato:
        elegido = formato.lower()
    elif ruta_salida:
        elegido = Path(ruta_salida).suffix.lstrip(".").lower() or "png"
    else:
        return "show"
    if elegido == "jpeg":
        elegido = "jpg"
    if elegido not in FORMATOS:
        print(f"[ERROR] Formato no soportado: {elegido}. Usa: {', '.join(FORMATOS)}")
        sys.exit(1)
    return elegido


def _ruta_con_formato(ruta_salida: str, formato: str) -> Path:
    ruta = Path(ruta_salida)
    sufijo = ".jpg" if formato == "jpg" else f".{formato}"
    return ruta.with_suffix(sufijo)


def _fondo_transparente(fig, ax) -> None:
    """Quita el fondo blanco de figura, ejes y barra de color."""
    fig.patch.set_facecolor("none")
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    ax.patch.set_alpha(0)
    if not ax.collections:
        return
    cbar = ax.collections[0].colorbar
    if cbar is None:
        return
    cbar.ax.set_facecolor("none")
    cbar.ax.patch.set_alpha(0)


def _guardar_figura(fig, ruta: Path, formato: str) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    if formato == "pkl":
        with ruta.open("wb") as fh:
            pickle.dump(fig, fh, protocol=pickle.HIGHEST_PROTOCOL)
        print(f"Figura interactiva guardada en: {ruta}")
        print(f"  Abrir: python mapa_calor.py --mostrar {ruta}")
        return
    kwargs = {
        "dpi": 150,
        "bbox_inches": "tight",
        "pad_inches": 0.25,
        "transparent": True,
        "facecolor": "none",
        "edgecolor": "none",
    }
    if formato == "jpg":
        # JPEG no soporta canal alpha: se aplana sobre blanco.
        kwargs["format"] = "jpeg"
        kwargs["transparent"] = False
        kwargs["facecolor"] = "white"
        kwargs["edgecolor"] = "white"
    fig.savefig(ruta, **kwargs)
    print(f"Gráfico guardado en: {ruta}")


def mostrar_figura(ruta_pkl: str) -> None:
    """Abre un .pkl y muestra la ventana interactiva de Matplotlib."""
    ruta = Path(ruta_pkl)
    if not ruta.is_file():
        print(f"[ERROR] No se encontró el archivo: {ruta}")
        sys.exit(1)
    try:
        with ruta.open("rb") as fh:
            fig = pickle.load(fh)
    except Exception as e:
        print(f"[ERROR] No se pudo abrir la figura: {e}")
        sys.exit(1)
    plt.show()
    plt.close(fig)


def generar_mapa_calor(
    df: pd.DataFrame,
    titulo: str = "Mapa de Calor",
    paleta: str = NOMBRE_PALETA_DEFAULT,
    mostrar_anotaciones: bool = True,
    decimales: int = 0,
    ruta_salida: str = None,
    formato: str = None,
):
    """
    Genera y muestra (o guarda) un mapa de calor con Seaborn.

    Parámetros
    ----------
    df               : DataFrame con los datos (índice=filas, columnas=columnas del mapa).
    titulo           : Título del gráfico.
    paleta           : Paleta de colores. Por defecto beige_vino (custom).
    mostrar_anotaciones : Si True, muestra los valores numéricos en cada celda.
    decimales        : Número de decimales a mostrar en las anotaciones (por defecto 0).
    ruta_salida      : Si se indica, guarda el gráfico en esa ruta en lugar de mostrarlo.
    formato          : png, pdf, svg, jpg, pkl o show. Si falta, se infiere de la ruta.
    """
    df_etiquetas = preparar_dataframe(df)
    df_numerico = df_etiquetas.apply(pd.to_numeric, errors="coerce")

    n_filas, n_cols = df_numerico.shape
    lineas_y = _max_lineas(df_numerico.index)
    lineas_x = _max_lineas(df_numerico.columns)
    max_y = _max_ancho_linea(df_numerico.index)
    max_x = _max_ancho_linea(df_numerico.columns)
    fuentes = _tamanos_fuente(n_filas, n_cols, lineas_y)

    margen_y = max(1.7, max_y * (fuentes["ytick"] / 72.0) * 0.62)
    ancho_celda = 1.55 if n_cols == 1 else (1.10 if n_cols < 8 else 1.05)
    if n_cols >= 8:
        extra_x = max(2.6, max_x * (fuentes["xtick"] / 72.0) * 0.70)
    elif n_cols == 1:
        extra_x = 0.55 * lineas_x
    else:
        extra_x = 0.42 * lineas_x + 1.1
    fig_w = min(24, max(7.5, margen_y + n_cols * ancho_celda + 1.8))
    fig_h = min(32, max(4.2, n_filas * _alto_celda(n_filas, lineas_y, fuentes["ytick"]) + 2.2 + extra_x))

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))

    fmt_str = f",.{decimales}f" if mostrar_anotaciones else ""

    sns.heatmap(
        df_numerico,
        annot=mostrar_anotaciones,
        fmt=fmt_str,
        cmap=_resolver_paleta(paleta),
        linewidths=0.5,
        linecolor="white",
        annot_kws={"size": fuentes["annot"], "weight": "medium"},
        cbar_kws={"label": "Valor", "shrink": 0.65},
        ax=ax,
    )
    cbar = ax.collections[0].colorbar
    if cbar is not None:
        cbar.formatter = StrMethodFormatter(f"{{x:,.{decimales}f}}")
        cbar.update_ticks()

    titulo_visible = "\n".join(textwrap.wrap(titulo, width=56)) or titulo
    ax.set_title(titulo_visible, fontsize=fuentes["titulo"], fontweight="bold", pad=14)
    ax.set_xlabel(df.columns.name or "", fontsize=11)
    ax.set_ylabel(df.index.name or "", fontsize=11)

    if n_cols >= 8:
        plt.setp(
            ax.get_xticklabels(),
            rotation=90,
            ha="center",
            va="top",
            fontsize=fuentes["xtick"],
        )
    elif n_cols > 1:
        plt.setp(
            ax.get_xticklabels(),
            rotation=38,
            ha="right",
            rotation_mode="anchor",
            fontsize=fuentes["xtick"],
            linespacing=1.05,
        )
    else:
        plt.setp(
            ax.get_xticklabels(),
            rotation=0,
            ha="center",
            fontsize=fuentes["xtick"],
            linespacing=1.05,
        )

    plt.setp(
        ax.get_yticklabels(),
        rotation=0,
        ha="right",
        va="center",
        fontsize=fuentes["ytick"],
        linespacing=1.12,
    )
    ax.tick_params(axis="both", length=0, pad=4)

    fig.tight_layout()
    _fondo_transparente(fig, ax)

    formato_final = _normalizar_formato(formato, ruta_salida)
    if formato_final == "show":
        plt.show()
    elif not ruta_salida:
        print("[ERROR] --salida es obligatorio salvo con --formato show.")
        plt.close(fig)
        sys.exit(1)
    else:
        _guardar_figura(fig, _ruta_con_formato(ruta_salida, formato_final), formato_final)

    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(
        description="Genera un mapa de calor a partir de un archivo Excel."
    )
    parser.add_argument(
        "excel",
        nargs="?",
        default="input/datos_prueba.xlsx",
        help="Ruta al archivo Excel (por defecto: input/datos_prueba.xlsx).",
    )
    parser.add_argument(
        "--hoja",
        default=None,
        help="Nombre de la hoja a leer (por defecto la primera).",
    )
    parser.add_argument(
        "--titulo",
        default="Sin Título",
        help="Título del gráfico.",
    )
    parser.add_argument(
        "--paleta",
        default=NOMBRE_PALETA_DEFAULT,
        help=(
            "Paleta de colores. Por defecto: beige_vino (beige a vino). "
            "Otras: YlOrRd, Blues, viridis, coolwarm, RdYlGn, magma."
        ),
    )
    parser.add_argument(
        "--decimales",
        type=int,
        default=0,
        help="Decimales en anotaciones. Miles y millones con coma (1,234,567).",
    )
    parser.add_argument(
        "--sin-anotaciones",
        action="store_true",
        help="Oculta los valores numéricos dentro de las celdas.",
    )
    parser.add_argument(
        "--salida",
        default=None,
        help="Ruta de salida para guardar el gráfico (ej: mapa_calor.png).",
    )
    parser.add_argument(
        "--formato",
        choices=list(FORMATOS),
        default=None,
        help=(
            "png, pdf, svg, jpg: archivo. "
            "pkl: figura para reabrir con --mostrar. "
            "show: ventana interactiva. "
            "Si falta, se infiere de --salida o se usa show."
        ),
    )
    parser.add_argument(
        "--mostrar",
        default=None,
        help="Abre un .pkl guardado en la ventana interactiva de Matplotlib.",
    )

    args = parser.parse_args()

    if args.mostrar:
        mostrar_figura(args.mostrar)
        return

    df = cargar_datos(args.excel, hoja=args.hoja)

    generar_mapa_calor(
        df,
        titulo=args.titulo,
        paleta=args.paleta,
        mostrar_anotaciones=not args.sin_anotaciones,
        decimales=args.decimales,
        ruta_salida=args.salida,
        formato=args.formato,
    )


if __name__ == "__main__":
    main()
