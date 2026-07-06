# ==========================================================
# GRAFICOS.PY
# Librería de gráficos del Generador de Informes
# ==========================================================

from pathlib import Path
import textwrap

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
from matplotlib.ticker import MaxNLocator
# ==========================================================
# RUTAS
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

DPI = 300

FUENTE = "DejaVu Sans"

plt.rcParams["font.family"] = FUENTE

plt.rcParams["figure.facecolor"] = "none"
plt.rcParams["axes.facecolor"] = "none"

# ==========================================================
# COLORES INSTITUCIONALES
# ==========================================================

AZUL = "#013051"

VERDE = "#30A907"

GRIS = "#808080"

GRIS_CLARO = "#D9D9D9"

AZUL_CLARO = "#4472C4"

AZUL_MEDIO = "#5B9BD5"

AZUL_OSCURO = "#255E91"

CELESTE = "#9DC3E6"

VERDE_CLARO = "#70AD47"

GRIS_MEDIO = "#A5A5A5"

GRIS_OSCURO = "#636363"

PALETA = [

    AZUL_MEDIO,

    GRIS_MEDIO,

    AZUL_CLARO,

    AZUL_OSCURO,

    GRIS_OSCURO,

    VERDE,

    VERDE_CLARO,

    CELESTE,

]

# ==========================================================
# CONFIGURACIÓN DE BARRAS
# ==========================================================

FIG_BARRAS = (10,6)

ANCHO_BARRA = 0.60

COLOR_BARRAS = VERDE

COLOR_TEXTO = AZUL

MOSTRAR_EJE_Y = False

MOSTRAR_GRID = False

# ==========================================================
# CONFIGURACIÓN DE PASTELES
# ==========================================================

FIG_PASTEL = (6,6)

STARTANGLE = 90

LABEL_DISTANCE = 1.08

PCT_DISTANCE = 0.72

# ==========================================================
# TIPOGRAFÍAS
# ==========================================================

FONT_TITULO = 16

FONT_LABEL = 12

FONT_PORCENTAJE = 13

FONT_VALOR = 12

# ==========================================================
# MÁRGENES
# ==========================================================

BOTTOM_SPACE = 0.20

TOP_SPACE = 0.95

LEFT_SPACE = 0.08

RIGHT_SPACE = 0.92


# ==========================================================
# UTILIDADES
# ==========================================================

def guardar_figura(fig, nombre_archivo):
    """
    Guarda una figura de matplotlib dentro de la carpeta assets.
    """

    ruta = ASSETS_DIR / nombre_archivo

    fig.savefig(
        ruta,
        dpi=DPI,
        bbox_inches="tight",
        transparent=True
    )

    plt.close(fig)

    return str(ruta)


# ----------------------------------------------------------

def limpiar_porcentajes(valores):
    """
    Convierte porcentajes provenientes de Excel a float.

    Acepta:
        25%
        25,5%
        25.5
        25
    """

    resultado = []

    for valor in valores:

        if pd.isna(valor):
            resultado.append(0)
            continue

        texto = str(valor)

        texto = texto.replace("%", "")
        texto = texto.replace(",", ".")

        try:
            resultado.append(float(texto))
        except:
            resultado.append(0)

    return resultado



# ----------------------------------------------------------

def escribir_valores(

    ax,

    barras,

    porcentaje=True,

    color=COLOR_TEXTO,

    fontsize=FONT_VALOR

):
    """
    Escribe el valor encima de cada barra.
    """

    for barra in barras:

        valor = barra.get_height()

        if porcentaje:
            texto = f"{valor:.2f}%"
        else:
            texto = f"{valor:.0f}"

        ax.text(

            barra.get_x() + barra.get_width()/2,

            valor,

            texto,

            ha="center",

            va="bottom",

            fontsize=fontsize,

            fontweight="bold",

            color=color,

            path_effects=[

                pe.withStroke(

                    linewidth=2,

                    foreground="white"

                )

            ]

        )


# ----------------------------------------------------------

def envolver_texto(labels, ancho=18):
    """
    Divide etiquetas largas en varias líneas.
    """

    nuevas = []

    for texto in labels:

        nuevas.append(

            "\n".join(

                textwrap.wrap(

                    str(texto),

                    width=ancho

                )

            )

        )

    return nuevas


# ----------------------------------------------------------

def obtener_paleta(cantidad):
    """
    Devuelve la cantidad de colores necesarios.
    """

    if cantidad <= len(PALETA):

        return PALETA[:cantidad]

    colores = []

    while len(colores) < cantidad:

        colores.extend(PALETA)

    return colores[:cantidad]


# ==========================================================
# FUNCIÓN MAESTRA - BARRAS VERTICALES
# ==========================================================


def crear_barras(
    labels,
    valores,
    nombre_archivo,
    colores=None,
    titulo=None,
    porcentaje=True,
    figsize=FIG_BARRAS,
    ancho_barra=ANCHO_BARRA,
    fontsize_labels=FONT_LABEL,
    fontsize_valores=FONT_VALOR,
    fontsize_titulo=FONT_TITULO,
    rotacion=0,
    ylim_extra=0.10,
    mostrar_grid=MOSTRAR_GRID,
    mostrar_eje_y=MOSTRAR_EJE_Y,
):

    # ---------------------------------------------
    # Datos
    # ---------------------------------------------

    labels = envolver_texto(labels, ancho=15)

    valores = [float(v) for v in valores]

    if colores is None:
        colores = obtener_paleta(len(labels))

    # ---------------------------------------------
    # Figura
    # ---------------------------------------------

    fig, ax = plt.subplots(figsize=figsize)

    fig.patch.set_alpha(0)

    ax.set_facecolor("none")

    # ---------------------------------------------
    # Barras
    # ---------------------------------------------

    barras = ax.bar(

        labels,

        valores,

        color=colores,

        width=ancho_barra,

        edgecolor="white",

        linewidth=1.2

    )

    # ---------------------------------------------
    # Estilo
    # ---------------------------------------------

    ax.spines["top"].set_visible(False)

    ax.spines["right"].set_visible(False)

    ax.spines["left"].set_visible(False)

    if not mostrar_eje_y:

        ax.set_yticks([])

        ax.spines["bottom"].set_visible(False)

    if mostrar_grid:

        ax.grid(

            axis="y",

            linestyle="--",

            alpha=0.25

        )

    ax.tick_params(

        axis="x",

        labelsize=fontsize_labels,

        length=0

    )

    plt.xticks(rotation=rotacion)

    # ---------------------------------------------
    # Título
    # ---------------------------------------------

    if titulo:

        ax.set_title(

            titulo,

            fontsize=fontsize_titulo,

            color=AZUL,

            fontweight="bold",

            pad=18

        )

    # ---------------------------------------------
    # Altura máxima
    # ---------------------------------------------

    maximo = max(valores)

    ax.set_ylim(

        0,

        maximo * (1 + ylim_extra)

    )

    # ---------------------------------------------
    # Valores sobre barras
    # ---------------------------------------------

    escribir_valores(

        ax,

        barras,

        porcentaje=porcentaje,

        fontsize=fontsize_valores

    )

    # ---------------------------------------------
    # Márgenes
    # ---------------------------------------------

    plt.subplots_adjust(

        left=LEFT_SPACE,

        right=RIGHT_SPACE,

        top=TOP_SPACE,

        bottom=BOTTOM_SPACE

    )

    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    
    # ---------------------------------------------
    # Guardar
    # ---------------------------------------------

    return guardar_figura(

        fig,

        nombre_archivo

    )

# ==========================================================
# FUNCIÓN MAESTRA - PASTEL
# ==========================================================

def crear_pastel(
    labels,
    valores,
    nombre_archivo,
    colores=None,
    figsize=FIG_PASTEL,
    mostrar_labels=True,
    porcentaje_size=FONT_PORCENTAJE,
    label_size=FONT_LABEL,
    startangle=STARTANGLE,
    pctdistance=PCT_DISTANCE,
    labeldistance=LABEL_DISTANCE,
):

    labels = envolver_texto(labels)

    valores = [float(v) for v in valores]

    if colores is None:
        colores = obtener_paleta(len(labels))

    fig, ax = plt.subplots(figsize=figsize)

    fig.patch.set_alpha(0)

    ax.set_facecolor("none")

    wedges, texts, autotexts = ax.pie(

        valores,

        labels=labels if mostrar_labels else None,

        colors=colores,

        startangle=startangle,

        autopct=lambda p: f"{p:.2f}%",

        pctdistance=pctdistance,

        labeldistance=labeldistance,

        wedgeprops=dict(
            edgecolor="white",
            linewidth=1
        )

    )

    for texto in autotexts:

        texto.set_fontsize(porcentaje_size)

        texto.set_fontweight("bold")

        texto.set_color("white")

        texto.set_path_effects([

            pe.withStroke(

                linewidth=2,

                foreground="black"

            )

        ])

    if mostrar_labels:

        for texto in texts:

            texto.set_fontsize(label_size)

            texto.set_color(AZUL)

    ax.axis("equal")

    return guardar_figura(

        fig,

        nombre_archivo

    )


# ==========================================================
# FUNCIÓN MAESTRA - PASTEL CON ETIQUETAS EXTERNAS
# ==========================================================

def crear_pastel_labels(

    labels,

    valores,

    nombre_archivo,

    colores=None,

    figsize=FIG_PASTEL,

    porcentaje_size=FONT_PORCENTAJE,

    label_size=FONT_LABEL,

    startangle=STARTANGLE

):

    valores = [float(v) for v in valores]

    if colores is None:

        colores = obtener_paleta(len(labels))

    fig, ax = plt.subplots(figsize=figsize)

    fig.patch.set_alpha(0)

    ax.set_facecolor("none")

    wedges, _, autotexts = ax.pie(

        valores,

        labels=None,

        colors=colores,

        startangle=startangle,

        autopct=lambda p: f"{p:.2f}%",

        pctdistance=.70,

        wedgeprops=dict(

            edgecolor="white",

            linewidth=1

        )

    )

    for texto in autotexts:

        texto.set_fontsize(porcentaje_size)

        texto.set_fontweight("bold")

        texto.set_color("white")

        texto.set_path_effects([

            pe.withStroke(

                linewidth=2,

                foreground="black"

            )

        ])

    for i, wedge in enumerate(wedges):

        angulo = (wedge.theta1 + wedge.theta2) / 2

        x = np.cos(np.deg2rad(angulo))

        y = np.sin(np.deg2rad(angulo))

        ax.text(

            x * 1.25,

            y * 1.25,

            "\n".join(textwrap.wrap(str(labels[i]),18)),

            ha="center",

            va="center",

            fontsize=label_size,

            color=AZUL,

            fontweight="bold"

        )

    ax.axis("equal")

    return guardar_figura(

        fig,

        nombre_archivo

    )


# ==========================================================
# FUNCIÓN MAESTRA - BARRAS HORIZONTALES
# ==========================================================

def crear_barras_horizontal(
    labels,
    valores,
    nombre_archivo,
    colores=None,
    figsize=(10,6),
    porcentaje=False,
    fontsize_labels=FONT_LABEL,
    fontsize_valores=FONT_VALOR,
    titulo=None,
):

    labels = envolver_texto(labels)

    valores = [float(v) for v in valores]

    if colores is None:
        colores = obtener_paleta(len(labels))

    fig, ax = plt.subplots(figsize=figsize)

    fig.patch.set_alpha(0)
    ax.set_facecolor("none")

    barras = ax.barh(
        labels,
        valores,
        color=colores,
        edgecolor="white",
        linewidth=1.2
    )

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_visible(False)

    ax.set_xticks([])

    ax.tick_params(
        axis="y",
        labelsize=fontsize_labels,
        length=0
    )

    if titulo:

        ax.set_title(
            titulo,
            fontsize=FONT_TITULO,
            fontweight="bold",
            color=AZUL,
            pad=18
        )

    for barra in barras:

        valor = barra.get_width()

        if porcentaje:
            texto = f"{valor:.2f}%"
        else:
            texto = f"{valor:.0f}"

        ax.text(
            valor,
            barra.get_y()+barra.get_height()/2,
            texto,
            va="center",
            ha="left",
            fontsize=fontsize_valores,
            fontweight="bold",
            color=AZUL
        )

    plt.subplots_adjust(
        left=.20,
        right=.96,
        top=.93,
        bottom=.08
    )

    ax.xaxis.set_major_locator(MaxNLocator(integer=True))
        
    return guardar_figura(
        fig,
        nombre_archivo
    )


# ==========================================================
# GRÁFICO TIPO DONA
# ==========================================================

def crear_dona(
    labels,
    valores,
    nombre_archivo,
    colores=None,
    figsize=(6,6),
    ancho=.45
):

    valores=[float(v) for v in valores]

    if colores is None:
        colores=obtener_paleta(len(labels))

    fig,ax=plt.subplots(figsize=figsize)

    fig.patch.set_alpha(0)

    ax.set_facecolor("none")

    wedges,texts,autotexts=ax.pie(

        valores,

        labels=labels,

        colors=colores,

        autopct=lambda p:f"{p:.1f}%",

        startangle=90,

        wedgeprops=dict(
            width=ancho,
            edgecolor="white"
        )

    )

    for t in texts:

        t.set_fontsize(FONT_LABEL)

        t.set_color(AZUL)

    for t in autotexts:

        t.set_fontsize(FONT_PORCENTAJE)

        t.set_fontweight("bold")

        t.set_color("white")

    centro=plt.Circle(

        (0,0),

        1-ancho,

        fc="white"

    )

    ax.add_artist(centro)

    ax.axis("equal")

    return guardar_figura(
        fig,
        nombre_archivo
    )


# ==========================================================
# LIMPIAR FIGURAS
# ==========================================================

def cerrar_figuras():
    """
    Cierra todas las figuras abiertas.
    """

    plt.close("all")


# ==========================================================
# EXPORTACIÓN
# ==========================================================

__all__ = [

    "crear_barras",

    "crear_barras_horizontal",

    "crear_pastel",

    "crear_pastel_labels",

    "crear_dona",

    "guardar_figura",

    "limpiar_porcentajes",

    "aplicar_estilo",

    "escribir_valores",

    "envolver_texto",

    "obtener_paleta",

    "cerrar_figuras",

]


def colores_verde(cantidad):
    return [VERDE] * cantidad


def colores_azul(cantidad):
    return [AZUL] * cantidad


def colores_paleta(cantidad):
    return obtener_paleta(cantidad)
    
    

    
