"""
==============================================================
GRAFICOS.PY

Librería de gráficos para el Generador de Informes
Programa Sembremos Seguridad

Toda la apariencia de los gráficos se controla desde este archivo.
==============================================================
"""

# ==========================================================
# IMPORTS
# ==========================================================

from pathlib import Path
import textwrap

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import matplotlib.patheffects as pe

# ==========================================================
# RUTAS
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

# ==========================================================
# CONFIGURACION GENERAL
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

AZUL_CLARO = "#4472C4"

AZUL_MEDIO = "#5B9BD5"

AZUL_OSCURO = "#255E91"

CELESTE = "#9DC3E6"

VERDE_CLARO = "#70AD47"

GRIS = "#A5A5A5"

GRIS_OSCURO = "#636363"

BLANCO = "#FFFFFF"

# ==========================================================
# PALETA GENERAL
# ==========================================================

PALETA = [

    AZUL_MEDIO,

    GRIS,

    AZUL_CLARO,

    AZUL_OSCURO,

    VERDE,

    VERDE_CLARO,

    CELESTE,

    GRIS_OSCURO

]

# ==========================================================
# ESTILOS
# ==========================================================

# ----------------------------------------------------------
# BARRAS PEQUEÑAS
# ----------------------------------------------------------

BARRAS_S = {

    "figsize": (6,4),

    "fontsize_labels": 9,

    "fontsize_valores": 13,

    "fontsize_titulo": 11,

    "ancho_barra": .55,

    "mostrar_grid": False,

    "mostrar_eje_y": False,

    "rotacion": 0,

    "colores": [VERDE]

}

# ----------------------------------------------------------
# BARRAS MEDIANAS
# ----------------------------------------------------------

BARRAS_M = {

    "figsize": (8,5),

    "fontsize_labels": 11,

    "fontsize_valores": 13,

    "fontsize_titulo": 13,

    "ancho_barra": .60,

    "mostrar_grid": False,

    "mostrar_eje_y": False,

    "rotacion": 0,

    "colores": [VERDE]

}

# ----------------------------------------------------------
# BARRAS GRANDES
# ----------------------------------------------------------

BARRAS_L = {

    "figsize": (10,6),

    "fontsize_labels": 13,

    "fontsize_valores": 13,

    "fontsize_titulo": 15,

    "ancho_barra": .65,

    "mostrar_grid": False,

    "mostrar_eje_y": False,

    "rotacion": 0,

    "colores": [VERDE]

}

# ----------------------------------------------------------
# PASTEL PEQUEÑO
# ----------------------------------------------------------

PASTEL_S = {

    "figsize": (5,5),

    "fontsize_labels": 13,

    "fontsize_porcentajes": 13,

    "startangle":90,

    "labeldistance":1.10,

    "pctdistance":0.65,

    "colores": [
        "#5B9BD5",
        "#A5A5A5",
        "#4472C4",
        "#255E91",
        "#30A907",
        "#70AD47",
        "#9DC3E6",
        "#636363"
    ]

}

# ----------------------------------------------------------
# PASTEL MEDIANO
# ----------------------------------------------------------

PASTEL_M = {

    "figsize": (7,7),

    "fontsize_labels": 14,

    "fontsize_porcentajes": 16,

    "startangle":90,

    "labeldistance":1.15,

    "pctdistance":0.65,

    "colores": [
        "#5B9BD5",
        "#A5A5A5",
        "#4472C4",
        "#255E91",
        "#30A907",
        "#70AD47",
        "#9DC3E6",
        "#636363"
    ]

}

# ----------------------------------------------------------
# GRAFICO DE LINEA
# ----------------------------------------------------------

LINEA = {

    "figsize": (10,5),

    "linewidth":3,

    "marker":"o",

    "markersize":7,

    "fontsize_labels":11,

    "fontsize_titulo":14,

    "mostrar_grid":True

}


# ==========================================================
# UTILIDADES
# ==========================================================

def guardar_figura(fig, nombre_archivo):
    """
    Guarda una figura en la carpeta assets y devuelve la ruta.
    """

    ruta = ASSETS_DIR / nombre_archivo

    fig.savefig(
        ruta,
        dpi=DPI,
        transparent=True,
        bbox_inches=None
    )

    plt.close(fig)

    return str(ruta)


# ----------------------------------------------------------

def envolver_texto(labels, ancho=18):
    """
    Divide etiquetas largas en varias líneas.
    """

    resultado = []

    for texto in labels:

        texto = str(texto).replace("_", " ").title()

        texto = "\n".join(
            textwrap.wrap(texto, width=ancho)
        )

        resultado.append(texto)

    return resultado


# ----------------------------------------------------------

def obtener_colores(cantidad, estilo):

    colores = estilo["colores"]

    if len(colores) == 1:

        return colores * cantidad

    if cantidad <= len(colores):

        return colores[:cantidad]

    salida = []

    while len(salida) < cantidad:

        salida.extend(colores)

    return salida[:cantidad]


# ----------------------------------------------------------

def escribir_valores(
    ax,
    barras,
    porcentaje,
    fontsize,
):

    for barra in barras:

        valor = barra.get_height()

        if porcentaje:

            texto = f"{valor:.0f}%"

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

            color=AZUL,

            path_effects=[

                pe.withStroke(

                    linewidth=2,

                    foreground="white"

                )

            ]

        )


# ----------------------------------------------------------

def normalizar_estilo(estilo):

    """
    Completa cualquier propiedad faltante del estilo.
    """

    base = {

        "figsize": (8,5),

        "fontsize_labels":11,

        "fontsize_valores":10,

        "fontsize_titulo":13,

        "fontsize_porcentajes":12,

        "ancho_barra":0.60,

        "mostrar_grid":False,

        "mostrar_eje_y":False,

        "rotacion":0,

        "labeldistance":1.10,

        "pctdistance":0.65,

        "startangle":90,

        "linewidth":2,

        "marker":"o",

        "markersize":6,

        "colores":[VERDE]

    }

    copia = base.copy()

    copia.update(estilo)

    return copia


# ----------------------------------------------------------

def aplicar_estilo_barras(ax, estilo):

    ax.spines["top"].set_visible(False)

    ax.spines["right"].set_visible(False)

    ax.spines["left"].set_visible(False)

    if not estilo["mostrar_eje_y"]:

        ax.spines["bottom"].set_visible(False)

        ax.set_yticks([])

    ax.tick_params(

        axis="x",

        labelsize=estilo["fontsize_labels"],

        length=0

    )

    ax.tick_params(

        axis="y",

        length=0

    )

    if estilo["mostrar_grid"]:

        ax.grid(

            axis="y",

            linestyle="--",

            alpha=.25

        )

# ==========================================================
# GRAFICO DE BARRAS
# ==========================================================

def crear_barras(
    labels,
    valores,
    nombre_archivo,
    estilo=BARRAS_M,
    titulo=None,
    porcentaje=False
):

    # ------------------------------------------
    # Estilo
    # ------------------------------------------

    estilo = normalizar_estilo(estilo)

    labels = envolver_texto(labels)

    valores = [float(v) for v in valores]

    colores = obtener_colores(

        len(labels),

        estilo

    )

    # ------------------------------------------
    # Figura
    # ------------------------------------------

    fig, ax = plt.subplots(

        figsize=estilo["figsize"]

    )

    fig.patch.set_alpha(0)

    ax.set_facecolor("none")

    # ------------------------------------------
    # Barras
    # ------------------------------------------

    barras = ax.bar(

        labels,

        valores,

        width=estilo["ancho_barra"],

        color=colores,

        edgecolor="white",

        linewidth=1.2,

        zorder=3

    )

    # ------------------------------------------
    # Estilo del eje
    # ------------------------------------------

    aplicar_estilo_barras(

        ax,

        estilo

    )

    # ------------------------------------------
    # Rotación
    # ------------------------------------------

    plt.xticks(

        rotation=estilo["rotacion"]

    )

    # ------------------------------------------
    # Limite superior
    # ------------------------------------------

    maximo = max(valores)

    ax.set_ylim(

        0,

        maximo * 1.15

    )

    # ------------------------------------------
    # Valores
    # ------------------------------------------

    escribir_valores(

        ax,

        barras,

        porcentaje,

        estilo["fontsize_valores"]

    )

    # ------------------------------------------
    # Título
    # ------------------------------------------

    if titulo:

        ax.set_title(

            titulo,

            fontsize=estilo["fontsize_titulo"],

            color=AZUL,

            fontweight="bold",

            pad=18

        )

    # ------------------------------------------
    # Márgenes
    # ------------------------------------------

    plt.subplots_adjust(

        left=.08,

        right=.97,

        bottom=.18,

        top=.92

    )

    # ------------------------------------------
    # Guardar
    # ------------------------------------------

    return guardar_figura(

        fig,

        nombre_archivo

    )


# ==========================================================
# GRAFICO DE PASTEL
# ==========================================================

def crear_pastel(
    labels,
    valores,
    nombre_archivo,
    estilo=PASTEL_M,
    mostrar_labels=True
):

    # ------------------------------------------
    # Estilo
    # ------------------------------------------

    estilo = normalizar_estilo(estilo)

    labels = envolver_texto(labels)

    valores = [float(v) for v in valores]

    colores = obtener_colores(

        len(labels),

        estilo

    )

    # ------------------------------------------
    # Figura
    # ------------------------------------------

    fig, ax = plt.subplots(
        figsize=(12,12)
    )

    fig.patch.set_alpha(0)

    ax.set_facecolor("none")

    # ------------------------------------------
    # Pastel
    # ------------------------------------------

    wedges, texts, autotexts = ax.pie(

        valores,

        labels=labels if mostrar_labels else None,

        colors=colores,

        startangle=estilo["startangle"],

        autopct=lambda p: f"{p:.0f}%",

        labeldistance=estilo["labeldistance"],

        pctdistance=estilo["pctdistance"],

        wedgeprops=dict(

            edgecolor="white",

            linewidth=1.5

        )

    )

    # ------------------------------------------
    # Etiquetas
    # ------------------------------------------

    if mostrar_labels:

        for texto in texts:

            texto.set_fontsize(

                estilo["fontsize_labels"]

            )

            texto.set_color(AZUL)

            texto.set_fontweight("bold")

    # ------------------------------------------
    # Porcentajes
    # ------------------------------------------

    for texto in autotexts:

        texto.set_fontsize(

            estilo["fontsize_porcentajes"]

        )

        texto.set_fontweight("bold")

        texto.set_color("white")

        texto.set_path_effects([

            pe.withStroke(

                linewidth=2,

                foreground="black"

            )

        ])

    # ------------------------------------------
    # Mantener círculo perfecto
    # ------------------------------------------

    ax.axis("equal")

    plt.subplots_adjust(
        left=0.05,
        right=0.95,
        top=0.95,
        bottom=0.05
    )

    # ------------------------------------------
    # Guardar
    # ------------------------------------------

    return guardar_figura(

        fig,

        nombre_archivo

    )


# ==========================================================
# GRAFICO DE LINEA
# ==========================================================

def crear_linea(
    labels,
    valores,
    nombre_archivo,
    estilo=LINEA,
    titulo=None
):

    # ------------------------------------------
    # Estilo
    # ------------------------------------------

    estilo = normalizar_estilo(estilo)

    labels = envolver_texto(labels)

    valores = [float(v) for v in valores]

    # ------------------------------------------
    # Figura
    # ------------------------------------------

    fig, ax = plt.subplots(
        figsize=estilo["figsize"]
    )

    fig.patch.set_alpha(0)

    ax.set_facecolor("none")

    # ------------------------------------------
    # Línea
    # ------------------------------------------

    ax.plot(

        labels,

        valores,

        color=AZUL,

        linewidth=estilo["linewidth"],

        marker=estilo["marker"],

        markersize=estilo["markersize"],

        markerfacecolor=VERDE,

        markeredgecolor=AZUL,

        zorder=3

    )

    # ------------------------------------------
    # Valores
    # ------------------------------------------

    for x, y in zip(labels, valores):

        ax.text(

            x,

            y,

            f"{y:.0f}",

            fontsize=estilo["fontsize_labels"],

            fontweight="bold",

            ha="center",

            va="bottom",

            color=AZUL,

            path_effects=[

                pe.withStroke(

                    linewidth=2,

                    foreground="white"

                )

            ]

        )

    # ------------------------------------------
    # Grid
    # ------------------------------------------

    if estilo["mostrar_grid"]:

        ax.grid(

            linestyle="--",

            alpha=.30

        )

    # ------------------------------------------
    # Ejes
    # ------------------------------------------

    ax.spines["top"].set_visible(False)

    ax.spines["right"].set_visible(False)

    ax.tick_params(

        axis="x",

        labelsize=estilo["fontsize_labels"]

    )

    ax.tick_params(

        axis="y",

        labelsize=estilo["fontsize_labels"]

    )

    # ------------------------------------------
    # Título
    # ------------------------------------------

    if titulo:

        ax.set_title(

            titulo,

            fontsize=estilo["fontsize_titulo"],

            fontweight="bold",

            color=AZUL,

            pad=18

        )

    # ------------------------------------------
    # Guardar
    # ------------------------------------------

    return guardar_figura(

        fig,

        nombre_archivo

    )


# ==========================================================
# GRAFICO DE BARRAS HORIZONTALES
# ==========================================================

def crear_barras_horizontal(
    labels,
    valores,
    nombre_archivo,
    estilo=BARRAS_M,
    titulo=None
):

    estilo = normalizar_estilo(estilo)

    labels = envolver_texto(labels)

    valores = [float(v) for v in valores]

    colores = obtener_colores(

        len(labels),

        estilo

    )

    fig, ax = plt.subplots(

        figsize=estilo["figsize"]

    )

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

    ax.tick_params(

        axis="y",

        labelsize=estilo["fontsize_labels"],

        length=0

    )

    ax.set_xticks([])

    for barra in barras:

        valor = barra.get_width()

        ax.text(

            valor,

            barra.get_y()+barra.get_height()/2,

            f"{valor:.0f}",

            va="center",

            ha="left",

            fontsize=estilo["fontsize_valores"],

            fontweight="bold",

            color=AZUL,

            path_effects=[

                pe.withStroke(

                    linewidth=2,

                    foreground="white"

                )

            ]

        )

    if titulo:

        ax.set_title(

            titulo,

            fontsize=estilo["fontsize_titulo"],

            color=AZUL,

            fontweight="bold"

        )

    return guardar_figura(

        fig,

        nombre_archivo

    )


# ==========================================================
# EXPORTACIONES
# ==========================================================

__all__ = [

    "BARRAS_S",
    "BARRAS_M",
    "BARRAS_L",

    "PASTEL_S",
    "PASTEL_M",

    "LINEA",

    "crear_barras",

    "crear_barras_horizontal",

    "crear_pastel",

    "crear_linea",

]
