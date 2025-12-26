import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import base64
from io import BytesIO
from pathlib import Path

# ================== CONFIG ==================
st.set_page_config(page_title="Generador de Informes", layout="wide")
st.title("Generador de Informes")

# ================== UTILIDADES ==================
def img_to_base64(path: Path) -> str:
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()

def fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=200)
    plt.close(fig)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()

def extraer_bloque(df, fila_ini, fila_fin, col_label, col_valor, factor=1):
    bloque = df.iloc[fila_ini:fila_fin, [col_label, col_valor]].copy()
    bloque.columns = ["label", "value"]
    bloque["value"] = pd.to_numeric(bloque["value"], errors="coerce") * factor
    bloque = bloque.dropna()
    return bloque

def grafico_pie(bloque, titulo):
    if bloque.empty:
        return ""
    fig, ax = plt.subplots()
    ax.pie(
        bloque["value"],
        labels=bloque["label"],
        autopct="%1.0f%%",
        startangle=90
    )
    ax.axis("equal")
    ax.set_title(titulo)
    return fig_to_base64(fig)

def grafico_barras(bloque, titulo):
    if bloque.empty:
        return ""
    fig, ax = plt.subplots()
    ax.bar(bloque["label"], bloque["value"])
    ax.set_ylim(0, 100)
    ax.set_ylabel("Porcentaje")
    ax.set_title(titulo)
    plt.xticks(rotation=20)
    return fig_to_base64(fig)

# ================== CARGA HTML / CSS ==================
html = Path("Informe.html").read_text(encoding="utf-8")
css = Path("estilos.css").read_text(encoding="utf-8")

# 🔴 QUITAR LINK CSS (SOLO PARA STREAMLIT)
html = html.replace(
    '<link rel="stylesheet" href="estilos.css">',
    ''
)

# ================== IMÁGENES BASE64 ==================
imagenes = [
    "portada.png","intro.png","conformacion.png","participacion.png",
    "metodologico.png","estadistica.png","lineas.png",
    "percepcion.png","final.png","header.png","footer.png"
]

for img in imagenes:
    p = Path(img)
    if p.exists():
        html = html.replace(
            f'src="{img}"',
            f'src="{img_to_base64(p)}"'
        )

# ================== SUBIR EXCEL ==================
archivo_excel = st.file_uploader("Subir matriz Excel", type=["xlsx"])

if archivo_excel:
    part = pd.read_excel(archivo_excel, sheet_name="participacion", header=None)
    rel = pd.read_excel(archivo_excel, sheet_name="relevante", header=None)

    # ================== PARTICIPACIÓN ==================
    html = html.replace(
        "{{tabla_distritos}}",
        f"<tr><td>{part.iloc[33,1]}</td><td>{part.iloc[33,2]*100:.0f}%</td><td>{int(part.iloc[33,3])}</td></tr>"
    )

    html = html.replace(
        "{{grafico_relacion}}",
        grafico_barras(extraer_bloque(part,33,34,4,6,100),"Relación de participación")
    )

    html = html.replace(
        "{{grafico_edad}}",
        grafico_pie(extraer_bloque(part,1,6,1,2,100),"Participación por edad")
    )

    html = html.replace(
        "{{grafico_escolaridad}}",
        grafico_pie(extraer_bloque(part,11,18,1,2,100),"Participación por escolaridad")
    )

    html = html.replace(
        "{{grafico_genero}}",
        grafico_pie(extraer_bloque(part,23,26,1,2,100),"Participación por género")
    )

    # ================== PERCEPCIÓN CIUDADANA ==================
    html = html.replace(
        "{{grafico_seguridad_comunidad}}",
        grafico_pie(extraer_bloque(rel,1,4,0,1),"¿Se siente seguro en su comunidad?")
    )

    html = html.replace(
        "{{grafico_seguridad_barrio}}",
        grafico_barras(extraer_bloque(rel,6,9,0,1),"Seguridad en el barrio este año")
    )

    tabla = ""
    for i in range(1,4):
        tabla += f"<tr><td>{rel.iloc[i,0]}</td><td>{rel.iloc[i,1]}%</td></tr>"
    html = html.replace("{{tabla_seguridad}}", tabla)

    # ================== VICTIMIZACIÓN ==================
    html = html.replace(
        "{{grafico_victima_denuncia}}",
        grafico_barras(extraer_bloque(rel,12,14,0,1),"Victimización y denuncia")
    )

    html = html.replace(
        "{{grafico_no_denuncia}}",
        grafico_barras(extraer_bloque(rel,15,18,0,1),"Motivos de no denuncia")
    )

    # ================== DINÁMICA DELICTIVA ==================
    html = html.replace(
        "{{grafico_horario}}",
        grafico_pie(extraer_bloque(rel,20,24,0,1),"Horario delictivo")
    )

    html = html.replace(
        "{{grafico_armas}}",
        grafico_pie(extraer_bloque(rel,25,29,0,1),"Uso de armas")
    )

    # ================== SERVICIO POLICIAL ==================
    html = html.replace(
        "{{grafico_servicio_policial}}",
        grafico_barras(extraer_bloque(rel,31,35,0,1),"Percepción del servicio policial")
    )

    html = html.replace(
        "{{grafico_calificacion}}",
        grafico_pie(extraer_bloque(rel,36,40,0,1),"Calificación del servicio policial")
    )

    html = html.replace(
        "{{grafico_conoce_policias}}",
        grafico_pie(extraer_bloque(rel,41,43,0,1),"¿Conoce a los policías?")
    )

    html = html.replace(
        "{{grafico_conversado}}",
        grafico_pie(extraer_bloque(rel,44,46,0,1),"¿Ha conversado con ellos?")
    )

    # ================== SECTOR COMERCIAL ==================
    html = html.replace(
        "{{grafico_seguridad_local}}",
        grafico_pie(extraer_bloque(rel,48,50,0,1),"Seguridad local comercial")
    )

    html = html.replace(
        "{{grafico_conoce_programa}}",
        grafico_pie(extraer_bloque(rel,51,53,0,1),"Conocimiento del programa")
    )

    html = html.replace(
        "{{grafico_inscrito_programa}}",
        grafico_pie(extraer_bloque(rel,54,56,0,1),"Inscripción al programa")
    )

    html = html.replace(
        "{{grafico_contacto_programa}}",
        grafico_pie(extraer_bloque(rel,57,59,0,1),"Interés de contacto")
    )

# ================== INYECTAR CSS ==================
html = html.replace(
    "</head>",
    f"<style>{css}</style></head>"
)

# ================== PREVIEW ==================
st.components.v1.html(html, height=1200, scrolling=True)

# ================== DESCARGA ==================
st.download_button(
    "⬇️ Descargar HTML (listo para PDF)",
    html,
    "Informe.html",
    "text/html"
)

