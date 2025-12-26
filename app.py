import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import base64
from io import BytesIO
from pathlib import Path

st.set_page_config(page_title="Generador de Informes", layout="wide")
st.title("Generador de Informes")

def img_to_base64(path):
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()

def fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=200)
    plt.close(fig)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()

def grafico_pie(labels, values, title):
    df = pd.DataFrame({"l": labels, "v": values})
    df["v"] = pd.to_numeric(df["v"], errors="coerce")
    df = df.dropna()
    if df.empty:
        return img_to_base64(Path("final.png"))
    fig, ax = plt.subplots()
    ax.pie(df["v"], labels=df["l"], autopct="%1.0f%%", startangle=90)
    ax.set_title(title)
    return fig_to_base64(fig)

def grafico_barras(labels, values, title):
    df = pd.DataFrame({"l": labels, "v": values})
    df["v"] = pd.to_numeric(df["v"], errors="coerce")
    df = df.dropna()
    if df.empty:
        return img_to_base64(Path("final.png"))
    fig, ax = plt.subplots()
    ax.bar(df["l"], df["v"])
    ax.set_ylim(0, 100)
    ax.set_title(title)
    return fig_to_base64(fig)

html = Path("Informe.html").read_text(encoding="utf-8")
css = Path("estilos.css").read_text(encoding="utf-8")

imagenes = [
    "portada.png","intro.png","conformacion.png","participacion.png",
    "metodologico.png","estadistica.png","lineas.png","percepcion.png",
    "final.png","header.png","footer.png"
]

for img in imagenes:
    p = Path(img)
    if p.exists():
        html = html.replace(f'src="{img}"', f'src="{img_to_base64(p)}"')

archivo_excel = st.file_uploader("Subir matriz Excel", type=["xlsx"])

if archivo_excel:
    part = pd.read_excel(archivo_excel, sheet_name="participacion", header=None)
    rel = pd.read_excel(archivo_excel, sheet_name="relevante", header=None)

    html = html.replace("{{tabla_distritos}}", f"""
      <tr><td>{part.iloc[33,1]}</td><td>{part.iloc[33,2]*100:.0f}%</td><td>{int(part.iloc[33,3])}</td></tr>
    """)

    html = html.replace("{{grafico_relacion}}", grafico_barras(["Comunidad","Comercio","FP"], part.iloc[33,4:7]*100, "Relación"))
    html = html.replace("{{grafico_edad}}", grafico_pie(part.iloc[1:6,1], part.iloc[1:6,2]*100, "Edad"))
    html = html.replace("{{grafico_escolaridad}}", grafico_pie(part.iloc[11:18,1], part.iloc[11:18,2]*100, "Escolaridad"))
    html = html.replace("{{grafico_genero}}", grafico_pie(part.iloc[23:26,1], part.iloc[23:26,2]*100, "Género"))

html = html.replace("</head>", f"<style>{css}</style></head>")

st.components.v1.html(html, height=1000, scrolling=True)
st.download_button("⬇️ Descargar HTML", html, "Informe.html", "text/html")
