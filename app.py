import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pdf_generator import generar_pdf

st.set_page_config(page_title="Generador de Informes", layout="centered")
st.title("Generador de Informes PDF")

archivo = st.file_uploader("Subir matriz Excel", type=["xlsx"])

def crear_grafico_barras(labels, values, titulo):
    fig, ax = plt.subplots()
    ax.bar(labels, values)
    ax.set_title(titulo)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf

if archivo:
    part = pd.read_excel(archivo, sheet_name="participacion", header=None)
    rel = pd.read_excel(archivo, sheet_name="relevante", header=None)

    graficos = {
        "relacion": crear_grafico_barras(
            ["Comunidad", "Comercio", "FP"],
            part.iloc[33,4:7] * 100,
            "Relación de participación"
        ),
        "edad": crear_grafico_barras(
            part.iloc[1:6,1],
            part.iloc[1:6,2] * 100,
            "Participación por edad"
        ),
        "seguridad": crear_grafico_barras(
            rel.iloc[1:4,0],
            rel.iloc[1:4,1],
            "Seguridad en la comunidad"
        )
    }

    data = {
        "introduccion": "Texto institucional de introducción..."
    }

    if st.button("Generar PDF"):
        ruta_pdf = "Informe_Territorial.pdf"
        generar_pdf(ruta_pdf, data, graficos)

        with open(ruta_pdf, "rb") as f:
            st.download_button(
                "⬇️ Descargar PDF",
                f,
                file_name="Informe_Territorial.pdf",
                mime="application/pdf"
            )
