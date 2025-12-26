import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Image, Spacer
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

# ================= STREAMLIT =================
st.set_page_config(page_title="Generador de PDF", layout="centered")
st.title("Generador de PDF (Versión Base)")

# ================= RUTAS =================
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

# ================= UTILIDADES =================
def limpiar_series(labels, values):
    df = pd.DataFrame({"label": labels, "value": values})
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna()
    return df["label"].astype(str), df["value"]

def crear_grafico(labels, values):
    fig, ax = plt.subplots()
    ax.bar(labels, values)
    ax.set_ylim(0, 100)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf

def generar_pdf(ruta_pdf, grafico):
    styles = getSampleStyleSheet()

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=A4
    )

    story = []

    # ===== PORTADA =====
    story.append(
        Image(
            str(ASSETS_DIR / "portada.png"),
            width=A4[0],
            height=A4[1]
        )
    )

    # ===== CONTENIDO =====
    story.append(Spacer(1, 20))
    story.append(Paragraph("Datos de participación", styles["Heading1"]))
    story.append(Spacer(1, 12))
    story.append(Image(grafico, width=400, height=300))

    doc.build(story)

# ================= APP =================
archivo = st.file_uploader("Subir matriz Excel", type=["xlsx"])

if archivo:
    try:
        df = pd.read_excel(
            archivo,
            sheet_name="participacion",
            header=None,
            engine="openpyxl"
        )

        labels, values = limpiar_series(
            ["Comunidad", "Comercio", "Fuerza Pública"],
            df.iloc[33, 4:7] * 100
        )

        grafico = crear_grafico(labels, values)

        if st.button("Generar PDF"):
            ruta_pdf = BASE_DIR / "informe_base.pdf"
            generar_pdf(ruta_pdf, grafico)

            with open(ruta_pdf, "rb") as f:
                st.download_button(
                    "⬇️ Descargar PDF",
                    f,
                    file_name="informe_base.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"Error: {e}")
