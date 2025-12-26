import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path

from pdf_generator import generar_pdf

# ================= STREAMLIT =================
st.set_page_config(page_title="Generador de PDF", layout="centered")
st.title("Generador de PDF – Versión Estable")

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
    plt.xticks(rotation=20)

    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf

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
            pdf_path = BASE_DIR / "informe.pdf"
            grafico_path = BASE_DIR / "grafico.png"

            # guardar gráfico temporal
            with open(grafico_path, "wb") as f:
                f.write(grafico.getbuffer())

            generar_pdf(
                ruta_pdf=str(pdf_path),
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_path=str(grafico_path)
            )

            with open(pdf_path, "rb") as f:
                st.download_button(
                    "⬇️ Descargar PDF",
                    f,
                    file_name="informe.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"Error: {e}")
