import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path

from pdf_generator import generar_pdf

# ================= STREAMLIT =================
st.set_page_config(page_title="Generador de PDF", layout="centered")
st.title("Generador de PDF – Base Estable")

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
        # 🔹 Leer NUEVA matriz (ya procesada)
        df = pd.read_excel(
            archivo,
            sheet_name="Hoja1",
            header=None,
            engine="openpyxl"
        )

        # 🔹 Tomar valores YA CALCULADOS en la matriz
        # Ajustá SOLO estos índices si cambian
        labels = ["Comunidad", "Comercio", "Fuerza Pública"]
        values = df.iloc[180:183, 5]  # ejemplo: F181:F183

        labels, values = limpiar_series(labels, values)

        grafico_buffer = crear_grafico(labels, values)

        if st.button("Generar PDF"):
            # Guardar gráfico temporal
            grafico_path = BASE_DIR / "grafico_temp.png"
            with open(grafico_path, "wb") as f:
                f.write(grafico_buffer.getbuffer())

            # Generar PDF
            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_path=str(grafico_path)
            )

            st.download_button(
                label="⬇️ Descargar PDF",
                data=pdf_buffer,
                file_name="informe.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
