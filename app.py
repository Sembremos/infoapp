import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path

from pdf_generator import generar_pdf

# ================= STREAMLIT =================
st.set_page_config(page_title="Generador de Informe", layout="centered")
st.title("Generador de Informe – Vista previa")

# ================= RUTAS =================
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

# ================= UTILIDADES =================
def limpiar_series(labels, values):
    df = pd.DataFrame({"label": labels, "value": values})
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna()
    return df["label"].astype(str), df["value"]

def crear_grafico(labels, values, preview=False):
    fig, ax = plt.subplots()
    ax.bar(labels, values)
    ax.set_ylim(0, 100)
    ax.set_ylabel("Porcentaje")
    plt.xticks(rotation=20)

    if preview:
        st.pyplot(fig)

    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf

# ================= APP =================
archivo = st.file_uploader("Subir matriz Excel", type=["xlsx"])

if archivo:
    try:
        # Leer matriz ya procesada
        df = pd.read_excel(
            archivo,
            sheet_name="Hoja1",
            header=None,
            engine="openpyxl"
        )

        # ====== DATOS DE EJEMPLO (ajustables) ======
        titulo = df.iloc[4, 1]           # B5
        valores = df.iloc[180:183, 5]    # F181:F183
        labels = ["Comunidad", "Comercio", "Fuerza Pública"]

        labels, values = limpiar_series(labels, valores)

        # ================= VISTA PREVIA =================
        st.subheader("Vista previa del informe")

        st.markdown(f"### {titulo}")

        grafico_buffer = crear_grafico(labels, values, preview=True)

        # ================= PDF =================
        if st.button("Generar PDF"):
            grafico_path = BASE_DIR / "grafico_temp.png"
            with open(grafico_path, "wb") as f:
                f.write(grafico_buffer.getbuffer())

            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_path=str(grafico_path)
            )

            st.download_button(
                "⬇️ Descargar PDF",
                pdf_buffer,
                "informe.pdf",
                "application/pdf"
            )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")

