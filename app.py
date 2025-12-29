import streamlit as st
import pandas as pd
from pathlib import Path
import base64

from pdf_generator import generar_pdf

# ================= STREAMLIT =================
st.set_page_config(page_title="Generador de Informe", layout="wide")
st.title("Generador de Informe – Vista previa PDF")

# ================= RUTAS =================
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

# ================= APP =================
archivo = st.file_uploader("Subir matriz Excel", type=["xlsx"])

if archivo:
    try:
        # Leer matriz (ya procesada)
        df = pd.read_excel(
            archivo,
            sheet_name="Hoja1",
            header=None,
            engine="openpyxl"
        )

        # 👉 Generar PDF en memoria
        pdf_buffer = generar_pdf(
            portada_path=str(ASSETS_DIR / "portada.png"),
            data=df  # se pasa el dataframe completo
        )

        # ================= VISTA PREVIA PDF =================
        st.subheader("Vista previa del informe")

        pdf_bytes = pdf_buffer.getvalue()
        base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

        pdf_display = f"""
        <iframe
            src="data:application/pdf;base64,{base64_pdf}"
            width="100%"
            height="900px"
            type="application/pdf">
        </iframe>
        """

        st.components.v1.html(pdf_display, height=900, scrolling=True)

        # ================= DESCARGA =================
        st.download_button(
            "⬇️ Descargar PDF",
            data=pdf_bytes,
            file_name="informe.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
