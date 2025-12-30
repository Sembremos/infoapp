import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path
import base64

from pdf_generator import generar_pdf

# ================= STREAMLIT =================
st.set_page_config(page_title="Generador de PDF", layout="centered")
st.title("Generador de Informes SS")

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
archivo = st.file_uploader(
    "Aquí suba o arrastre el ENGINE de la Delegación correspondiente",
    type=["xlsx"]
)

if archivo:
    try:
        # 🔹 LECTURA DEL EXCEL
        df = pd.read_excel(
            archivo,
            sheet_name="Hoja1",
            header=None,
            engine="openpyxl"
        )

        # ================= DATOS BASE =================
        delegacion = str(df.iloc[1, 1])  # Hoja1!B2
        codigo = str(df.iloc[2, 1])      # Hoja1!B3

        # ================= GRÁFICO =================
        labels = ["Comunidad", "Comercio", "Fuerza Pública"]
        values = df.iloc[180:183, 5]

        labels, values = limpiar_series(labels, values)
        grafico_buffer = crear_grafico(labels, values)

        # ================= TABLA PARTICIPACIÓN POR DISTRITO =================
        tabla_df = df.iloc[6:23, 0:3]
        tabla_df = tabla_df.dropna(how="all")

        # eliminar filas donde todos los valores numéricos son 0
        tabla_df = tabla_df[
            ~(tabla_df.iloc[:, 1:].fillna(0) == 0).all(axis=1)
        ]

        def formatear(valor):
            if isinstance(valor, (int, float)):
                if 0 <= valor <= 1:
                    return f"{valor*100:.0f}%"
                return f"{valor:.0f}"
            return str(valor)

        tabla_participacion = [
            [formatear(celda) for celda in fila]
            for fila in tabla_df.fillna("").values.tolist()
        ]

        # ================= GENERAR PDF =================
        if st.button("Generar PDF"):
            grafico_path = BASE_DIR / "grafico_temp.png"
            with open(grafico_path, "wb") as f:
                f.write(grafico_buffer.getbuffer())

            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_path=str(grafico_path),
                delegacion=delegacion,
                codigo=codigo,
                tabla_participacion=tabla_participacion
            )

            # ================= VISTA PREVIA =================
            pdf_bytes = pdf_buffer.getvalue()
            base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

            st.subheader("Vista previa del informe")
            st.components.v1.html(
                f"""
                <iframe
                    src="data:application/pdf;base64,{base64_pdf}"
                    width="100%"
                    height="800px">
                </iframe>
                """,
                height=800,
                scrolling=True
            )

            st.download_button(
                label="⬇️ Descargar PDF",
                data=pdf_bytes,
                file_name="informe.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
