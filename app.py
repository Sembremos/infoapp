import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path
import base64
from openpyxl import load_workbook

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

    df["value"] = (
        df["value"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
    )

    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna()

    return df["label"].astype(str), df["value"]


def crear_grafico(labels, values):
    fig, ax = plt.subplots()
    ax.bar(labels, values)
    ax.set_ylim(0, 100)
    ax.set_ylabel("%")

    for i, v in enumerate(values):
        ax.text(i, v + 1, f"{v:.0f}%", ha="center", fontsize=9)

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
        # =========================================================
        # LECTURA EXACTA DEL EXCEL (valores reales, no fórmulas)
        # =========================================================
        wb = load_workbook(archivo, data_only=True)
        ws = wb["Hoja1"]
        df = pd.DataFrame(ws.values)

        # ================= DATOS BASE =================
        delegacion = str(df.iloc[1, 1])  # B2
        codigo = str(df.iloc[2, 1])      # B3

        # ================= TABLA PARTICIPACIÓN =================
        tabla_df = df.iloc[6:23, 0:3].dropna(how="all")

        tabla_df = tabla_df[
            ~(tabla_df.iloc[:, 1:].fillna(0).astype(str) == "0").all(axis=1)
        ]

        def formatear(v):
            if isinstance(v, (int, float)):
                if 0 <= v <= 1:
                    return f"{v*100:.0f}%"
                return f"{v:.0f}"
            return str(v)

        tabla_participacion = [
            [formatear(c) for c in fila]
            for fila in tabla_df.fillna("").values.tolist()
        ]

        # ================= GRÁFICO RELACIÓN =================
        # Nombres (G8:G11)
        rel_labels = df.iloc[7:11, 6].astype(str)

        # Valores reales base (I8:I11)
        rel_base_values = pd.to_numeric(df.iloc[7:11, 8], errors="coerce")

        # Porcentajes literales visibles (H8:H11)
        rel_percent_labels = (
            df.iloc[7:11, 7]
            .astype(str)
            .str.replace(",", ".", regex=False)
        )

        # Filtrar filas válidas
        mask = rel_base_values.notna()
        rel_labels = rel_labels[mask]
        rel_base_values = rel_base_values[mask]
        rel_percent_labels = rel_percent_labels[mask]

        # Crear gráfico
        fig, ax = plt.subplots(figsize=(9, 4))

        ax.bar(rel_labels, rel_base_values, color="#30a907")

        ax.set_ylabel("Cantidad")
        ax.set_title("")

        # >>> CAMBIO ÚNICO APLICADO <<<
        ax.margins(y=0.1)

        # Borrar marcos
        for spine in ax.spines.values():
            spine.set_visible(False)

        # Quitar líneas de ticks
        ax.tick_params(left=False, bottom=False)

        # Fondo transparente (clave para PDF)
        ax.set_facecolor("none")
        fig.patch.set_alpha(0)

        for i in range(len(rel_percent_labels)):
            porcentaje = float(rel_percent_labels.iloc[i]) * 100
            ax.text(
                i,
                rel_base_values.iloc[i],
                f"{porcentaje:.2f}%",
                ha="center",
                va="bottom",
                fontsize=9
            )

        buf_rel = BytesIO()
        fig.savefig(
            buf_rel,
            format="png",
            bbox_inches="tight",
            pad_inches=0,
            dpi=200,
            transparent=True
        )

        plt.close(fig)
        buf_rel.seek(0)

        grafico_rel_path = BASE_DIR / "grafico_relacion.png"
        with open(grafico_rel_path, "wb") as f:
            f.write(buf_rel.getbuffer())

        # ================= GENERAR PDF =================
        if st.button("HACER INFORME TERRITORIAL"):
            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_path=str(grafico_rel_path),
                grafico_relacion_path=str(grafico_rel_path),
                delegacion=delegacion,
                codigo=codigo,
                tabla_participacion=tabla_participacion
            )

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

