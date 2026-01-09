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

# ================= CONSTANTES DE GRÁFICOS =================
FIG_RELACION_SIZE = (9, 6)
FIG_EDAD_SIZE = (7, 7)

FONT_RELACION_LABEL = 12
FONT_EDAD_LABEL = 20
FONT_EDAD_PERCENT = 25

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
        wb = load_workbook(archivo, data_only=True)
        ws = wb["Hoja1"]
        df = pd.DataFrame(ws.values)

        delegacion = str(df.iloc[1, 1])
        codigo = str(df.iloc[2, 1])

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
        rel_labels = df.iloc[7:11, 6].astype(str)
        rel_base_values = pd.to_numeric(df.iloc[7:11, 8], errors="coerce")
        rel_percent_labels = (
            df.iloc[7:11, 7]
            .astype(str)
            .str.replace(",", ".", regex=False)
        )

        mask = rel_base_values.notna()
        rel_labels = rel_labels[mask]
        rel_base_values = rel_base_values[mask]
        rel_percent_labels = rel_percent_labels[mask]

        fig, ax = plt.subplots(figsize=FIG_RELACION_SIZE)
        ax.bar(rel_labels, rel_base_values, color="#30a907")
        ax.set_ylabel("Cantidad")
        ax.set_title("")
        ax.margins(y=0.1)

        for spine in ax.spines.values():
            spine.set_visible(False)

        ax.tick_params(left=False, bottom=False)
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
                fontsize=FONT_RELACION_LABEL
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

        # ================= GRÁFICO EDAD =================
        edad_labels = df.iloc[28:33, 0].astype(str)
        edad_percent_values = (
            df.iloc[28:33, 1]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        edad_percent_values = pd.to_numeric(edad_percent_values, errors="coerce")
        mask = edad_percent_values.notna()
        edad_labels = edad_labels[mask]
        edad_percent_values = edad_percent_values[mask]

        fig_edad, ax_edad = plt.subplots(figsize=FIG_EDAD_SIZE)

        colores = [
            "#5B9BD5",
            "#A5A5A5",
            "#4472C4",
            "#255E91",
            "#B7B7B7"
        ]

        wedges, texts, autotexts = ax_edad.pie(
            edad_percent_values,
            labels=edad_labels,
            autopct=lambda p: f"{p:.0f}%",
            pctdistance=0.6,
            labeldistance=1.35,
            startangle=90,
            colors=colores,
            textprops={"fontsize": FONT_EDAD_LABEL}
        )

        ax_edad.axis("equal")

        for text in texts:
            text.set_fontsize(FONT_EDAD_LABEL)

        for autotext in autotexts:
            autotext.set_fontsize(FONT_EDAD_PERCENT)

        ax_edad.set_facecolor("none")
        fig_edad.patch.set_alpha(0)

        buf_edad = BytesIO()
        fig_edad.savefig(
            buf_edad,
            format="png",
            bbox_inches="tight",
            pad_inches=0,
            dpi=200,
            transparent=True
        )

        plt.close(fig_edad)
        buf_edad.seek(0)

        grafico_edad_path = BASE_DIR / "grafico_participacion_edad.png"
        with open(grafico_edad_path, "wb") as f:
            f.write(buf_edad.getbuffer())

        if st.button("HACER INFORME TERRITORIAL"):
            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_relacion_path=str(grafico_rel_path),
                grafico_edad_path=str(grafico_edad_path),
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
