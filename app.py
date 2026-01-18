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


def generar_infografia_datos(
    template_path,
    output_path,
    datos,
    config
):
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.axis("off")

    img = plt.imread(template_path)
    ax.imshow(img)

    for key, cfg in config.items():
        ax.text(
            cfg["x"],
            cfg["y"],
            f"{datos[key]:,}",
            transform=ax.transAxes,
            fontsize=cfg["fontsize"],
            color=cfg["color"],
            ha=cfg.get("ha", "center"),
            va=cfg.get("va", "center"),
            weight=cfg.get("weight", "bold")
        )

    fig.savefig(
        output_path,
        dpi=200,
        bbox_inches="tight",
        transparent=True
    )
    plt.close(fig)


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

        rel_labels = df.iloc[7:11, 6].astype(str)
        rel_base_values = pd.to_numeric(df.iloc[7:11, 8], errors="coerce")
        rel_percent_labels = (
            df.iloc[7:11, 7].astype(str).str.replace(",", ".", regex=False)
        )

        mask = rel_base_values.notna()
        rel_labels = rel_labels[mask]
        rel_base_values = rel_base_values[mask]
        rel_percent_labels = rel_percent_labels[mask]

        fig, ax = plt.subplots(figsize=(9, 6))
        ax.bar(rel_labels, rel_base_values, color="#30a907")
        ax.margins(y=0.1)

        for spine in ax.spines.values():
            spine.set_visible(False)

        ax.tick_params(left=False, bottom=False)
        ax.set_facecolor("none")
        fig.patch.set_alpha(0)

        for i in range(len(rel_percent_labels)):
            porcentaje = float(rel_percent_labels.iloc[i]) * 100
            ax.text(i, rel_base_values.iloc[i], f"{porcentaje:.2f}%", ha="center")

        buf_rel = BytesIO()
        fig.savefig(buf_rel, dpi=200, transparent=True)
        plt.close(fig)
        buf_rel.seek(0)

        grafico_rel_path = BASE_DIR / "grafico_relacion.png"
        with open(grafico_rel_path, "wb") as f:
            f.write(buf_rel.getbuffer())

        datos_fuentes = {
            "encuesta_comunidad": int(df.iloc[82, 1]),
            "encuesta_policial": int(df.iloc[83, 1]),
            "encuesta_comercio": int(df.iloc[84, 1]),
            "estadistica_registrada": int(df.iloc[86, 2]),
            "total_datos": int(df.iloc[87, 1])
        }

        config_infografia = {
            "encuesta_comunidad": {"x": 0.27, "y": 0.62, "fontsize": 20, "color": "white"},
            "encuesta_policial": {"x": 0.73, "y": 0.62, "fontsize": 20, "color": "white"},
            "encuesta_comercio": {"x": 0.27, "y": 0.35, "fontsize": 20, "color": "white"},
            "estadistica_registrada": {"x": 0.73, "y": 0.35, "fontsize": 20, "color": "white"},
            "total_datos": {"x": 0.5, "y": 0.48, "fontsize": 26, "color": "#333333"}
        }

        infografia_datos_path = BASE_DIR / "datos_render.png"

        generar_infografia_datos(
            template_path=str(ASSETS_DIR / "datos.png"),
            output_path=str(infografia_datos_path),
            datos=datos_fuentes,
            config=config_infografia
        )

        if not infografia_datos_path.exists():
            raise RuntimeError("NO SE GENERÓ datos_render.png")

        if st.button("HACER INFORME TERRITORIAL"):
            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_relacion_path=str(grafico_rel_path),
                grafico_edad_path="",
                grafico_escolaridad_path="",
                grafico_genero_path="",
                infografia_datos_path=str(infografia_datos_path),
                delegacion=delegacion,
                codigo=codigo,
                tabla_participacion=tabla_participacion,
                tabla_edad=[],
                tabla_escolaridad=[],
                tabla_genero=[],
                tabla_encuesta_comunidad=[],
                tabla_otras_encuestas=[]
            )

            st.download_button(
                "⬇️ Descargar PDF",
                data=pdf_buffer.getvalue(),
                file_name="informe.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
