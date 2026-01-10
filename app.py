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
        # Lectura del excel
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
        fig, ax = plt.subplots(figsize=(9, 6))

        ax.bar(rel_labels, rel_base_values, color="#30a907")
        ax.set_ylabel("Cantidad")
        ax.set_title("")

        # CAMBIO ÚNICO APLICADO
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
                fontsize=12
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

        # ================= PARTICIPACIÓN POR EDAD =================
        # Intervalos de edad (A29:A33)
        edad_labels = df.iloc[28:33, 0].astype(str)

        # Porcentajes (B29:B33)
        edad_percent_values = (
            df.iloc[28:33, 1]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        edad_percent_values = pd.to_numeric(edad_percent_values, errors="coerce")

        # Filtrar filas válidas
        mask = edad_percent_values.notna()
        edad_labels = edad_labels[mask]
        edad_percent_values = edad_percent_values[mask]

        fig_edad, ax_edad = plt.subplots(figsize=(7, 7))

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
            pctdistance=0.65,     # ⬅️ porcentajes más centrados
            labeldistance=1.15,  # ⬅️ etiquetas más cerca del círculo
            startangle=90,
            colors=colores,
            textprops={"fontsize": 20}
        )


        ax_edad.axis("equal")


        # Tamaño de porcentajes
        for autotext in autotexts:
            autotext.set_fontsize(25)

        # Fondo transparente
        ax_edad.set_facecolor("none")
        fig_edad.patch.set_alpha(0)

        buf_edad = BytesIO()
        fig_edad.savefig(
            buf_edad,
            format="png",
            dpi=200,
            transparent=True
        )

        plt.close(fig_edad)
        buf_edad.seek(0)

        plt.close(fig_edad)
        buf_edad.seek(0)

        grafico_edad_path = BASE_DIR / "grafico_participacion_edad.png"
        with open(grafico_edad_path, "wb") as f:
            f.write(buf_edad.getbuffer())


        ## TAbla EDAD==============S=S=S=S===========S=S=
        tabla_edad_df = df.iloc[28:33, 0:2].copy()
        tabla_edad_df.iloc[:, 1] = tabla_edad_df.iloc[:, 1].apply(lambda x: f"{x*100:.0f}%" if isinstance(x, (int, float)) else x)
        tabla_edad = tabla_edad_df.fillna("").values.tolist()

        #####________________________________BLOQUE ESCOLARIDAD___________________________________

        # Intervalos / niveles (A39:A46)
        escolaridad_labels = df.iloc[38:46, 0].astype(str)

        # Porcentajes (B39:B46)
        escolaridad_percent_values = (
            df.iloc[38:46, 1]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        escolaridad_percent_values = pd.to_numeric(escolaridad_percent_values, errors="coerce")

        # Filtrar filas válidas
        mask = escolaridad_percent_values.notna()
        escolaridad_labels = escolaridad_labels[mask]
        escolaridad_percent_values = escolaridad_percent_values[mask]

        # -------- GRÁFICO --------
        fig_esco, ax_esco = plt.subplots(figsize=(7, 7))

        colores_esco = [
            "#5B9BD5",
            "#A5A5A5",
            "#4472C4",
            "#255E91",
            "#B7B7B7",
            "#9DC3E6",
            "#8FAADC",
            "#D9E1F2"
        ]

        wedges, texts, autotexts = ax_esco.pie(
            escolaridad_percent_values,
            labels=None,
            autopct=lambda p: f"{p:.0f}%",
            pctdistance=0.65,
            labeldistance=1.15,
            startangle=90,
            colors=colores_esco,
            textprops={"fontsize": 20}
        )

        ax_esco.axis("equal")

        # Tamaño de porcentajes
        for autotext in autotexts:
            autotext.set_fontsize(25)

        # Fondo transparente
        ax_esco.set_facecolor("none")
        fig_esco.patch.set_alpha(0)

    
        buf_esco = BytesIO()
        fig_esco.savefig(
            buf_esco,
            format="png",
            dpi=200,
            transparent=True
        )

        plt.close(fig_esco)
        buf_edad.seek(0)

        grafico_escolaridad_path = BASE_DIR / "grafico_participacion_escolaridad.png"
        with open(grafico_escolaridad_path, "wb") as f:
            f.write(buf_esco.getbuffer())

        # -------- TABLA --------
        tabla_escolaridad_df = df.iloc[38:46, 0:2].copy()

        tabla_escolaridad_df.iloc[:, 1] = tabla_escolaridad_df.iloc[:, 1].apply(
            lambda x: f"{x*100:.0f}%" if isinstance(x, (int, float)) else x
        )

        tabla_escolaridad = tabla_escolaridad_df.fillna("").values.tolist()

        # ------------------------------------- Bloque de participación por género -----------------------------------------
        # Etiquetas (A52:A54)
        genero_labels = df.iloc[51:54, 0].astype(str)

        # Porcentajes (B52:B54)
        genero_percent_values = (
            df.iloc[51:54, 1]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )

        genero_percent_values = pd.to_numeric(genero_percent_values, errors="coerce")

        # Filtrar válidos
        mask = genero_percent_values.notna()
        genero_labels = genero_labels[mask]
        genero_percent_values = genero_percent_values[mask]

        # -------- GRÁFICO --------
        fig_gen, ax_gen = plt.subplots(figsize=(7, 7))

        # mismos colores que edad/escolaridad
        colores_genero = [
            "#5B9BD5",
            "#A5A5A5",
            "#4472C4"
        ]

        wedges, texts, autotexts = ax_gen.pie(
            genero_percent_values,
            labels=None,
            autopct=lambda p: f"{p:.0f}%",
            pctdistance=0.65,
            labeldistance=1.15,
            startangle=90,
            colors=colores_genero,
            textprops={"fontsize": 20}
        )

        ax_gen.axis("equal")

        for autotext in autotexts:
            autotext.set_fontsize(25)

        ax_gen.set_facecolor("none")
        fig_gen.patch.set_alpha(0)

        buf_gen = BytesIO()
        fig_gen.savefig(
            buf_gen,
            format="png",
            dpi=200,
            transparent=True
        )

        plt.close(fig_gen)
        buf_gen.seek(0)

        grafico_genero_path = BASE_DIR / "grafico_participacion_genero.png"
        with open(grafico_genero_path, "wb") as f:
            f.write(buf_gen.getbuffer())

        # -------- TABLA --------
        tabla_genero_df = df.iloc[51:54, 0:2].copy()
        tabla_genero_df.iloc[:, 1] = tabla_genero_df.iloc[:, 1].apply(
            lambda x: f"{x*100:.0f}%" if isinstance(x, (int, float)) else x
        )

        tabla_genero = tabla_genero_df.fillna("").values.tolist() 

        #______________________________________________________________________________________________________
        #______________________________________________________________________________________________________
        # ================= GENERAR PDF =================
        if st.button("HACER INFORME TERRITORIAL"):
            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_relacion_path=str(grafico_rel_path),
                grafico_edad_path=str(grafico_edad_path),
                grafico_escolaridad_path=str(grafico_escolaridad_path),
                grafico_genero_path=str(grafico_genero_path),   # 👈 NUEVO
                delegacion=delegacion,
                codigo=codigo,
                tabla_participacion=tabla_participacion,
                tabla_edad=tabla_edad,
                tabla_escolaridad=tabla_escolaridad,
                tabla_genero=tabla_genero                        # 👈 NUEVO
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

