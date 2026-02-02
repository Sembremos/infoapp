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
            labels=None,
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
        buf_esco.seek(0)

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

        #__________________-----------------------_______________TAblas de encuestas
        #ENCUESTAS COMUNIDAD
        tabla_encuesta_comunidad_df = df.iloc[58:60, 0:4].copy()
        tabla_encuesta_comunidad = tabla_encuesta_comunidad_df.fillna("").values.tolist()

        #OTRAS ENCUESTAS
        tabla_otras_encuestas_df = df.iloc[62:65, 6:10].copy()
        tabla_otras_encuestas = tabla_otras_encuestas_df.fillna("").values.tolist()

        ## imagen datos
        datos_pagina_8 = {
            "encuesta_comunidad": int(df.iloc[82, 1]),   # B83
            "encuesta_policial": int(df.iloc[83, 1]),    # B84
            "encuesta_comercio": int(df.iloc[84, 1]),    # B85
            "estadistica": int(df.iloc[86, 2]),           # C86
            "total_datos": int(df.iloc[87, 1])            # B88
        }

        ##==============================PARETO===================================
        datos_pagina_9 = {
            "lado_izquierdo": str(df.iloc[92, 0]),   # A93
            "derecha_superior": str(df.iloc[92, 1]), # B93
            "derecha_inferior": str(df.iloc[92, 2])  # C93
        }
        #=========================LISTAS PERETO================================
        # ================= DELITOS =================
        tabla_delitos_raw = df.iloc[96:117, 1]  # B97:B117

        tabla_delitos = [
            [str(v)]
            for v in tabla_delitos_raw
            if pd.notna(v) and str(v) != "0"
        ]

        # =================RIESGOS SOCIALES =================
        tabla_riesgos_raw = df.iloc[96:117, 2]  # C97:C117

        tabla_riesgos = [
            [str(v)]
            for v in tabla_riesgos_raw
            if pd.notna(v) and str(v) != "0"
        ]

        # ================= PORCENTAJES PARETO =================
        porcentaje_delitos = f"{float(df.iloc[118, 1]) * 100:.2f}%"
        porcentaje_riesgos = f"{float(df.iloc[118, 2]) * 100:.2f}%"


        # ================= CANTIDAD DELITOS =================
        cantidad_delitos = int(df.iloc[117, 1])  # B118
        #==================cantidad riesgos=======
        cantidad_riesgos = int(df.iloc[117, 2])

       # ================= MICMAC =================

        def limpiar_lista(col):
            return [
                [str(v)]
                for v in col
                if pd.notna(v) and str(v).strip() != ""
            ]
        
        # Poder
        micmac_poder = limpiar_lista(df.iloc[123:140, 1])      # B124:B125
        # Conflicto
        micmac_conflicto = limpiar_lista(df.iloc[123:140, 2])  # C124:C125
        # Autónomas
        micmac_autonomas = limpiar_lista(df.iloc[123:140, 3])  # D124:D125
        # Resultados
        micmac_resultados = limpiar_lista(df.iloc[123:140, 4]) # E124:E125

        #_________________________micmac2_______________________________
        
        def limpiar_lista_simple(col):
            return [
                [str(v)]
                for v in col
                if pd.notna(v) and str(v).strip() != ""
            ]
        
        # ===== TABLAS MICMAC 2 =====
        tabla_riesgos_micmac2 = limpiar_lista_simple(df.iloc[123:140, 10])  # K124:K140
        tabla_delitos_micmac2 = limpiar_lista_simple(df.iloc[123:140, 11])  # L124:L140
        
        # ===== DATOS SUELTOS =====
        cantidad_problematicas = int(df.iloc[140, 12])  # M141
        riesgos_total = int(df.iloc[140, 10])           # K141
        delitos_total = int(df.iloc[140, 11])           # L141

        #_________________________Triangulo de violes______________________
        
        causas_identificadas = int(df.iloc[117, 3])   # D118
        factores_micmac = int(df.iloc[140, 12])       # M141
        
        triangulo_directa = int(df.iloc[146, 0])      # A147
        triangulo_sociocultural = int(df.iloc[146, 1])# B147
        triangulo_estructural = int(df.iloc[146, 2])  # C147

        #______________________LISTA DE INSTIS_____-------_-------___-
        
        tabla_instituciones_df = df.iloc[149:159, 1:3].copy()  # B150:C160
        
        # Eliminar filas completamente vacías
        tabla_instituciones_df = tabla_instituciones_df.dropna(how="all")
        
        tabla_instituciones = []
        for _, row in tabla_instituciones_df.iterrows():
            col1 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            col2 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        
            if col1 or col2:  # al menos una columna con contenido
                tabla_instituciones.append([col1, col2])
        
        
        #_______________________________PAGINA ESTADISTICA_______________________________
        df_denuncias = pd.read_excel(
            "datos.xlsx",
            sheet_name="Hoja1",
            usecols="A,C",
            skiprows=165,
            nrows=11
        )
        
        df_denuncias.columns = ["categoria", "porcentaje"]
        
        tabla_denuncias = [[row["categoria"]] for _, row in df_denuncias.iterrows()]

        #_____Tabla$$$$$$$$$$$$$$$$$$$$$$$
        total_denuncias = pd.read_excel(
            "datos.xlsx",
            sheet_name="Hoja1",
            usecols="B",
            skiprows=176,
            nrows=1
        ).iloc[0, 0]

        # ================== PROCESAMIENTO DENUNCIAS ==================

        # Tabla para el PDF 
        tabla_denuncias = [[row["categoria"]] for _, row in df_denuncias.iterrows()]
        
        # ================== GRÁFICO CIRCULAR DENUNCIAS ==================
        
        def generar_grafico_denuncias(df):
            colores = [
                "#4472C4", "#5B9BD5", "#A5A5A5", "#70AD47",
                "#255E91", "#9DC3E6", "#264478", "#B7B7B7",
                "#30A907", "#8FAADC", "#D9E1F2"
            ]
        
            fig, ax = plt.subplots(figsize=(6, 6))
        
            ax.pie(
                df["porcentaje"],
                labels=df["categoria"],
                autopct="%1.0f%%",
                startangle=90,
                colors=colores[:len(df)]
            )
        
            ax.axis("equal")
            plt.tight_layout()
            plt.savefig("assets/grafico_denuncias.png", dpi=300)
            plt.close()
        
        
        # execute
        generar_grafico_denuncias(df_denuncias)
        #______________________________________________________________________________________________________
        # ================= GENERAR PDF =================
        if st.button("HACER INFORME TERRITORIAL"):
            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                grafico_relacion_path=str(grafico_rel_path),
                grafico_edad_path=str(grafico_edad_path),
                grafico_escolaridad_path=str(grafico_escolaridad_path),
                grafico_genero_path=str(grafico_genero_path),
                delegacion=delegacion,
                codigo=codigo,
                tabla_participacion=tabla_participacion,
                tabla_edad=tabla_edad,
                tabla_escolaridad=tabla_escolaridad,
                tabla_genero=tabla_genero,
                tabla_encuesta_comunidad=tabla_encuesta_comunidad,
                tabla_otras_encuestas=tabla_otras_encuestas,
                datos_pagina_8=datos_pagina_8,
                datos_pagina_9=datos_pagina_9,
                tabla_delitos=tabla_delitos,
                tabla_riesgos=tabla_riesgos,
                porcentaje_delitos=porcentaje_delitos,
                porcentaje_riesgos=porcentaje_riesgos,
                cantidad_delitos=cantidad_delitos,
                cantidad_riesgos=cantidad_riesgos,
                micmac_poder=micmac_poder,
                micmac_conflicto=micmac_conflicto,
                micmac_autonomas=micmac_autonomas,
                micmac_resultados=micmac_resultados,
                tabla_riesgos_micmac2=tabla_riesgos_micmac2,
                tabla_delitos_micmac2=tabla_delitos_micmac2,
                cantidad_problematicas=cantidad_problematicas,
                riesgos_total=riesgos_total,
                delitos_total=delitos_total,
                causas_identificadas=causas_identificadas,
                factores_micmac=factores_micmac,
                triangulo_directa=triangulo_directa,
                triangulo_sociocultural=triangulo_sociocultural,
                triangulo_estructural=triangulo_estructural,
                tabla_instituciones=tabla_instituciones,
                grafico_denuncias_path="assets/grafico_denuncias.png",
                tabla_denuncias=tabla_denuncias,
                total_denuncias=total_denuncias
            )

            pdf_bytes = pdf_buffer.getvalue()

            st.subheader("Informe generado correctamente")

            st.download_button(
                label="⬇️ Descargar PDF",
                data=pdf_bytes,
                file_name="informe.pdf",
                mime="application/pdf"
            )
    
    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")

