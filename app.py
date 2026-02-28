

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

#####=============valores vacios

def seguro_int(valor, default=0):
    if pd.isna(valor):
        return default
    try:
        return int(float(valor))
    except:
        return default

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
        # ================= CORRESPONSABLES DESDE EXCEL =================

        columnas_corresponsable = [
            "J", "P", "V", "AB", "AH", "AN",
            "AT", "AZ", "BF", "BL", "BR", "BX"
        ]
        
        corresponsables_excel = []
        
        for col in columnas_corresponsable:
            valor = ws[f"{col}246"].value
        
            if valor:
                corresponsables_excel.append(str(valor).strip())
            else:
                corresponsables_excel.append("")
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
            "encuesta_comunidad": seguro_int(df.iloc[82, 1]),   # B83
            "encuesta_policial": seguro_int(df.iloc[83, 1]),    # B84
            "encuesta_comercio": seguro_int(df.iloc[84, 1]),    # B85
            "estadistica": seguro_int(df.iloc[86, 2]),          # C86
            "total_datos": seguro_int(df.iloc[87, 1])           # B88
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
        
        df_grafico_denuncias = df.iloc[165:176, [0, 2]].copy()
        df_grafico_denuncias.columns = ["categoria", "porcentaje"]
        
        df_grafico_denuncias["porcentaje"] = (
            df_grafico_denuncias["porcentaje"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        
        df_grafico_denuncias["porcentaje"] = pd.to_numeric(
            df_grafico_denuncias["porcentaje"],
            errors="coerce"
        )
        
        df_grafico_denuncias = df_grafico_denuncias.dropna()
        
        # ----- DATOS PARA TABLA (A y B) -----
        df_tabla_denuncias = df.iloc[165:176, [0, 1]].copy()
        df_tabla_denuncias.columns = ["categoria", "cantidad"]
        
        df_tabla_denuncias = df_tabla_denuncias.dropna(how="all")
        
        tabla_denuncias = [
            [str(row["categoria"]), str(int(row["cantidad"]))]
            for _, row in df_tabla_denuncias.iterrows()
            if pd.notna(row["categoria"]) and pd.notna(row["cantidad"])
        ]
        
        # ----- TOTAL DENUNCIAS (B177) -----
        total_denuncias = int(df.iloc[176, 1])

       
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
            plt.savefig(ASSETS_DIR / "grafico_denuncias.png", dpi=300)
            plt.close()
        
        
        # execute
        generar_grafico_denuncias(df_grafico_denuncias)


        #====================================HORARIOS DE DELITOS=============================================

        # ----- GRAFICO PASTEL (A y C) -----
        df_grafico_horario = df.iloc[179:188, [0, 2]].copy()
        df_grafico_horario.columns = ["horario", "porcentaje"]
        
        df_grafico_horario["porcentaje"] = (
            df_grafico_horario["porcentaje"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        
        df_grafico_horario["porcentaje"] = pd.to_numeric(
            df_grafico_horario["porcentaje"],
            errors="coerce"
        )
        
        df_grafico_horario = df_grafico_horario.dropna()

        # ----- TABLA GRAFICO (A y B) -----
        df_tabla_horario = df.iloc[179:188, [0, 1]].copy()
        df_tabla_horario.columns = ["horario", "cantidad"]
        
        tabla_horario = [
            [str(row["horario"]), str(row["cantidad"])]
            for _, row in df_tabla_horario.iterrows()
            if pd.notna(row["horario"])
        ]

        # ----- CUADROS AM / PM -----
        total_am = df.iloc[179, 3]  # D180
        total_pm = df.iloc[179, 4]  # E180
        
        def formatear_porcentaje(valor):
            if isinstance(valor, (int, float)):
                return f"{valor * 100:.2f}%"
            return str(valor)
        
        total_am = formatear_porcentaje(total_am)
        total_pm = formatear_porcentaje(total_pm)

        # ----- TABLA GRANDE POR DISTRITO (FORMATEADA CORRECTAMENTE) -----

        # 1️⃣ Tomar encabezados (fila 179)
        encabezados = df.iloc[178, 0:17].copy()
        
        # 2️⃣ Tomar datos (A180 a Q188)
        tabla_datos = df.iloc[179:188, 0:17].copy()
        
        # 3️⃣ Eliminar columnas D y E (índices 3 y 4)
        columnas_a_eliminar = [3, 4]
        encabezados = encabezados.drop(encabezados.index[columnas_a_eliminar])
        tabla_datos = tabla_datos.drop(tabla_datos.columns[columnas_a_eliminar], axis=1)
        
        # 4️⃣ Formatear porcentajes columna C (ahora índice 2)
        def formatear_porcentaje(valor):
            if pd.notna(valor):
                try:
                    return f"{float(valor) * 100:.2f}%"
                except:
                    return valor
            return ""
        
        tabla_datos.iloc[:, 2] = tabla_datos.iloc[:, 2].apply(formatear_porcentaje)
        
        # 5️⃣ Respetar celdas vacías pero eliminar columnas completamente vacías
        tabla_datos = tabla_datos.loc[:, ~(tabla_datos.isna().all())]
        
        # 6️⃣ Unir encabezados + datos
        tabla_horario_distrito = [
            encabezados.loc[tabla_datos.columns].fillna("").astype(str).tolist()
        ]
        
        tabla_horario_distrito += tabla_datos.fillna("").astype(str).values.tolist()

        #-----------------------------GRAFICO-------------
        def generar_grafico_horario(df):
            colores = [
                "#4472C4", "#5B9BD5", "#A5A5A5", "#70AD47",
                "#255E91", "#9DC3E6", "#264478", "#B7B7B7",
                "#30A907"
            ]
        
            fig, ax = plt.subplots(figsize=(6, 6))
        
            ax.pie(
                df["porcentaje"],
                labels=df["horario"],
                autopct="%1.0f%%",
                startangle=90,
                colors=colores[:len(df)]
            )
        
            ax.axis("equal")
            plt.tight_layout()
            plt.savefig(ASSETS_DIR / "grafico_horario.png", dpi=300, transparent=True)
            plt.close()

        generar_grafico_horario(df_grafico_horario)

        # =========================================
        # GRAFICO BARRAS PAGINA 14
        # =========================================
        
        # Datos desde A196:B204
        df_grafico_p14 = df.iloc[195:204, [0, 1]].copy()
        df_grafico_p14.columns = ["categoria", "valor"]
        
        df_grafico_p14 = df_grafico_p14.dropna()
        
        df_grafico_p14["valor"] = pd.to_numeric(
            df_grafico_p14["valor"],
            errors="coerce"
        )
        
        df_grafico_p14 = df_grafico_p14.dropna()
        
        def generar_grafico_p14(df):
        
            # ===== VARIABLES CONFIGURABLES =====
            COLOR_BARRAS = "#30A907"      # Verde institucional
            COLOR_TEXTO = "#013051"       # Azul institucional
            FIG_WIDTH = 8
            FIG_HEIGHT = 5
            DPI = 300
            # ===================================
        
            fig, ax = plt.subplots(figsize=(FIG_WIDTH, FIG_HEIGHT))
        
            barras = ax.bar(
                df["categoria"],
                df["valor"],
                color=COLOR_BARRAS
            )
        
            ax.set_ylabel("")
            ax.set_xlabel("")
            ax.set_title("")
        
            ax.tick_params(axis="x", rotation=45)
        
            # Quitar marcos
            for spine in ax.spines.values():
                spine.set_visible(False)
        
            ax.tick_params(left=False, bottom=False)
        
            # Etiquetas encima de barras
            for bar in barras:
                height = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    height,
                    f"{int(height)}",
                    ha="center",
                    va="bottom",
                    fontsize=10,
                    color=COLOR_TEXTO
                )
        
            plt.tight_layout()
        
            plt.savefig(
                ASSETS_DIR / "grafico_p14.png",
                dpi=DPI,
                transparent=True
            )
        
            plt.close()
        
        # Ejecutar
        generar_grafico_p14(df_grafico_p14)

        
        # =========================================
        # TABLA PAGINA 14 (MODALIDADES POR DISTRITO)
        # =========================================
        
        # Encabezados (Fila 207 → índice 206)
        encabezados_p14 = df.iloc[206, 0:9].copy()
        
        # Datos (A208:I219 → índices 207:219)
        tabla_p14_raw = df.iloc[207:219, 0:9].copy()
        
        # Eliminar filas completamente vacías
        tabla_p14_raw = tabla_p14_raw.dropna(how="all")
        
        # Mantener celdas vacías individuales
        tabla_p14 = []
        
        # Agregar encabezados primero
        tabla_p14.append(encabezados_p14.fillna("").astype(str).tolist())
        
        # Agregar filas válidas
        for _, row in tabla_p14_raw.iterrows():
            # Ignorar fila si TODAS las frecuencias están vacías
            if row.iloc[1:].isna().all():
                continue
        
            fila = []
            for cell in row:
                if pd.isna(cell):
                    fila.append("")
                else:
                    fila.append(str(int(cell)) if isinstance(cell, (int, float)) else str(cell))
        
            tabla_p14.append(fila)


        #==================================PAGINA 15============================================================
        # =========================================
        # GRAFICO LINEAL PAGINA 15 (DIAS SEMANA)
        # RANGO: A223:C229
        # =========================================
        
        df_grafico_p15 = df.iloc[222:229, 0:3].copy()
        df_grafico_p15.columns = ["dia", "frecuencia", "porcentaje"]
        
        df_grafico_p15 = df_grafico_p15.dropna(how="all")
        
        # Formatear frecuencia
        df_grafico_p15["frecuencia"] = pd.to_numeric(
            df_grafico_p15["frecuencia"],
            errors="coerce"
        )
        
        # Formatear porcentaje
        df_grafico_p15["porcentaje"] = (
            df_grafico_p15["porcentaje"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        
        df_grafico_p15["porcentaje"] = pd.to_numeric(
            df_grafico_p15["porcentaje"],
            errors="coerce"
        )
        
        df_grafico_p15 = df_grafico_p15.dropna()
        
        def generar_grafico_p15(df):
        
            # ===== VARIABLES CONFIGURABLES =====
            COLOR_LINEA = "#013051"
            COLOR_PUNTOS = "#30A907"
            COLOR_RELLENO = "#30A907"
            COLOR_TEXTO = "#013051"
            COLOR_GRILLA = "#E0E0E0"
        
            FIG_WIDTH = 8
            FIG_HEIGHT = 5
            DPI = 300
            # ===================================
        
            fig, ax = plt.subplots(figsize=(FIG_WIDTH, FIG_HEIGHT))
        
            # Convertir eje X a numérico para mayor control
            x = range(len(df))
        
            # Línea principal
            ax.plot(
                x,
                df["frecuencia"],
                marker="o",
                linewidth=3,
                markersize=8,
                color=COLOR_LINEA
            )
        
            # Relleno inferior
            ax.fill_between(
                x,
                df["frecuencia"],
                color=COLOR_RELLENO,
                alpha=0.15
            )
        
            # Grilla horizontal
            ax.grid(axis="y", linestyle="--", alpha=0.4, color=COLOR_GRILLA)
        
            # Etiquetas frecuencia
            for i, row in enumerate(df.itertuples()):
                ax.text(
                    i,
                    row.frecuencia,
                    f"{int(row.frecuencia)}",
                    ha="center",
                    va="bottom",
                    fontsize=10,
                    fontweight="bold",
                    color=COLOR_TEXTO
                )
        
            # Ajuste eje Y
            ax.set_ylim(0, df["frecuencia"].max() * 1.25)
        
            # Porcentajes debajo del punto
            offset = df["frecuencia"].max() * 0.08
        
            for i, row in enumerate(df.itertuples()):
                ax.text(
                    i,
                    row.frecuencia - offset,
                    f"{row.porcentaje * 100:.2f}%",
                    ha="center",
                    va="top",
                    fontsize=9,
                    color=COLOR_TEXTO
                )
        
            # Etiquetas del eje X (días)
            ax.set_xticks(x)
            ax.set_xticklabels(df["dia"])
        
            # Quitar bordes
            for spine in ax.spines.values():
                spine.set_visible(False)
        
            ax.tick_params(left=False, bottom=False)
        
            ax.set_ylabel("")
            ax.set_xlabel("")
            ax.set_title("")
        
            plt.tight_layout()
        
            plt.savefig(
                ASSETS_DIR / "grafico_p15.png",
                dpi=DPI,
                transparent=True
            )
        
            plt.close()

        generar_grafico_p15(df_grafico_p15)

       #========tabla p15
        # =========================================
        # TABLA PAGINA 15 (FRECUENCIA POR DISTRITO Y DIA)
        # RANGO: A222:O229
        # IGNORAR COLUMNAS B y C
        # =========================================
        
        # Encabezados (Dias) → Columna A (A223:A229)
        dias_p15 = df.iloc[222:229, 0].copy().astype(str).tolist()
        
        # Distritos → Fila 222 (D222:O222)
        distritos_p15 = df.iloc[221, 3:15].copy().astype(str).tolist()
        
        # Frecuencias → D223:O229
        valores_p15_raw = df.iloc[222:229, 3:15].copy()
        
        tabla_p15 = []
        
        # Primera fila → encabezados (vacío + días)
        tabla_p15.append(["Distrito"] + dias_p15)
        
        # Construir filas
        for i, distrito in enumerate(distritos_p15):
        
            fila = [distrito]
        
            for valor in valores_p15_raw.iloc[:, i]:
        
                if pd.isna(valor):
                    fila.append("")
                else:
                    fila.append(str(int(valor)) if isinstance(valor, (int, float)) else str(valor))
        
            # Ignorar fila si completamente vacía (excepto nombre distrito)
            if all(v == "" for v in fila[1:]):
                continue
        
            tabla_p15.append(fila)

        # =========================================
        # DATOS LINEAS DE ACCION (PAGINA 17)
        # =========================================
        
        # Región (D2)
        region_numero = seguro_int(df.iloc[1, 3])  # D2
        
        # Delegación (B3) → formato "D-28"
        delegacion_codigo = str(df.iloc[2, 1])  # B3
        
        # Extraer número 28 de "D-28"
        numero_delegacion = int(delegacion_codigo.replace("D-", "").strip())
        
        # Totales líneas
        # ================= SAFE INT =================
        def safe_int(value):
            if pd.isna(value) or value == "":
                return 0
            return int(value)
        
        # Totales líneas
        lineas_municipalidad = seguro_int(df.iloc[238, 0])  # A239
        lineas_fp = seguro_int(df.iloc[238, 1])             # B239
        lineas_mixtas = safe_int(df.iloc[238, 2])         # C239
        total_lineas = seguro_int(df.iloc[238, 3])          # D239
        
        # Si mixtas es 0 no se muestra
        if lineas_mixtas == 0:
            lineas_mixtas = None
                
        # Construcción automática del path del logo municipal
        logo_muni_path = (
            ASSETS_DIR /
            "Municipalidades" /
            str(region_numero) /
            f"{numero_delegacion}.png"
        )


        # =========================================================
        # ================= PERCEPCION CIUDADANA ==================
        # =========================================================
        
        st.write("Filas del DataFrame:", df.shape[0])
        st.write("Columnas del DataFrame:", df.shape[1])
        # ================= PREGUNTA ACTUAL =================
        df_percepcion_actual = df.iloc[283:285, 0:2].copy()
        df_percepcion_actual.columns = ["respuesta", "porcentaje"]
        
        df_percepcion_actual["porcentaje"] = (
            df_percepcion_actual["porcentaje"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        
        df_percepcion_actual["porcentaje"] = pd.to_numeric(
            df_percepcion_actual["porcentaje"],
            errors="coerce"
        )
        
        df_percepcion_actual = df_percepcion_actual.dropna()
        
        # ===== GRAFICO PASTEL =====
        def generar_grafico_percepcion_actual(df):
        
            FIG_WIDTH = 5
            FIG_HEIGHT = 5
            DPI = 300
        
            fig, ax = plt.subplots(figsize=(FIG_WIDTH, FIG_HEIGHT))
        
            COLORES_PERCEPCION = [
                "#30A907",  # Verde institucional
                "#013051",  # Azul institucional
                "#A5A5A5"   # Gris (si hubiera tercera categoría)
            ]
        
            wedges, texts, autotexts = ax.pie(
                df["porcentaje"],
                labels=df["respuesta"],
                autopct=lambda p: f"{p:.2f}%",
                startangle=90,
                colors=COLORES_PERCEPCION[:len(df)],  # <-- AQUI ESTABA LO QUE FALTABA
                textprops={"fontsize": 18}
            )
        
            # Tamaño porcentajes
            for autotext in autotexts:
                autotext.set_fontsize(20)
                autotext.set_color("white")  # opcional, se ve más profesional
        
            ax.axis("equal")
        
            plt.tight_layout()
        
            plt.savefig(
                ASSETS_DIR / "grafico_percepcion_actual.png",
                dpi=DPI,
                transparent=True
            )
        
            plt.close()
        
        generar_grafico_percepcion_actual(df_percepcion_actual)
        
        
        # ================= COMPARACION AÑO ANTERIOR =================
        df_percepcion_comparacion = df.iloc[290:293, 0:2].copy()
        df_percepcion_comparacion.columns = ["categoria", "porcentaje"]
        
        df_percepcion_comparacion["porcentaje"] = (
            df_percepcion_comparacion["porcentaje"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        
        df_percepcion_comparacion["porcentaje"] = pd.to_numeric(
            df_percepcion_comparacion["porcentaje"],
            errors="coerce"
        )
        
        df_percepcion_comparacion = df_percepcion_comparacion.dropna()
        
        # ===== GRAFICO BARRAS =====
        def generar_grafico_percepcion_comparacion(df):
        
            FIG_WIDTH = 6
            FIG_HEIGHT = 5
            DPI = 300
            COLOR_BARRAS = "#30A907"
        
            fig, ax = plt.subplots(figsize=(FIG_WIDTH, FIG_HEIGHT))
        
            barras = ax.bar(
                df["categoria"],
                df["porcentaje"],
                color=COLOR_BARRAS
            )
        
            ax.set_ylim(0, df["porcentaje"].max() * 1.25)
        
            for bar in barras:
                height = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width()/2,
                    height,
                    f"{height:.2f}%",
                    ha="center",
                    va="bottom",
                    fontsize=16
                )
        
            ax.tick_params(axis="x", rotation=45)
            ax.tick_params(axis="x", labelsize=14)
            ax.tick_params(axis="y", labelsize=14)
        
            for spine in ax.spines.values():
                spine.set_visible(False)
        
            plt.tight_layout()
        
            plt.savefig(
                ASSETS_DIR / "grafico_percepcion_comparacion.png",
                dpi=DPI,
                transparent=True
            )
        
            plt.close()
        
        generar_grafico_percepcion_comparacion(df_percepcion_comparacion)
        
        # =========================================================
        # ================= VICTIMIZACION CIUDADANA PAGINA " PARTE FINAL===============
        # =========================================================
        
        # ----- GRAFICO 1 (A314:B316) -----
        df_victimizacion = df.iloc[313:316, 0:2].copy()
        df_victimizacion.columns = ["categoria", "porcentaje"]
        
        df_victimizacion["porcentaje"] = (
            df_victimizacion["porcentaje"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        
        df_victimizacion["porcentaje"] = pd.to_numeric(
            df_victimizacion["porcentaje"],
            errors="coerce"
        )
        
        df_victimizacion = df_victimizacion.dropna()


        # ----- GRAFICO 2 (A323:B330) -----
        df_no_denuncia = df.iloc[322:330, 0:2].copy()
        df_no_denuncia.columns = ["categoria", "porcentaje"]
        
        df_no_denuncia["porcentaje"] = (
            df_no_denuncia["porcentaje"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        
        df_no_denuncia["porcentaje"] = pd.to_numeric(
            df_no_denuncia["porcentaje"],
            errors="coerce"
        )
        
        df_no_denuncia = df_no_denuncia.dropna()
       
       # ----- TABLA INFERIOR (A323:C330 ignorando B) -----

        tabla_no_denuncia_df = df.iloc[322:330, [0, 2]].copy()
        tabla_no_denuncia_df = tabla_no_denuncia_df.dropna(how="all")
        
        ## TABLA
        tabla_no_denuncia = []

        for _, row in tabla_no_denuncia_df.iterrows():
        
            categoria = row.iloc[0] if len(row) > 0 else None
            valor = row.iloc[1] if len(row) > 1 else None
        
            if pd.notna(categoria) and pd.notna(valor):
                tabla_no_denuncia.append([
                    str(categoria),
                    f"{float(valor)*100:.2f}%"
                ])

        ## mayor frecuencia

        # Obtener motivo con mayor frecuencia
        if not df_no_denuncia.empty and df_no_denuncia["porcentaje"].notna().any():
            fila_max = df_no_denuncia.loc[df_no_denuncia["porcentaje"].idxmax()]
            motivo_principal = str(fila_max["categoria"])
        else:
            motivo_principal = "No especificado"
            
        if df.shape[0] > 321 and df.shape[1] > 6:
            total_omitidas = seguro_int(df.iloc[321, 6])
        else:
            total_omitidas = 0 # G322


        # =========================================================
        # ============== TABLA COMPARATIVA PERCEPCION =============
        # =========================================================
        
        # ----- Validar tamaño mínimo del DataFrame -----
        FILAS_MINIMAS = 309
        COLUMNAS_MINIMAS = 7
        
        filas_actuales, columnas_actuales = df.shape
        
        # Expandir filas si faltan
        if filas_actuales < FILAS_MINIMAS:
            filas_extra = FILAS_MINIMAS - filas_actuales
            df_extra = pd.DataFrame(
                [[None] * columnas_actuales] * filas_extra
            )
            df = pd.concat([df, df_extra], ignore_index=True)
        
        # Expandir columnas si faltan
        if columnas_actuales < COLUMNAS_MINIMAS:
            columnas_extra = COLUMNAS_MINIMAS - columnas_actuales
            for i in range(columnas_extra):
                df[f"_extra_{i}"] = None
        
        # ----- Extraer rango seguro -----
        tabla_percepcion_df = df.iloc[297:309, 0:7].copy()
        
        # Eliminar columnas B, D y F (índices 1,3,5)
        try:
            tabla_percepcion_df = tabla_percepcion_df.drop(
                tabla_percepcion_df.columns[[1,3,5]],
                axis=1
            )
        except:
            pass
        
        # ----- Formatear porcentajes SOLO desde fila 300 en adelante -----
        # (las dos primeras filas son encabezado)
        # ===== FORMATEAR SOLO COLUMNAS PORCENTUALES =====
        # Columnas 1,2,3 después de eliminar B,D,F
        for col in [1, 2, 3]:
            for fila in range(2, tabla_percepcion_df.shape[0]):
                valor = tabla_percepcion_df.iat[fila, col]
                if pd.notna(valor) and valor != "":
                    try:
                        tabla_percepcion_df.iat[fila, col] = f"{float(valor)*100:.2f}%"
                    except:
                        pass
                else:
                    tabla_percepcion_df.iat[fila, col] = ""
        
        # ----- Limpiar nulos -----
        tabla_percepcion = (
            tabla_percepcion_df
            .fillna("")
            .astype(str)
            .values
            .tolist()
        )
        
        
        # =========================================
        # LINEAS DE ACCION DINAMICAS PORTADAS
        # =========================================
        
        import string
        
        # Columnas donde están los corresponsables
        columnas_corresponsables = [
            9,   # J246
            15,  # P246
            21,  # V246
            27,  # AB246
            33,  # AH246
            39,  # AN246
            45,  # AT246
            51,  # AZ246
            57,  # BF246
            63,  # BL246
            69,  # BR246
            75   # BX246
        ]

        # Columnas porcentaje total (fila 242)
        columnas_total_porcentaje = [
            9,   # J242
            15,  # P242
            21,  # V242
            27,  # AB242
            33,  # AH242
            39,  # AN242
            45,  # AT242
            51,  # AZ242
            57,  # BF242
            63,  # BL242
            69,  # BR242
            75   # BX242
        ]

        

        # ================= COLUMNAS DINAMICAS LINEAS DE ACCION =================

        columnas_causas = [5, 11, 17, 23, 29, 35, 41, 47, 53, 59, 65, 71]
        columnas_problemas = [6, 12, 18, 24, 30, 36, 42, 48, 54, 60, 66, 72]
        columnas_lider = [9, 15, 21, 27, 33, 39, 45, 51, 57, 63, 69, 75]
        columnas_acciones = [8, 14, 20, 26, 32, 38, 44, 50, 56, 62, 68, 74]
        columnas_cogestores = [8, 14, 20, 26, 32, 38, 44, 50, 56, 62, 68, 74]
        columnas_corresponsable = [9,15, 21, 27, 33, 39, 45, 51, 57, 63, 69, 75]

        lineas_accion_data = []

        for i in range(int(total_lineas)):

            fila_problematicas = 241 + i  # B242 empieza en índice 241
        
            # ===== PROBLEMATICAS =====
            problematicas = []
            for col in [1, 2, 3]:  # B, C, D
                valor = df.iloc[fila_problematicas, col]
                if pd.notna(valor) and str(valor).strip() != "":
                    problematicas.append(str(valor).strip())
        
            # ===== CAUSAS =====
            col_causa = columnas_causas[i]
            causas = []
        
            for fila in range(246, 276):  # F247 a F276
                valor = df.iloc[fila, col_causa]
                if pd.notna(valor) and str(valor).strip() != "":
                    causas.append([str(valor).strip()])
        
            # ===== PROBLEMAS INFLUYENTES =====
            col_problema = columnas_problemas[i]
            problemas = []
        
            for fila in range(246, 276):
                valor = df.iloc[fila, col_problema]
                if pd.notna(valor) and str(valor).strip() != "":
                    problemas.append([str(valor).strip()])
        
            # ===== PORCENTAJE TOTAL =====
            col_total = columnas_total_porcentaje[i]
            valor_total = df.iloc[241, col_total]  # Fila 242
        
            if pd.notna(valor_total):
                try:
                    total_porcentaje = float(valor_total) * 100
                    total_porcentaje = f"{total_porcentaje:.2f}%"
                except:
                    total_porcentaje = "0.00%"
            else:
                total_porcentaje = "0.00%"
        
            # ===== LIDER ESTRATEGICO =====
            col_lider = columnas_lider[i]
            lider_estrategico = df.iloc[245, col_lider]

            # ===== CORRESPONSABLE =====
            col_corresponsable = columnas_corresponsable[i]
            corresponsable = df.iloc[245, col_corresponsable]
            
            if pd.notna(corresponsable):
                corresponsable = str(corresponsable).strip()
            else:
                corresponsable = ""
        
            if pd.notna(lider_estrategico):
                lider_estrategico = str(lider_estrategico).strip()
            else:
                lider_estrategico = ""
        
            # ===== ACCIONES ESTRATEGICAS =====
            col_acciones = columnas_acciones[i]
            acciones = []
        
            for fila in range(248, 257):  # I249 a I257
                valor = df.iloc[fila, col_acciones]
                if pd.notna(valor) and str(valor).strip() != "":
                    acciones.append(str(valor).strip())
        
            # ===== COGESTORES =====
            col_cogestores = columnas_cogestores[i]
            cogestores = []
        
            for fila in range(261, 263):  # I262 a I263
                valor = df.iloc[fila, col_cogestores]
                if pd.notna(valor) and str(valor).strip() != "":
                    partes = str(valor).split(",")
                    for p in partes:
                        if p.strip():
                            cogestores.append(p.strip())
        
            # ===== APPEND FINAL =====
            lineas_accion_data.append({
                "numero": i + 1,
                "problematicas": problematicas,
                "causas": causas,
                "problemas_influyentes": problemas,
                "total_porcentaje": total_porcentaje,
                "lider_estrategico": lider_estrategico,
                "acciones": acciones,
                "cogestores": cogestores,
                "corresponsable": corresponsable  # 👈 NUEVO
            })
                
        st.write("Cantidad de lineas detectadas:", len(lineas_accion_data)) ###debugging


        ##==================== graficos pagina 2===========

        def generar_grafico_victimizacion(df, nombre_archivo):

            COLOR_BARRAS = "#30A907"
            COLOR_TEXTO = "#013051"
            FIG_WIDTH = 12
            FIG_HEIGHT = 5
            DPI = 300
        
            fig, ax = plt.subplots(figsize=(FIG_WIDTH, FIG_HEIGHT))
        
            barras = ax.bar(
                df["categoria"],
                df["porcentaje"],
                color=COLOR_BARRAS
            )
        
            ax.set_ylim(0, df["porcentaje"].max() * 1.25)
        
            ax.set_xticklabels(
                df["categoria"],
                rotation=30,
                ha="right"
            )
        
            for bar in barras:
                height = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width()/2,
                    height,
                    f"{height:.2f}%",
                    ha="center",
                    va="bottom",
                    fontsize=12,
                    color=COLOR_TEXTO
                )
        
            for spine in ax.spines.values():
                spine.set_visible(False)
        
            plt.subplots_adjust(bottom=0.30)
            plt.tight_layout()
        
            plt.savefig(
                ASSETS_DIR / nombre_archivo,
                dpi=DPI,
                transparent=True
            )
        
            plt.close()
        

        
            
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
                total_denuncias=total_denuncias,
                grafico_horario_path=str(ASSETS_DIR / "grafico_horario.png"),
                tabla_horario=tabla_horario,
                total_am=total_am,
                total_pm=total_pm,
                tabla_horario_distrito=tabla_horario_distrito,
                grafico_p14_path=str(ASSETS_DIR / "grafico_p14.png"),
                tabla_p14=tabla_p14,
                grafico_p15_path=str(ASSETS_DIR / "grafico_p15.png"),
                tabla_p15=tabla_p15,
                total_lineas=total_lineas,
                lineas_municipalidad=lineas_municipalidad,
                lineas_fp=lineas_fp,
                lineas_mixtas=lineas_mixtas,
                logo_muni_path=str(logo_muni_path),
                lineas_accion_data=lineas_accion_data,
                grafico_percepcion_actual_path=str(ASSETS_DIR / "grafico_percepcion_actual.png"),
                grafico_percepcion_comparacion_path=str(ASSETS_DIR / "grafico_percepcion_comparacion.png"),
                tabla_percepcion=tabla_percepcion,
                grafico_victimizacion_path=str(ASSETS_DIR / "grafico_victimizacion.png"),
                grafico_no_denuncia_path=str(ASSETS_DIR / "grafico_no_denuncia.png"),
                tabla_no_denuncia=tabla_no_denuncia,
                motivo_principal=motivo_principal,
                total_omitidas=total_omitidas,
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
        import traceback
        st.error("ERROR DETALLADO:")
        st.code(traceback.format_exc())
