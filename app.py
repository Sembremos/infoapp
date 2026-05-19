import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path
from openpyxl import load_workbook
import numpy as np

# Forzar backend seguro para no-interactivos en hilos web
plt.switch_backend('Agg')

from pdf_generator import generar_pdf

st.set_page_config(page_title="Generador de Informes SS Level Up", layout="centered")
st.title("Generador Automático de Informes Territoriales")

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
if not ASSETS_DIR.exists():
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

# ================= MÓDULOS DE PROCESAMIENTO SEGURO =================
def seguro_int(valor, default=0):
    if pd.isna(valor) or str(valor).strip() == "":
        return default
    try:
        return int(float(valor))
    except:
        return default

def safe_float(valor, default=0.0):
    if pd.isna(valor) or str(valor).strip() == "":
        return default
    try:
        return float(valor)
    except:
        return default

# ================= CONVERSOR DE GRÁFICOS INTERNOS (BUFFER) =================
def salvar_grafico_a_disco(fig, filename):
    ruta = ASSETS_DIR / filename
    fig.savefig(ruta, dpi=250, bbox_inches="tight", transparent=True)
    plt.close(fig)
    return str(ruta)

# ================= FLUJO PRINCIPAL STREAMLIT =================
archivo = st.file_uploader("Cargue el archivo binario ENGINE (.xlsx) de la Delegación", type=["xlsx"])

if archivo:
    try:
        wb = load_workbook(archivo, data_only=True)
        ws = wb["Hoja1"]
        df = pd.DataFrame(ws.values)
        
        st.success("⚡ Estructura del Engine leída correctamente. Procesando rangos cartográficos...")

        # 1. Extracción de Identificadores Base
        delegacion = str(df.iloc[1, 1]).strip() if pd.notna(df.iloc[1, 1]) else "Desconocida"
        codigo = str(df.iloc[2, 1]).strip() if pd.notna(df.iloc[2, 1]) else "D0"
        
        # 2. Tabla Muestral de Participación por Distrito
        tabla_df = df.iloc[6:23, 0:3].dropna(how="all")
        def formatear_celda(v):
            if isinstance(v, (int, float)):
                return f"{v*100:.1f}%" if 0 < v < 1 else f"{int(v)}"
            return str(v)
        tabla_participacion = [[formatear_celda(c) for c in fila] for fila in tabla_df.fillna("").values.tolist()]

        # 3. Gráfico de Barras: Relación por Distrito
        rel_labels = df.iloc[7:11, 6].astype(str).tolist()
        rel_base_values = pd.to_numeric(df.iloc[7:11, 8], errors="coerce").fillna(0).tolist()
        
        fig_rel, ax_rel = plt.subplots(figsize=(7, 4))
        barras = ax_rel.bar(rel_labels, rel_base_values, color="#30A907", width=0.5)
        ax_rel.set_facecolor("none")
        for spine in ax_rel.spines.values(): spine.set_visible(False)
        ax_rel.tick_params(left=False, bottom=False, labelsize=10, colors="#013051")
        for bar in barras:
            h = bar.get_height()
            ax_rel.text(bar.get_x() + bar.get_width()/2, h + 0.5, f"{int(h)}", ha="center", va="bottom", fontsize=10, color="#013051", fontweight='bold')
        grafico_rel_path = salvar_grafico_a_disco(fig_rel, "grafico_relacion.png")

        # 4. Variables Demográficas (Edad, Escolaridad y Género)
        edad_labels = df.iloc[28:33, 0].astype(str).tolist()
        edad_vals = pd.to_numeric(df.iloc[28:33, 1], errors="coerce").fillna(0).tolist()
        fig_e, ax_e = plt.subplots(figsize=(4, 4))
        ax_e.pie(edad_vals, labels=edad_labels, autopct='%1.0f%%', startangle=90, colors=["#5B9BD5", "#A5A5A5", "#4472C4", "#255E91", "#B7B7B7"])
        grafico_edad_path = salvar_grafico_a_disco(fig_e, "grafico_edad.png")
        tabla_edad = [[l, f"{v*100:.0f}%"] for l, v in zip(edad_labels, edad_vals)]

        escolaridad_labels = df.iloc[38:46, 0].astype(str).tolist()
        esco_vals = pd.to_numeric(df.iloc[38:46, 1], errors="coerce").fillna(0).tolist()
        fig_es, ax_es = plt.subplots(figsize=(4, 4))
        ax_es.pie(esco_vals, labels=escolaridad_labels, autopct='%1.0f%%', startangle=90)
        grafico_escolaridad_path = salvar_grafico_a_disco(fig_es, "grafico_escolaridad.png")
        tabla_escolaridad = [[l, f"{v*100:.0f}%"] for l, v in zip(escolaridad_labels, esco_vals)]

        genero_labels = df.iloc[51:54, 0].astype(str).tolist()
        gen_vals = pd.to_numeric(df.iloc[51:54, 1], errors="coerce").fillna(0).tolist()
        fig_g, ax_g = plt.subplots(figsize=(4, 4))
        ax_g.pie(gen_vals, labels=genero_labels, autopct='%1.0f%%', startangle=90, colors=["#5B9BD5", "#A5A5A5", "#4472C4"])
        grafico_genero_path = salvar_grafico_a_disco(fig_g, "grafico_genero.png")
        tabla_genero = [[l, f"{v*100:.0f}%"] for l, v in zip(genero_labels, gen_vals)]

        # 5. Datos Métricos de Validación General
        tabla_encuesta_comunidad = df.iloc[58:60, 0:4].fillna("").astype(str).values.tolist()
        tabla_otras_encuestas = df.iloc[62:65, 6:10].fillna("").astype(str).values.tolist()
        
        datos_pagina_8 = {
            "encuesta_comunidad": seguro_int(df.iloc[82, 1]),
            "encuesta_policial": seguro_int(df.iloc[83, 1]),
            "encuesta_comercio": seguro_int(df.iloc[84, 1]),
            "estadistica": seguro_int(df.iloc[86, 2]),
            "total_datos": seguro_int(df.iloc[87, 1])
        }

        # 6. Diagramación Estratégica (Pareto y MICMAC)
        datos_pagina_9 = {
            "lado_izquierdo": str(df.iloc[92, 0]),
            "derecha_superior": str(df.iloc[92, 1]),
            "derecha_inferior": str(df.iloc[92, 2])
        }
        
        tabla_delitos = [[str(v)] for v in df.iloc[96:117, 1] if pd.notna(v) and str(v).strip() != "0"]
        tabla_riesgos = [[str(v)] for v in df.iloc[96:117, 2] if pd.notna(v) and str(v).strip() != "0"]
        
        porcentaje_delitos = f"{safe_float(df.iloc[118, 1]) * 100:.2f}%"
        porcentaje_riesgos = f"{safe_float(df.iloc[118, 2]) * 100:.2f}%"
        cantidad_delitos = seguro_int(df.iloc[117, 1])
        cantidad_riesgos = seguro_int(df.iloc[117, 2])

        # MICMAC Quadrants
        def build_clean_list(col_series):
            return [[str(v).strip()] for v in col_series if pd.notna(v) and str(v).strip() != ""]
            
        micmac_poder = build_clean_list(df.iloc[123:140, 1])
        micmac_conflicto = build_clean_list(df.iloc[123:140, 2])
        micmac_autonomas = build_clean_list(df.iloc[123:140, 3])
        micmac_resultados = build_clean_list(df.iloc[123:140, 4])
        
        tabla_riesgos_micmac2 = build_clean_list(df.iloc[123:140, 10])
        tabla_delitos_micmac2 = build_clean_list(df.iloc[123:140, 11])
        cantidad_problematicas = seguro_int(df.iloc[140, 12])
        riesgos_total = seguro_int(df.iloc[140, 10])
        delitos_total = seguro_int(df.iloc[140, 11])

        # 7. Causalidad y Alianzas Operativas
        causas_identificadas = seguro_int(df.iloc[117, 3])
        factores_micmac = seguro_int(df.iloc[140, 12])
        triangulo_directa = seguro_int(df.iloc[146, 0])
        triangulo_sociocultural = seguro_int(df.iloc[146, 1])
        triangulo_estructural = seguro_int(df.iloc[146, 2])
        
        tabla_instituciones = []
        for _, r in df.iloc[149:159, 1:3].dropna(how="all").iterrows():
            tabla_instituciones.append([str(r.iloc[0]).strip(), str(r.iloc[1]).strip()])

        # 8. Analítica Operativa Judicial (Incidencia OIJ)
        df_denuncias = df.iloc[165:176, [0, 2]].dropna()
        df_denuncias.columns = ["cat", "pct"]
        df_denuncias["pct"] = pd.to_numeric(df_denuncias["pct"].astype(str).str.replace("%","").str.replace(",","."), errors="coerce").fillna(0)
        
        fig_den, ax_den = plt.subplots(figsize=(5, 5))
        ax_den.pie(df_denuncias["pct"], labels=df_denuncias["cat"], autopct='%1.0f%%', startangle=90)
        grafico_denuncias_path = salvar_grafico_a_disco(fig_den, "grafico_denuncias.png")
        
        tabla_denuncias = df.iloc[165:176, [0, 1]].dropna().astype(str).values.tolist()
        total_denuncias = seguro_int(df.iloc[177, 1])

        # Rangos Cronológicos Horarios y Distribuciones Distritales
        df_horario = df.iloc[179:188, [0, 2]].dropna()
        df_horario.columns = ["hor", "pct"]
        df_horario["pct"] = pd.to_numeric(df_horario["pct"].astype(str).str.replace("%","").str.replace(",","."), errors="coerce").fillna(0)
        fig_hor, ax_hor = plt.subplots(figsize=(5, 5))
        ax_hor.pie(df_horario["pct"], labels=df_horario["hor"], autopct='%1.0f%%', startangle=90)
        grafico_horario_path = salvar_grafico_a_disco(fig_hor, "grafico_horario.png")
        
        tabla_horario = df.iloc[179:188, [0, 1]].dropna().astype(str).values.tolist()
        total_am = f"{safe_float(df.iloc[179, 3]) * 100:.1f}%"
        total_pm = f"{safe_float(df.iloc[179, 4]) * 100:.1f}%"
        
        # Maquetación limpia de tabla cruzada dinámica Horario-Distrito
        headers_hor_dist = df.iloc[178, 0:17].drop(df.iloc[178, [3,4]].index).fillna("").astype(str).tolist()
        rows_hor_dist = df.iloc[179:188, 0:17].drop(df.iloc[179:188, [3,4]].columns, axis=1).fillna("").astype(str).values.tolist()
        tabla_horario_distrito = [headers_hor_dist] + rows_hor_dist

        # Modalidades de Hechos (Página 14)
        df_p14 = df.iloc[195:204, [0, 1]].dropna()
        df_p14.columns = ["cat", "val"]
        df_p14["val"] = pd.to_numeric(df_p14["val"], errors="coerce").fillna(0)
        fig_p14, ax_p14 = plt.subplots(figsize=(6, 3.5))
        ax_p14.bar(df_p14["cat"], df_p14["val"], color="#30A907")
        ax_p14.set_facecolor("none")
        for spine in ax_p14.spines.values(): spine.set_visible(False)
        ax_p14.tick_params(axis='x', rotation=45, labelsize=8)
        grafico_p14_path = salvar_grafico_a_disco(fig_p14, "grafico_p14.png")
        tabla_p14 = df.iloc[206:219, 0:9].fillna("").astype(str).values.tolist()

        # Días de la semana (Página 15)
        df_p15 = df.iloc[222:229, 0:3].dropna()
        df_p15.columns = ["dia", "freq", "pct"]
        df_p15["freq"] = pd.to_numeric(df_p15["freq"], errors="coerce").fillna(0)
        fig_p15, ax_p15 = plt.subplots(figsize=(6, 3.5))
        ax_p15.plot(df_p15["dia"], df_p15["freq"], marker="o", color="#013051", linewidth=2)
        ax_p15.set_facecolor("none")
        for spine in ax_p15.spines.values(): spine.set_visible(False)
        grafico_p15_path = salvar_grafico_a_disco(fig_p15, "grafico_p15.png")
        
        # Tabla temporal cruzada días
        dias_h = ["Distrito"] + df.iloc[222:229, 0].astype(str).tolist()
        distritos_p15 = df.iloc[221, 3:15].fillna("").astype(str).tolist()
        valores_p15 = df.iloc[222:229, 3:15].fillna(0).astype(int).astype(str).values.T.tolist()
        tabla_p15 = [dias_h] + [[distritos_p15[idx]] + valores_p15[idx] for idx in range(len(distritos_p15))]

        # Resumen Operativo de Responsabilidades de Líneas
        total_lineas = seguro_int(df.iloc[238, 3])
        lineas_municipalidad = seguro_int(df.iloc[238, 0])
        lineas_fp = seguro_int(df.iloc[238, 1])
        lineas_mixtas = seguro_int(df.iloc[238, 2])
        
        # Búsqueda automatizada del logotipo municipal en base a ID geográfico regional
        region_id = seguro_int(df.iloc[1, 3])
        deleg_id = int(codigo.replace("D-", "").replace("D", "").strip())
        logo_muni_path = str(ASSETS_DIR / f"Municipalidades/{region_id}/{deleg_id}.png")
        if not os.path.exists(logo_muni_path):
            logo_muni_path = "assets/muni_default.png"  # Respaldo seguro para evitar excepciones de ReportLab

        # 9. Construcción del Segmento Dinámico Transversal: Líneas de Acción
        columnas_causas = [5, 11, 17, 23, 29, 35, 41, 47, 53, 59, 65, 71]
        columnas_problemas = [6, 12, 18, 24, 30, 36, 42, 48, 54, 60, 66, 72]
        columnas_lider = [9, 15, 21, 27, 33, 39, 45, 51, 57, 63, 69, 75]
        columnas_acciones = [8, 14, 20, 26, 32, 38, 44, 50, 56, 62, 68, 74]
        
        lineas_accion_data = []
        for idx in range(total_lineas):
            fila_prob = 241 + idx
            probs = [str(df.iloc[fila_prob, col]).strip() for col in [1, 2, 3] if pd.notna(df.iloc[fila_prob, col]) and str(df.iloc[fila_prob, col]).strip() != ""]
            
            causas = [[str(df.iloc[f, columnas_causas[idx]]).strip()] for f in range(246, 276) if pd.notna(df.iloc[f, columnas_causas[idx]]) and str(df.iloc[f, columnas_causas[idx]]).strip() != ""]
            pro_inf = [[str(df.iloc[f, columnas_problemas[idx]]).strip()] for f in range(246, 276) if pd.notna(df.iloc[f, columnas_problemas[idx]]) and str(df.iloc[f, columnas_problemas[idx]]).strip() != ""]
            
            pct_val = df.iloc[241, columnas_lider[idx]]
            total_pct = f"{float(pct_val)*100:.2f}%" if pd.notna(pct_val) else "0.00%"
            
            lider = str(df.iloc[245, columnas_lider[idx]]).strip() if pd.notna(df.iloc[245, columnas_lider[idx]]) else ""
            corresp = str(df.iloc[245, columnas_lider[idx]]).strip() if pd.notna(df.iloc[245, columnas_lider[idx]]) else ""
            
            accs = [str(df.iloc[f, columnas_acciones[idx]]).strip() for f in range(248, 257) if pd.notna(df.iloc[f, columnas_acciones[idx]]) and str(df.iloc[f, columnas_acciones[idx]]).strip() != ""]
            cogs = [str(df.iloc[f, columnas_acciones[idx]]).strip() for f in range(261, 263) if pd.notna(df.iloc[f, columnas_acciones[idx]]) and str(df.iloc[f, columnas_acciones[idx]]).strip() != ""]
            
            lineas_accion_data.append({
                "numero": idx + 1,
                "problematicas": probs,
                "causas": causas,
                "problemas_influyentes": pro_inf,
                "total_porcentaje": total_pct,
                "lider_estrategico": lider,
                "corresponsable": corresp,
                "acciones": accs,
                "cogestores": cogs
            })

        # 10. Percepción Ciudadana Final y Cifra Negra
        df_p_act = df.iloc[283:285, 0:2].dropna()
        df_p_act.columns = ["resp", "pct"]
        df_p_act["pct"] = pd.to_numeric(df_p_act["pct"].astype(str).str.replace("%","").str.replace(",","."), errors="coerce").fillna(0)
        fig_pa, ax_pa = plt.subplots(figsize=(4, 4))
        ax_pa.pie(df_p_act["pct"], labels=df_p_act["resp"], autopct='%1.1f%%', colors=["#30A907", "#013051"])
        grafico_percepcion_actual_path = salvar_grafico_a_disco(fig_pa, "grafico_percepcion_actual.png")

        df_p_comp = df.iloc[290:293, 0:2].dropna()
        df_p_comp.columns = ["cat", "pct"]
        df_p_comp["pct"] = pd.to_numeric(df_p_comp["pct"].astype(str).str.replace("%","").str.replace(",","."), errors="coerce").fillna(0)
        fig_pc, ax_pc = plt.subplots(figsize=(5, 3.5))
        ax_pc.bar(df_p_comp["cat"], df_p_comp["pct"], color="#30A907")
        ax_pc.set_facecolor("none")
        for spine in ax_pc.spines.values(): spine.set_visible(False)
        grafico_percepcion_comparacion_path = salvar_grafico_a_disco(fig_pc, "grafico_percepcion_comparacion.png")

        tabla_percepcion = df.iloc[297:309, [0,2,4,6]].fillna("").astype(str).values.tolist()

        # Victimización e Indicios No Denunciados
        df_vic = df.iloc[313:316, 0:2].dropna()
        df_vic.columns = ["cat", "pct"]
        df_vic["pct"] = pd.to_numeric(df_vic["pct"].astype(str).str.replace("%","").str.replace(",","."), errors="coerce").fillna(0)
        fig_v, ax_v = plt.subplots(figsize=(4, 4))
        ax_v.pie(df_vic["pct"], labels=df_vic["cat"], autopct='%1.1f%%')
        grafico_victimizacion_path = salvar_grafico_a_disco(fig_v, "grafico_victimizacion.png")

        df_nd = df.iloc[322:330, 0:2].dropna()
        df_nd.columns = ["cat", "pct"]
        df_nd["pct"] = pd.to_numeric(df_nd["pct"].astype(str).str.replace("%","").str.replace(",","."), errors="coerce").fillna(0)
        fig_nd, ax_nd = plt.subplots(figsize=(5, 3.5))
        ax_nd.bar(df_nd["cat"], df_nd["pct"], color="#4472C4")
        ax_nd.set_facecolor("none")
        for spine in ax_nd.spines.values(): spine.set_visible(False)
        ax_nd.tick_params(axis='x', rotation=30, labelsize=8)
        grafico_no_denuncia_path = salvar_grafico_a_disco(fig_nd, "grafico_no_denuncia.png")

        tabla_no_denuncia = [[str(r.iloc[0]), seguro_int(r.iloc[1])] for _, r in df.iloc[322:330, [0, 2]].dropna().iterrows()]
        motivo_principal = str(df_nd.loc[df_nd["pct"].idxmax()]["cat"]) if not df_nd.empty else "No determinado"
        total_omitidas = seguro_int(df.iloc[321, 6])

        # Métricas de Entorno Perceptual (Horas y Armas comunidad)
        hor_labels = df.iloc[335:344, 0].fillna("").astype(str).tolist()
        hor_pacts = pd.to_numeric(df.iloc[335:344, 1].astype(str).str.replace("%","").str.replace(",","."), errors="coerce").fillna(0).tolist()
        fig_hp, ax_hp = plt.subplots(figsize=(4, 4))
        ax_hp.pie(hor_pacts, labels=hor_labels, autopct='%1.0f%%')
        grafico_horarios_percepcion = salvar_grafico_a_disco(fig_hp, "grafico_horarios_percepcion.png")
        tabla_horarios_percepcion = [[str(r[0]), seguro_int(r[1])] for r in df.iloc[335:344, [0, 2]].fillna(0).values.tolist()]
        horario_mayor = hor_labels[np.argmax(hor_pacts)] if hor_pacts else "No especificado"

        arm_labels = df.iloc[349:357, 0].fillna("").astype(str).tolist()
        arm_pacts = pd.to_numeric(df.iloc[349:357, 1].astype(str).str.replace("%","").str.replace(",","."), errors="coerce").fillna(0).tolist()
        fig_ap, ax_ap = plt.subplots(figsize=(4, 4))
        ax_ap.pie(arm_pacts, labels=arm_labels, autopct='%1.0f%%')
        grafico_armas_percepcion = salvar_grafico_a_disco(fig_ap, "grafico_armas_percepcion.png")
        tabla_armas = [[str(r[0]), seguro_int(r[1])] for r in df.iloc[349:357, [0, 2]].fillna(0).values.tolist()]
        metodo_mas_usado = arm_labels[np.argmax(arm_pacts)] if arm_pacts else "No especificado"
        omitidas_aportes = seguro_int(df.iloc[321, 6])

        # Calificación del Servicio Policial (Fuerza Pública)
        serv_labels = df.iloc[362:367, 0].fillna("").astype(str).tolist()
        serv_vals = (pd.to_numeric(df.iloc[362:367, 1], errors="coerce").fillna(0) * 100).tolist()
        fig_sp, ax_sp = plt.subplots(figsize=(6, 3.5))
        ax_sp.bar(serv_labels, serv_vals, color="#013051")
        ax_sp.set_facecolor("none")
        for spine in ax_sp.spines.values(): spine.set_visible(False)
        grafico_servicio_policial = salvar_grafico_a_disco(fig_sp, "grafico_servicio_policial.png")

        # Comercio Especializado
        com_s_l = df.iloc[398:400, 0].fillna("").astype(str).tolist()
        com_s_v = pd.to_numeric(df.iloc[398:400, 1], errors="coerce").fillna(0).tolist()
        fig_cs, ax_cs = plt.subplots(figsize=(4, 4))
        ax_cs.pie(com_s_v, labels=com_s_l, autopct='%1.1f%%')
        grafico_comercio_seguridad = salvar_grafico_a_disco(fig_cs, "grafico_comercio_seguridad.png")

        com_p_l = df.iloc[405:407, 0].fillna("").astype(str).tolist()
        com_p_v = pd.to_numeric(df.iloc[405:407, 1], errors="coerce").fillna(0).tolist()
        fig_cp, ax_cp = plt.subplots(figsize=(4, 4))
        ax_cp.pie(com_p_v, labels=com_p_l, autopct='%1.1f%%')
        grafico_comercio_programa = salvar_grafico_a_disco(fig_cp, "grafico_comercio_programa.png")

        com_i_l = df.iloc[412:414, 0].fillna("").astype(str).tolist()
        com_i_v = pd.to_numeric(df.iloc[412:414, 1], errors="coerce").fillna(0).tolist()
        fig_ci, ax_ci = plt.subplots(figsize=(4, 4))
        ax_ci.pie(com_i_v, labels=com_i_l, autopct='%1.1f%%')
        grafico_comercio_inscrito = salvar_grafico_a_disco(fig_ci, "grafico_comercio_inscrito.png")

        com_c_l = df.iloc[419:421, 0].fillna("").astype(str).tolist()
        com_c_v = pd.to_numeric(df.iloc[419:421, 1], errors="coerce").fillna(0).tolist()
        fig_cc, ax_cc = plt.subplots(figsize=(4, 4))
        ax_cc.pie(com_c_v, labels=com_c_l, autopct='%1.1f%%')
        grafico_comercio_contacto = salvar_grafico_a_disco(fig_cc, "grafico_comercio_contacto.png")

        # Variables sueltas requeridas por firmas
        omitidas_servicio = seguro_int(df.iloc[386, 6])
        total_respuestas_servicio = seguro_int(df.iloc[382, 2])

        # ================= INVOCACIÓN DISPARADORA DEL PDF =================
        if st.button("🚀 GENERAR INFORME TERRITORIAL OPTIMIZADO"):
            pdf_buffer = generar_pdf(
                portada_path=str(ASSETS_DIR / "portada.png"),
                delegacion=delegacion, codigo=codigo,
                tabla_participacion=tabla_participacion,
                grafico_relacion_path=grafico_rel_path,
                grafico_edad_path=grafico_edad_path, tabla_edad=tabla_edad,
                grafico_escolaridad_path=grafico_escolaridad_path, tabla_escolaridad=tabla_escolaridad,
                grafico_genero_path=grafico_genero_path, tabla_genero=tabla_genero,
                tabla_encuesta_comunidad=tabla_encuesta_comunidad, tabla_otras_encuestas=tabla_otras_encuestas,
                datos_pagina_8=datos_pagina_8, datos_pagina_9=datos_pagina_9,
                tabla_delitos=tabla_delitos, tabla_riesgos=tabla_riesgos,
                porcentaje_delitos=porcentaje_delitos, porcentaje_riesgos=porcentaje_riesgos,
                cantidad_delitos=cantidad_delitos, cantidad_riesgos=cantidad_riesgos,
                micmac_poder=micmac_poder, micmac_conflicto=micmac_conflicto,
                micmac_autonomas=micmac_autonomas, micmac_resultados=micmac_resultados,
                tabla_riesgos_micmac2=tabla_riesgos_micmac2, tabla_delitos_micmac2=tabla_delitos_micmac2,
                cantidad_problematicas=cantidad_problematicas, riesgos_total=riesgos_total, delitos_total=delitos_total,
                causas_identificadas=causas_identificadas, factores_micmac=factores_micmac,
                triangulo_directa=triangulo_directa, triangulo_sociocultural=triangulo_sociocultural, triangulo_estructural=triangulo_estructural,
                tabla_instituciones=tabla_instituciones,
                grafico_denuncias_path=grafico_denuncias_path, tabla_denuncias=tabla_denuncias, total_denuncias=total_denuncias,
                grafico_horario_path=grafico_horario_path, tabla_horario=tabla_horario, total_am=total_am, total_pm=total_pm,
                tabla_horario_distrito=tabla_horario_distrito,
                grafico_p14_path=grafico_p14_path, tabla_p14=tabla_p14,
                grafico_p15_path=grafico_p15_path, tabla_p15=tabla_p15,
                total_lineas=total_lineas, lineas_municipalidad=lineas_municipalidad, lineas_fp=lineas_fp, lineas_mixtas=lineas_mixtas,
                logo_muni_path=logo_muni_path, lineas_accion_data=lineas_accion_data,
                grafico_percepcion_actual_path=grafico_percepcion_actual_path,
                grafico_percepcion_comparacion_path=grafico_percepcion_comparacion_path,
                tabla_percepcion=tabla_percepcion,
                grafico_victimizacion_path=grafico_victimizacion_path, grafico_no_denuncia_path=grafico_no_denuncia_path,
                tabla_no_denuncia=tabla_no_denuncia, motivo_principal=motivo_principal, total_omitidas=total_omitidas,
                grafico_horarios_percepcion=grafico_horarios_percepcion, grafico_armas_percepcion=grafico_armas_percepcion,
                tabla_horarios_percepcion=tabla_horarios_percepcion, tabla_armas=tabla_armas,
                horario_mayor=horario_mayor, metodo_mas_usado=metodo_mas_usado, omitidas_aportes=omitidas_aportes,
                grafico_servicio_policial=grafico_servicio_policial, omitidas_servicio=omitidas_servicio, total_respuestas_servicio=total_respuestas_servicio,
                grafico_comercio_seguridad=grafico_comercio_seguridad, grafico_comercio_programa=grafico_comercio_programa,
                grafico_comercio_inscrito=grafico_comercio_inscrito, grafico_comercio_contacto=grafico_comercio_contacto
            )
            
            st.subheader("🎉 ¡Informe construido con éxito!")
            st.download_button(
                label="⬇️ Descargar PDF de Alta Calidad",
                data=pdf_buffer.getvalue(),
                file_name=f"Informe_Territorial_{delegacion.replace(' ', '_')}.pdf",
                mime="application/pdf"
            )
            
    except Exception as e:
        import traceback
        st.error("🚨 Ocurrió un error crítico durante la compilación:")
        st.code(traceback.format_exc())
