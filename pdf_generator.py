import unicodedata
import re
import os
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Image,
    PageBreak,
    Spacer,
    Table,
    TableStyle,
    KeepTogether
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_CENTER, TA_RIGHT
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from io import BytesIO

# ================= PALETAS Y ESTILOS GLOBALES =================
COLOR_PRIMARIO = colors.HexColor("#013051")    # Azul Institucional
COLOR_SECUNDARIO = colors.HexColor("#30A907")  # Verde Institucional
COLOR_FONDO_CLARO = colors.HexColor("#F7F9FC")
COLOR_GRIS_TEXTO = colors.HexColor("#222222")

# Paleta Percepción
P3_PALETA_GRAFICO = ["#5b9bd5", "#a5a5a5", "#4472c4", "#255e91", "#636363", "#264478", "#7cafdd", "#b7b7b7", "#698ed0"]

# ================= BACKGROUNDS (FONDO Y MARCOS) =================
def draw_background_templates(canvas, doc):
    """Maneja de forma segura los fondos de página institucional basados en el número de página real."""
    page_width, page_height = A4
    
    # 1. Portada
    if doc.page == 1:
        if os.path.exists(doc.portada_path):
            canvas.drawImage(doc.portada_path, 0, 0, width=page_width, height=page_height, preserveAspectRatio=True)
        return

    # 2. Páginas de Transición / Portadillas Completas
    transiciones = {
        2: "assets/intro.png",
        4: "assets/participacion.png",
        7: "assets/metodologico.png",
        12: "assets/estadistica.png",
        16: "assets/lineas.png"
    }
    
    if doc.page in transiciones:
        ruta = transiciones[doc.page]
        if os.path.exists(ruta):
            canvas.drawImage(ruta, 0, 0, width=page_width, height=page_height, preserveAspectRatio=True)
        return

    # 3. Páginas de Flujo de Líneas de Acción Dinámicas
    # Identificar si es una portada de Línea de Acción (Cada primera página del bloque de 3 páginas)
    if doc.page >= 18 and doc.page < doc.percepcion_inicio:
        posicion_bloque = (doc.page - 18) % 3
        if posicion_bloque == 0:  # Portada de la Línea de Acción
            if os.path.exists("assets/la.png"):
                canvas.drawImage("assets/la.png", 0, 0, width=page_width, height=page_height, preserveAspectRatio=True)
            return

    # 4. Portada de Percepción Ciudadana
    if doc.page == doc.percepcion_inicio:
        if os.path.exists("assets/percepcion.png"):
            canvas.drawImage("assets/percepcion.png", 0, 0, width=page_width, height=page_height, preserveAspectRatio=True)
        return

    # 5. Contraportada Final
    if doc.page == doc.total_paginas_estimadas:
        if os.path.exists("assets/final.png"):
            canvas.drawImage("assets/final.png", 0, 0, width=page_width, height=page_height, preserveAspectRatio=True)
        return

    # 6. Páginas Estándar con Encabezado y Pie de página
    if os.path.exists("assets/header.png"):
        canvas.drawImage("assets/header.png", 0, page_height - 60, width=page_width, height=45, mask="auto")
    if os.path.exists("assets/footer.png"):
        canvas.drawImage("assets/footer.png", 0, 15, width=page_width, height=40, mask="auto")
        
    # Numeración de página limpia en el Footer
    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(COLOR_PRIMARIO)
    canvas.drawRightString(page_width - 40, 30, f"Página {doc.page}")


# ================= CONSTRUIDORES DE COMPONENTES DE TABLAS =================
def crear_tabla_estilizada(data, col_widths, primary_color=COLOR_PRIMARIO, alternar_filas=True, font_size=9):
    """Genera tablas con envoltorios de párrafo automáticos para evitar desbordes de columnas."""
    style_cell = ParagraphStyle('CellSt', fontName='Helvetica', fontSize=font_size, leading=font_size+3, textColor=COLOR_GRIS_TEXTO)
    style_header = ParagraphStyle('HeadSt', fontName='Helvetica-Bold', fontSize=font_size+1, leading=font_size+4, textColor=colors.white, alignment=TA_CENTER)
    
    processed_data = []
    for r_idx, row in enumerate(data):
        processed_row = []
        for cell in row:
            if r_idx == 0:
                processed_row.append(Paragraph(str(cell), style_header))
            else:
                processed_row.append(Paragraph(str(cell), style_cell))
        processed_data.append(processed_row)
        
    t = Table(processed_data, colWidths=col_widths, repeatRows=1)
    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), primary_color),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]
    if alternar_filas:
        t_style.append(('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, COLOR_FONDO_CLARO]))
    t.setStyle(TableStyle(t_style))
    return t

def normalizar_nombre(texto):
    texto = texto.lower()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = texto.replace(" ", "_")
    return re.sub(r"[^a-z0-9_()]", "", texto)


# ================= GENERADOR COMPLETO DE PDF =================
def generar_pdf(portada_path, **kwargs):
    buffer = BytesIO()
    
    # Configuración base del documento flotante
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=70,  # Espacio para respetar el Header institucional
        bottomMargin=65  # Espacio para respetar el Footer institucional
    )
    
    # Inyectar variables globales del flujo al documento para el canvas
    doc.portada_path = portada_path
    doc.total_lineas = int(kwargs.get("total_lineas", 0))
    doc.percepcion_inicio = 18 + (doc.total_lineas * 3)
    doc.total_paginas_estimadas = doc.percepcion_inicio + 7
    
    styles = getSampleStyleSheet()
    
    # Paleta tipográfica avanzada
    style_h1 = ParagraphStyle('H1_Custom', fontName='Helvetica-Bold', fontSize=22, leading=26, textColor=COLOR_PRIMARIO, spaceAfter=15, spaceBefore=10)
    style_h2 = ParagraphStyle('H2_Custom', fontName='Helvetica-Bold', fontSize=15, leading=19, textColor=COLOR_PRIMARIO, spaceAfter=10, spaceBefore=15)
    style_body = ParagraphStyle('Body_Custom', fontName='Helvetica', fontSize=10.5, leading=15, textColor=COLOR_GRIS_TEXTO, alignment=TA_JUSTIFY, spaceAfter=10)
    
    story = []
    
    # ---------------- PAGINA 1: PORTADA ----------------
    # El fondo se maneja en draw_background_templates. Usamos Flowables para colocar el texto en los cuadros transparentes exactos.
    story.append(Spacer(1, 120))
    story.append(Paragraph("INFORME TERRITORIAL", ParagraphStyle('PortSub', fontName='Helvetica-Bold', fontSize=18, textColor=colors.white, spaceAfter=5)))
    story.append(Paragraph(kwargs.get("delegacion", "DELEGACIÓN").upper(), ParagraphStyle('PortDel', fontName='Helvetica-Bold', fontSize=32, leading=36, textColor=COLOR_SECUNDARIO, spaceAfter=5)))
    story.append(Paragraph(f"Código Territorial: {kwargs.get('codigo', '')}", ParagraphStyle('PortCod', fontName='Helvetica', fontSize=14, textColor=colors.white)))
    story.append(PageBreak())
    
    # ---------------- PAGINA 2: INTRODUCCIÓN (Fondo PNG) ----------------
    story.append(Spacer(1, 40))
    story.append(PageBreak())
    
    # ---------------- PAGINA 3: TEXTO INTRODUCTORIO ----------------
    story.append(Paragraph("Introducción", style_h1))
    story.append(Paragraph("Desde el año 2022, el Ministerio de Seguridad Pública ha implementado en todo el territorio nacional el Modelo Preventivo de Gestión Policial, una iniciativa estratégica destinada a fortalecer la seguridad pública a través de un enfoque proactivo y colaborativo...", style_body))
    story.append(Spacer(1, 15))
    if os.path.exists("assets/conformacion.png"):
        story.append(Image("assets/conformacion.png", width=480, height=320, preserveAspectRatio=True))
    story.append(PageBreak())
    
    # ---------------- PAGINA 4: PORTADILLA PARTICIPACIÓN (Fondo PNG) ----------------
    story.append(Spacer(1, 40))
    story.append(PageBreak())
    
    # ---------------- PAGINA 5: PARTICIPACIÓN POR DISTRITO ----------------
    story.append(Paragraph("Datos de Participación Ciudadana", style_h1))
    story.append(Paragraph("La recolección de impresiones en el campo permite entender las dinámicas específicas de cada sector habitacional y comercial. A continuación se detalla la cantidad muestral obtenida:", style_body))
    
    tabla_part = kwargs.get("tabla_participacion", [])
    if tabla_part:
        # Forzar fila de encabezado si no viene explícita
        if len(tabla_part[0]) == 3 and "distrito" in str(tabla_part[0][0]).lower():
            pass
        else:
            tabla_part.insert(0, ["Distrito", "Porcentaje", "Total Casos"])
        story.append(crear_tabla_estilizada(tabla_part, [220, 130, 130], COLOR_PRIMARIO))
        
    story.append(Spacer(1, 20))
    if os.path.exists(kwargs.get("grafico_relacion_path", "")):
        story.append(Paragraph("Relación del Muestreo por Distrito", style_h2))
        story.append(Image(kwargs.get("grafico_relacion_path"), width=490, height=220, preserveAspectRatio=True))
    story.append(PageBreak())
    
    # ---------------- PAGINA 6: DEMOGRAFÍA (EDAD, ESCOLARIDAD, GÉNERO) ----------------
    story.append(Paragraph("Variables Demográficas Recopiladas", style_h1))
    
    # Fusión estética de gráficos y tablas usando sub-tablas de ReportLab para maquetación lado a lado libre de solapamientos
    demo_data = []
    
    # Bloque Edad
    if os.path.exists(kwargs.get("grafico_edad_path", "")):
        t_edad = crear_tabla_estilizada([["Rango Edad", "%"]] + kwargs.get("tabla_edad", []), [130, 70], COLOR_SECUNDARIO, alternar_filas=False)
        demo_data.append([Image(kwargs.get("grafico_edad_path"), width=240, height=150), t_edad])
        
    # Bloque Escolaridad
    if os.path.exists(kwargs.get("grafico_escolaridad_path", "")):
        t_esco = crear_tabla_estilizada([["Escolaridad", "%"]] + kwargs.get("tabla_escolaridad", []), [130, 70], COLOR_PRIMARIO, alternar_filas=False)
        demo_data.append([t_esco, Image(kwargs.get("grafico_escolaridad_path"), width=240, height=150)])
        
    if demo_data:
        layout_demo = Table(demo_data, colWidths=[250, 250])
        layout_demo.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 15)
        ]))
        story.append(layout_demo)
    story.append(PageBreak())
    
    # ---------------- PAGINA 7: PORTADILLA PROCESO METODOLÓGICO ----------------
    story.append(Spacer(1, 40))
    story.append(PageBreak())
    
    # ---------------- PAGINA 8: DETALLE MUESTREO GENERAL ----------------
    story.append(Paragraph("Metodología Muestral Aplicada", style_h1))
    
    t_comunidad = kwargs.get("tabla_encuesta_comunidad", [])
    if t_comunidad:
        story.append(Paragraph("Muestreo del Sector Residencial/Comunidad", style_h2))
        story.append(crear_tabla_estilizada(t_comunidad, [125]*4, COLOR_PRIMARIO))
        
    story.append(Spacer(1, 15))
    t_otras = kwargs.get("tabla_otras_encuestas", [])
    if t_otras:
        story.append(Paragraph("Validaciones Multiactor Adicionales", style_h2))
        story.append(crear_tabla_estilizada(t_otras, [125]*4, COLOR_PRIMARIO))
        
    story.append(Spacer(1, 25))
    if os.path.exists("assets/datos.png"):
        # Contenedor absoluto simulado sobre flujo dinámico para los datos numéricos de la tarjeta "datos.png"
        story.append(Paragraph("Consolidado General de Respuestas", style_h2))
        story.append(Image("assets/datos.png", width=500, height=200, preserveAspectRatio=True))
    story.append(PageBreak())
    
    # ---------------- PAGINA 9: PARETO DELICTIVO Y SOCIAL ----------------
    story.append(Paragraph("Análisis de Problemáticas Prioritarias (Principio de Pareto)", style_h1))
    if os.path.exists("assets/pareto.png"):
        story.append(Image("assets/pareto.png", width=500, height=180, preserveAspectRatio=True))
        
    story.append(Spacer(1, 15))
    
    # Tablas lado a lado de Delitos vs Riesgos Sociales prioritarios
    t_delitos_raw = kwargs.get("tabla_delitos", [])
    t_riesgos_raw = kwargs.get("tabla_riesgos", [])
    
    t_delitos = crear_tabla_estilizada([["Delitos Prioritarios (Vitales)"]] + t_delitos_raw, [240], COLOR_SECUNDARIO)
    t_riesgos = crear_tabla_estilizada([["Riesgos Sociales Prioritarios"]] + t_riesgos_raw, [240], COLOR_PRIMARIO)
    
    layout_pareto = Table([[t_delitos, t_riesgos]], colWidths=[250, 250])
    layout_pareto.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(layout_pareto)
    story.append(PageBreak())
    
    # ---------------- PAGINA 10: MATRIZ MICMAC ----------------
    story.append(Paragraph("Análisis Estructural Cruzado (Matriz MICMAC)", style_h1))
    if os.path.exists("assets/micmac.png"):
        story.append(Image("assets/micmac.png", width=500, height=350, preserveAspectRatio=True))
    story.append(PageBreak())
    
    # ---------------- PAGINA 11: TRIÁNGULO DE LAS VIOLENCIAS E INSTITUCIONES ----------------
    story.append(Paragraph("Análisis Causal: Triángulo de las Violencias", style_h1))
    story.append(Paragraph("Bajo el modelo del sociólogo Johan Galtung, se segmentan las causas identificadas en los sectores estructurales, socioculturales y de violencia directa.", style_body))
    
    if os.path.exists("assets/triangulo.png"):
        story.append(Image("assets/triangulo.png", width=480, height=220, preserveAspectRatio=True))
        
    story.append(Spacer(1, 15))
    t_inst = kwargs.get("tabla_instituciones", [])
    if t_inst:
        story.append(Paragraph("Mesas de Trabajo de Calificación Coordinada", style_h2))
        story.append(crear_tabla_estilizada([["Institución Participante", "Sector de Aporte"]] + t_inst, [300, 200], COLOR_PRIMARIO))
    story.append(PageBreak())
    
    # ---------------- PAGINA 12: PORTADILLA ESTADÍSTICA (Fondo PNG) ----------------
    story.append(Spacer(1, 40))
    story.append(PageBreak())
    
    # ---------------- PAGINA 13: DENUNCIAS JUDICIALES (OIJ) Y MAPA HORARIO ----------------
    story.append(Paragraph("Análisis Técnico Operativo de Denuncias OIJ", style_h1))
    
    if os.path.exists(kwargs.get("grafico_denuncias_path", "")):
        story.append(Paragraph("Distribución de Frecuencia de Hechos por Distrito", style_h2))
        story.append(Image(kwargs.get("grafico_denuncias_path"), width=480, height=220, preserveAspectRatio=True))
        
    story.append(Spacer(1, 15))
    t_horario_dist = kwargs.get("tabla_horario_distrito", [])
    if t_horario_dist:
        story.append(Paragraph("Consolidado de Incidencia por Bloques Horarios", style_h2))
        # Ajuste dinámico de columnas según los distritos leídos del excel
        c_w = 500 / len(t_horario_dist[0])
        story.append(crear_tabla_estilizada(t_horario_dist, [c_w]*len(t_horario_dist[0]), COLOR_SECUNDARIO, font_size=7.5))
    story.append(PageBreak())
    
    # ---------------- PAGINA 14: MODALIDADES DELICTIVAS PRIORITARIAS ----------------
    story.append(Paragraph("Análisis Técnico por Modalidad de Ejecución", style_h1))
    if os.path.exists(kwargs.get("grafico_p14_path", "")):
        story.append(Image(kwargs.get("grafico_p14_path"), width=490, height=220, preserveAspectRatio=True))
        
    story.append(Spacer(1, 15))
    t_p14 = kwargs.get("tabla_p14", [])
    if t_p14:
        c_w = 500 / len(t_p14[0])
        story.append(crear_tabla_estilizada(t_p14, [c_w]*len(t_p14[0]), COLOR_PRIMARIO, font_size=7.5))
    story.append(PageBreak())
    
    # ---------------- PAGINA 15: INCIDENCIA CRONOLÓGICA (DÍAS DE LA SEMANA) ----------------
    story.append(Paragraph("Comportamiento y Tendencia Cronológica", style_h1))
    if os.path.exists(kwargs.get("grafico_p15_path", "")):
        story.append(Image(kwargs.get("grafico_p15_path"), width=490, height=220, preserveAspectRatio=True))
        
    story.append(Spacer(1, 15))
    t_p15 = kwargs.get("tabla_p15", [])
    if t_p15:
        c_w = 500 / len(t_p15[0])
        story.append(crear_tabla_estilizada(t_p15, [c_w]*len(t_p15[0]), COLOR_SECUNDARIO, font_size=7.5))
    story.append(PageBreak())
    
    # ---------------- PAGINA 16: PORTADILLA LÍNEAS DE ACCIÓN (Fondo PNG) ----------------
    story.append(Spacer(1, 40))
    story.append(PageBreak())
    
    # ---------------- PAGINA 17: RESUMEN DE LÍNEAS OPERATIVAS DETECTADAS ----------------
    story.append(Paragraph("Estrategia Integral Operativa Local", style_h1))
    story.append(Paragraph("A partir de la triangulación metodológica, se definen las siguientes líneas prioritarias de articulación interinstitucional:", style_body))
    
    if os.path.exists("assets/lins.png"):
        story.append(Image("assets/lins.png", width=500, height=260, preserveAspectRatio=True))
    story.append(PageBreak())
    
    # ---------------- PAGINAS 18+ : FLUJO DINÁMICO DE LÍNEAS DE ACCIÓN ----------------
    # Iteración limpia sin hardcoding de páginas. Cada línea de acción genera exactamente su bloque modular.
    lineas_data = kwargs.get("lineas_accion_data", [])
    for linea in lineas_data:
        # Sub-página A: Portada de la Línea (Manejada por el Canvas)
        story.append(Spacer(1, 200))
        story.append(Paragraph(f"LÍNEA DE ACCIÓN NÚMERO {linea['numero']}", ParagraphStyle('LNum', fontName='Helvetica-Bold', fontSize=24, textColor=colors.white, alignment=TA_CENTER)))
        story.append(Spacer(1, 20))
        story.append(Paragraph("<br/>".join(linea["problematicas"]), ParagraphStyle('LProb', fontName='Helvetica-Bold', fontSize=16, textColor=colors.white, alignment=TA_CENTER)))
        story.append(PageBreak())
        
        # Sub-página B: Diagnóstico de Vectores, Causas e Influencias
        story.append(Paragraph(f"Diagnóstico de la Línea Operativa {linea['numero']}", style_h1))
        
        # Inserción dinámica de iconos vectoriales si existen en los recursos
        iconos_flowables = []
        for prob in linea["problematicas"]:
            v_name = f"assets/vectores/{normalizar_nombre(prob)}.png"
            if os.path.exists(v_name):
                iconos_flowables.append(Image(v_name, width=90, height=90, preserveAspectRatio=True))
        if iconos_flowables:
            t_iconos = Table([iconos_flowables], rowHeights=100)
            t_iconos.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
            story.append(t_iconos)
            
        story.append(Spacer(1, 15))
        if linea["causas"]:
            story.append(Paragraph("Causas Socio-Culturales y Estructurantes Identificadas", style_h2))
            story.append(crear_tabla_estilizada(linea["causas"], [500], COLOR_SECUNDARIO))
            
        story.append(Spacer(1, 15))
        if linea["problemas_influyentes"]:
            story.append(Paragraph("Problemáticas Co-dependientes e Influyentes", style_h2))
            story.append(crear_tabla_estilizada(linea["problemas_influyentes"], [500], COLOR_PRIMARIO))
        story.append(PageBreak())
        
        # Sub-página C: Objetivos, Líder y Plan de Acción
        story.append(Paragraph(f"Plan de Articulación Estratégica - Línea {linea['numero']}", style_h1))
        story.append(Paragraph("<b>Objetivo general de la intervención:</b>", style_body))
        story.append(Paragraph("Desplegar acciones tácticas conjuntas para la contención, disuasión y mitigación de la problemática priorizada a través de indicadores de trazabilidad compartida.", style_body))
        
        if linea["lider_estrategico"]:
            t_lider = Table([["Líder Estratégico Asociado:", linea["lider_estrategico"]]], colWidths=[200, 300])
            t_lider.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (0,0), COLOR_SECUNDARIO),
                ('BACKGROUND', (1,0), (1,0), COLOR_FONDO_CLARO),
                ('TEXTCOLOR', (0,0), (0,0), colors.white),
                ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),
                ('GRID', (0,0), (-1,-1), 0.5, COLOR_PRIMARIO),
                ('PADDING', (0,0), (-1,-1), 8),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
            ]))
            story.append(t_lider)
            story.append(Spacer(1, 15))
            
        if linea["acciones"]:
            story.append(Paragraph("Acciones Tácticas Inmediatas", style_h2))
            rows_acc = [[f"• {acc}"] for acc in linea["acciones"]]
            story.append(crear_tabla_estilizada(rows_acc, [500], COLOR_PRIMARIO, alternar_filas=True))
            
        story.append(Spacer(1, 15))
        if linea["cogestores"]:
            story.append(Paragraph("Corresponsables Civiles e Institucionales (Co-gestores)", style_h2))
            rows_co = [[f"• {cog}"] for cog in linea["cogestores"]]
            story.append(crear_tabla_estilizada(rows_co, [500], COLOR_SECUNDARIO, alternar_filas=False))
            
        story.append(PageBreak())

    # ---------------- SECCIÓN FINAL: PERCEPCIÓN CIUDADANA Y VICTIMIZACIÓN ----------------
    # Sub-página Portada Percepción (Manejada por el Canvas)
    story.append(Spacer(1, 40))
    story.append(PageBreak())
    
    # Sub-página Diagnóstico Percepción Seguridad
    story.append(Paragraph("Estudio de Percepción de Seguridad Local", style_h1))
    perc_data = []
    if os.path.exists(kwargs.get("grafico_percepcion_actual_path", "")):
        perc_data.append(Image(kwargs.get("grafico_percepcion_actual_path"), width=240, height=220, preserveAspectRatio=True))
    if os.path.exists(kwargs.get("grafico_percepcion_comparacion_path", "")):
        perc_data.append(Image(kwargs.get("grafico_percepcion_comparacion_path"), width=240, height=220, preserveAspectRatio=True))
    if perc_data:
        t_perc = Table([perc_data], colWidths=[250, 250])
        story.append(t_perc)
        
    story.append(Spacer(1, 15))
    t_per_comp = kwargs.get("tabla_percepcion", [])
    if t_per_comp:
        c_w = 500 / len(t_per_comp[0])
        story.append(crear_tabla_estilizada(t_per_comp, [c_w]*len(t_per_comp[0]), COLOR_SECUNDARIO))
    story.append(PageBreak())
    
    # Sub-página Fenómeno de la No Denuncia y Cifra Negra
    story.append(Paragraph("Victimización y Fenómeno de la No Denuncia", style_h1))
    vic_data = []
    if os.path.exists(kwargs.get("grafico_victimizacion_path", "")):
        vic_data.append(Image(kwargs.get("grafico_victimizacion_path"), width=240, height=200, preserveAspectRatio=True))
    if os.path.exists(kwargs.get("grafico_no_denuncia_path", "")):
        vic_data.append(Image(kwargs.get("grafico_no_denuncia_path"), width=240, height=200, preserveAspectRatio=True))
    if vic_data:
        story.append(Table([vic_data], colWidths=[250, 250]))
        
    story.append(Spacer(1, 15))
    t_no_den = kwargs.get("tabla_no_denuncia", [])
    if t_no_den:
        # Formatear lista cruda en estructura de tabla aceptable
        t_data_no_den = [["Motivo Principal Declarado", "Frecuencia Absoluta"]] + t_no_den
        story.append(crear_tabla_estilizada(t_data_no_den, [350, 150], COLOR_PRIMARIO))
        
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"<b>Nota de campo:</b> La mayor cantidad de los encuestados que declararon ser víctimas de algún ilícito señalan que el principal motivo para la abstención de la denuncia penal es: <i>{kwargs.get('motivo_principal', 'No especificado')}</i>. Casos omitidos en este bloque: {kwargs.get('total_omitidas', 0)}.", style_body))
    story.append(PageBreak())
    
    # Sub-página Percepción del Entorno y Armas
    story.append(Paragraph("Factores Ambientales del Hecho Delictivo", style_h1))
    env_data = []
    if os.path.exists(kwargs.get("grafico_horarios_percepcion", "")):
        env_data.append(Image(kwargs.get("grafico_horarios_percepcion"), width=240, height=200, preserveAspectRatio=True))
    if os.path.exists(kwargs.get("grafico_armas_percepcion", "")):
        env_data.append(Image(kwargs.get("grafico_armas_percepcion"), width=240, height=200, preserveAspectRatio=True))
    if env_data:
        story.append(Table([env_data], colWidths=[250, 250]))
    story.append(PageBreak())
    
    # Sub-página Evaluación del Servicio de Fuerza Pública
    story.append(Paragraph("Evaluación de la Gestión Operativa del Servicio Policial", style_h1))
    if os.path.exists(kwargs.get("grafico_servicio_policial", "")):
        story.append(Image(kwargs.get("grafico_servicio_policial"), width=490, height=220, preserveAspectRatio=True))
    story.append(PageBreak())
    
    # Sub-página Diagnóstico Sector Comercial Especializado
    story.append(Paragraph("Percepción y Enlace con el Sector Comercial", style_h1))
    com_data_1 = []
    com_data_2 = []
    if os.path.exists(kwargs.get("grafico_comercio_seguridad", "")):
        com_data_1.append(Image(kwargs.get("grafico_comercio_seguridad"), width=240, height=180, preserveAspectRatio=True))
    if os.path.exists(kwargs.get("grafico_comercio_programa", "")):
        com_data_1.append(Image(kwargs.get("grafico_comercio_programa"), width=240, height=180, preserveAspectRatio=True))
    if os.path.exists(kwargs.get("grafico_comercio_inscrito", "")):
        com_data_2.append(Image(kwargs.get("grafico_comercio_inscrito"), width=240, height=180, preserveAspectRatio=True))
    if os.path.exists(kwargs.get("grafico_comercio_contacto", "")):
        com_data_2.append(Image(kwargs.get("grafico_comercio_contacto"), width=240, height=180, preserveAspectRatio=True))
        
    if com_data_1:
        story.append(Table([com_data_1], colWidths=[250, 250]))
    story.append(Spacer(1, 10))
    if com_data_2:
        story.append(Table([com_data_2], colWidths=[250, 250]))
    story.append(PageBreak())
    
    # ---------------- PAGINA FINAL: CONTRAPORTADA ----------------
    story.append(Spacer(1, 40))

    # Construcción de las capas del documento
    doc.build(
        story,
        onFirstPage=draw_background_templates,
        onLaterPages=draw_background_templates
    )
    
    buffer.seek(0)
    return buffer
