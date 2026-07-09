import unicodedata
import re
import os
from io import BytesIO

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Image,
    PageBreak,
    Spacer,
    Table,
    TableStyle
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

# =====================================================================
# CONFIGURACIÓN GLOBAL ESTÉTICA (COLORES, FUENTES, ESTILOS)
# =====================================================================
COLOR_PRIMARIO = colors.HexColor("#30A907")      # Verde Institucional
COLOR_SECUNDARIO = colors.HexColor("#013051")    # Azul Institucional
COLOR_HEADER_TABLA = colors.HexColor("#DEEBF7")
COLOR_FILA_ALTERNA = colors.HexColor("#F2F2F2")
COLOR_BORDE = colors.HexColor("#A0A0A0")         # Borde gris elegante
COLOR_TEXTO_OSCURO = colors.HexColor("#222222")

FONT_NAME_BOLD = "Helvetica-Bold"
FONT_NAME_REGULAR = "Helvetica"

# Paleta exportada para uso en app.py (Página 3 y general)
P3_PALETA_GRAFICO = [
    "#5b9bd5", "#a5a5a5", "#4472c4", "#255e91", "#636363",
    "#264478", "#7cafdd", "#b7b7b7", "#698ed0"
]

# Estilo global para tablas de Pareto
pareto_cell_style = ParagraphStyle(
    name="ParetoCell",
    fontName=FONT_NAME_REGULAR,
    fontSize=10,
    leading=12,
    alignment=TA_LEFT,
    wordWrap="CJK"
)

# Estilo base para tablas estandarizadas
def obtener_estilo_tabla_base(header_bg=COLOR_SECUNDARIO, header_fg=colors.white, body_bg=colors.white):
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
        ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
        ("FONTSIZE", (0, 0), (-1, 0), 11),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 1), (-1, -1), body_bg),
        ("FONTNAME", (0, 1), (-1, -1), FONT_NAME_REGULAR),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ])


# ================= UTILIDADES DE RENDERIZADO =================
def FullImage(path):
    def draw(canvas, doc):
        w, h = A4
        canvas.drawImage(
            path,
            0,
            0,
            width=w,
            height=h,
            preserveAspectRatio=True,
            mask="auto"
        )
    return draw

def header_footer(canvas, doc):
    page_width, page_height = A4
    if os.path.exists("assets/header.png"):
        canvas.drawImage("assets/header.png", 0, page_height - 80, width=page_width, height=60, mask="auto")
    if os.path.exists("assets/footer.png"):
        canvas.drawImage("assets/footer.png", 0, 0, width=page_width, height=60, mask="auto")


# ================= DIBUJO DE GRÁFICOS Y TABLAS (PÁGINA 6) =================
def draw_grafico_edad(canvas, doc, grafico_edad_path):
    page_width, page_height = A4
    img_width = 220
    img = ImageReader(grafico_edad_path)
    img_w, img_h = img.getSize()
    img_height = img_width * img_h / img_w
    x = 40
    y = page_height - img_height - 120
    canvas.drawImage(grafico_edad_path, x, y, width=img_width, height=img_height, preserveAspectRatio=True, mask="auto")

def draw_grafico_escolaridad(canvas, grafico_path):
    page_width, page_height = A4
    img_width = 220
    img = ImageReader(grafico_path)
    img_w, img_h = img.getSize()
    img_height = img_width * img_h / img_w
    x = page_width / 2 + 10
    y = page_height - img_height - 320
    canvas.drawImage(grafico_path, x, y, width=img_width, height=img_height, preserveAspectRatio=True, mask="auto")

def draw_grafico_genero(canvas, grafico_path):
    page_width, page_height = A4
    img_width = 220
    img = ImageReader(grafico_path)
    img_w, img_h = img.getSize()
    img_height = img_width * img_h / img_w
    x = 40
    y = page_height - img_height - 540
    canvas.drawImage(grafico_path, x, y, width=img_width, height=img_height, preserveAspectRatio=True, mask="auto")

def draw_grafico_relacion(canvas, grafico_path):
    TITULO_X = 40
    TITULO_Y = 350
    CAJA_X = 35
    CAJA_Y = 80
    CAJA_WIDTH = 520
    CAJA_HEIGHT = 250
    GRAFICO_X = 60
    GRAFICO_Y = 90
    GRAFICO_WIDTH = 470
    GRAFICO_HEIGHT = 220

    canvas.setFont(FONT_NAME_BOLD, 16)
    canvas.setFillColor(COLOR_SECUNDARIO)
    canvas.drawString(TITULO_X, TITULO_Y, "Relación por distrito")

    canvas.setFillColor(colors.HexColor("#F7F9FC"))
    canvas.roundRect(CAJA_X, CAJA_Y, CAJA_WIDTH, CAJA_HEIGHT, 12, stroke=0, fill=1)
    canvas.drawImage(grafico_path, GRAFICO_X, GRAFICO_Y, width=GRAFICO_WIDTH, height=GRAFICO_HEIGHT, preserveAspectRatio=True, mask="auto")


# ================= FUNCIONES DE TABLAS ESTANDARIZADAS =================
def generar_tabla_estilizada(canvas, data, titulo, x, y, table_width, colores_filas):
    # --- 1. REDUCIMOS EL TAMAÑO DE LETRA (Antes estaban en 11 y 10) ---
    font_size_header = 10
    font_size_body = 9
    
    table_data = [[titulo, ""]]
    table_data.extend(data)
    
    table = Table(table_data, colWidths=[table_width * 0.6, table_width * 0.4])
    
    estilo = TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_HEADER_TABLA),
        ("TEXTCOLOR", (0, 0), (-1, 0), COLOR_TEXTO_OSCURO),
        ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
        ("FONTSIZE", (0, 0), (-1, 0), font_size_header),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 1), (-1, -1), font_size_body),
        ("GRID", (0, 1), (-1, -1), 0.5, colors.white),
        # --- 2. REDUCIMOS EL ESPACIO INTERNO (Antes estaba en 5) ---
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ])
    
    for i, color in enumerate(colores_filas, start=1):
        if i < len(table_data):
            estilo.add("BACKGROUND", (0, i), (-1, i), color)
            estilo.add("TEXTCOLOR", (0, i), (-1, i), colors.white)
            
    table.setStyle(estilo)
    table.wrapOn(canvas, table_width, 200)
    table.drawOn(canvas, x, y - table._height)

# ================= TABLA EDAD =================
def draw_tabla_edad(canvas, doc, tabla_edad):
    page_width, page_height = A4
    colores = [colors.HexColor(c) for c in ["#5B9BD5", "#A5A5A5", "#4472C4", "#255E91", "#B7B7B7"]]
    
    # Parámetros de posición
    pos_x = (page_width / 2) + 10
    pos_y = page_height - 180  # Ajustado para alinear con el gráfico de edad
    
    generar_tabla_estilizada(canvas, tabla_edad, "Participación por Edad", pos_x, pos_y, 185, colores)

# ================= TABLA ESCOLARIDAD =================
def draw_tabla_escolaridad(canvas, tabla_escolaridad):
    page_width, page_height = A4
    colores = [colors.HexColor(c) for c in ["#5B9BD5", "#A5A5A5", "#4472C4", "#255E91", "#B7B7B7", "#9DC3E6", "#8FAADC", "#424e69"]]
    
    # Parámetros de posición
    pos_x = 50
    pos_y = page_height - 340  # Ajustado para alinear con el gráfico de escolaridad
    
    generar_tabla_estilizada(canvas, tabla_escolaridad, "Participación por Escolaridad", pos_x, pos_y, 185, colores)

# ================= TABLA GENERO =================
def draw_tabla_genero(canvas, tabla_genero):
    page_width, page_height = A4
    colores = [colors.HexColor(c) for c in ["#5B9BD5", "#A5A5A5", "#4472C4"]]
    
    # Parámetros de posición
    pos_x = (page_width / 2) + 10
    pos_y = page_height - 600  # Ajustado para alinear con el gráfico de género
    
    generar_tabla_estilizada(canvas, tabla_genero, "Participación por Género", pos_x, pos_y, 185, colores)


def draw_tabla_simple(canvas, data, titulo, x, y, col_widths, header_color=COLOR_SECUNDARIO, font_size_header=11, font_size_body=10):
    TABLE_WIDTH = sum(col_widths)
    header_style = ParagraphStyle(name="HeaderSimple", fontName=FONT_NAME_BOLD, fontSize=font_size_header, textColor=colors.white)
    cell_style = ParagraphStyle(name="CellSimple", fontName=FONT_NAME_REGULAR, fontSize=font_size_body)
    
    table_data = []
    table_data.append([Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1))
    
    for row in data:
        nueva_fila = [Paragraph(str(cell), cell_style) for cell in row]
        table_data.append(nueva_fila)
        
    tabla = Table(table_data, colWidths=col_widths)
    tabla.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    
    tabla.wrapOn(canvas, TABLE_WIDTH, 400)
    tabla.drawOn(canvas, x, y - tabla._height)


def draw_tabla_victimizacion(canvas, data, titulo, x, y, col_widths, header_color):
    TABLE_WIDTH = sum(col_widths)
    header_style = ParagraphStyle(name="HeaderVict", fontName=FONT_NAME_BOLD, fontSize=10, alignment=TA_LEFT, textColor=colors.white)
    cell_style = ParagraphStyle(name="CellVict", fontName=FONT_NAME_REGULAR, fontSize=9, leading=11, alignment=TA_LEFT, wordWrap="CJK")
    
    table_data = []
    table_data.append([Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1))
    for row in data:
        table_data.append([Paragraph(str(cell), cell_style) for cell in row])
        
    tabla = Table(table_data, colWidths=col_widths)
    tabla.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    tabla.wrapOn(canvas, TABLE_WIDTH, 400)
    tabla.drawOn(canvas, x, y - tabla._height)


def draw_tabla_pareto(canvas, titulo, data, x, y, table_width=180, table_height_max=280, font_title=FONT_NAME_BOLD, font_size_title=13, header_color=COLOR_SECUNDARIO, body_color=colors.white, text_color=colors.black):
    total_filas = len(data)
    
    if total_filas <= 10:
        font_size, leading, pad = 11, 14, 5
    elif total_filas <= 15:
        font_size, leading, pad = 10, 12, 4
    elif total_filas <= 20:
        font_size, leading, pad = 9, 11, 3
    else:
        font_size, leading, pad = 8, 10, 2
        
    dynamic_style = ParagraphStyle(name="ParetoDynamic", fontName=FONT_NAME_REGULAR, fontSize=font_size, leading=leading, alignment=TA_LEFT, wordWrap="CJK")
    
    # Fondo y Título
    canvas.setFillColor(header_color)
    canvas.rect(x, y + 8, table_width, 24, stroke=0, fill=1)
    canvas.setFont(font_title, font_size_title)
    canvas.setFillColor(text_color)
    canvas.drawCentredString(x + table_width / 2, y + 14, titulo)
    
    # Tabla
    wrapped_data = [[Paragraph(str(row[0]), dynamic_style)] for row in data]
    table = Table(wrapped_data, colWidths=[table_width], rowHeights=None)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("BACKGROUND", (0, 0), (-1, -1), body_color),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), pad),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
    ]))
    table.wrapOn(canvas, table_width, table_height_max)
    table.drawOn(canvas, x, y - table._height)


def draw_porcentaje(canvas, texto, x, y, font=FONT_NAME_BOLD, size=18, color=colors.black):
    canvas.setFont(font, size)
    canvas.setFillColor(color)
    canvas.drawCentredString(x, y, texto)

def draw_cantidad(canvas, texto, x, y, font=FONT_NAME_BOLD, size=14, color=colors.black):
    canvas.setFont(font, size)
    canvas.setFillColor(color)
    canvas.drawCentredString(x, y, texto)


# ================= MICMAC Y TRIÁNGULO =================
def draw_micmac_lista(canvas, data, x, y, width=200, max_height=180, font_size=8):
    style = ParagraphStyle(name="MicmacCell", fontName=FONT_NAME_REGULAR, fontSize=font_size, leading=10, alignment=TA_LEFT, wordWrap="CJK")
    paragraphs = [Paragraph(item[0], style) for item in data]
    cols = [[], []]
    for i, p in enumerate(paragraphs):
        cols[i % 2].append(p)
        
    max_rows = max(len(cols[0]), len(cols[1]))
    table_data = []
    for i in range(max_rows):
        row = [
            cols[0][i] if i < len(cols[0]) else "",
            cols[1][i] if i < len(cols[1]) else ""
        ]
        table_data.append(row)
        
    table = Table(table_data, colWidths=[width / 2, width / 2], rowHeights=None)
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    table.wrapOn(canvas, width, max_height)
    table.drawOn(canvas, x, y - table._height)

def draw_tabla_overlay(canvas, data, x, y, width=120, font_size=8, max_height=160):
    style = ParagraphStyle(name="OverlayCell", fontName=FONT_NAME_REGULAR, fontSize=font_size, leading=font_size + 2, alignment=TA_LEFT, wordWrap="CJK")
    table_data = [[Paragraph(row[0], style)] for row in data]
    if not table_data:
        return
        
    table = Table(table_data, colWidths=[width])
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    table.wrapOn(canvas, width, max_height)
    table.drawOn(canvas, x, y - table._height)

def draw_texto_overlay(canvas, texto, x, y, font=FONT_NAME_BOLD, size=20, color=colors.white):
    canvas.setFont(font, size)
    canvas.setFillColor(color)
    canvas.drawCentredString(x, y, str(texto))

def draw_texto_mixto(canvas, x, y, texto_antes, valor_1, texto_medio, valor_2, texto_despues, width=250, font_size=11, valor_size=15):
    normal = ParagraphStyle("normal", fontName=FONT_NAME_REGULAR, fontSize=font_size, leading=14, alignment=TA_JUSTIFY)
    html = (
        f"{texto_antes} "
        f"<b><font size='{valor_size}'>{valor_1}</font></b> "
        f"{texto_medio} "
        f"<b><font size='{valor_size}'>{valor_2}</font></b> "
        f"{texto_despues}"
    )
    p = Paragraph(html, normal)
    p.wrapOn(canvas, width, 200)
    p.drawOn(canvas, x, y)

def draw_tabla_instituciones(canvas, data, titulo="Lista de Instituciones", x=60, y=250, table_width=480, header_bg=COLOR_SECUNDARIO, header_text_color=colors.white, body_bg=COLOR_FILA_ALTERNA, border_color=COLOR_BORDE, font_header=12, font_body=10):
    if not data:
        return
    table_data = [[titulo, ""]]
    table_data.extend(data)
    col_widths = [table_width * 0.6, table_width * 0.4]
    
    table = Table(table_data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), header_text_color),
        ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
        ("FONTSIZE", (0, 0), (-1, 0), font_header),
        ("GRID", (0, 1), (-1, -1), 0.5, border_color),
        ("BACKGROUND", (0, 1), (-1, -1), body_bg),
        ("FONTNAME", (0, 1), (-1, -1), FONT_NAME_REGULAR),
        ("FONTSIZE", (0, 1), (-1, -1), font_body),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    table.wrapOn(canvas, table_width, 300)
    table.drawOn(canvas, x, y - table._height)


# ================= TABLAS DE HORARIOS Y MODALIDADES =================
def draw_tabla_horario_distrito(canvas, data, titulo, x, y, col_widths):
    TABLE_WIDTH = sum(col_widths)
    header_style = ParagraphStyle(name="HorarioHeader", fontName=FONT_NAME_BOLD, fontSize=9, alignment=TA_CENTER, leading=11)
    cell_style = ParagraphStyle(name="HorarioCell", fontName=FONT_NAME_REGULAR, fontSize=8, alignment=TA_CENTER, leading=10, wordWrap="CJK")
    
    table_data = []
    table_data.append([Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1))
    
    for row in data:
        table_data.append([Paragraph(str(cell), cell_style) for cell in row])
        
    table = Table(table_data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARIO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E2F0D9")),
        ("FONTNAME", (0, 1), (-1, 1), FONT_NAME_BOLD),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))
    table.wrapOn(canvas, TABLE_WIDTH, 2000)
    table.drawOn(canvas, x, y - table._height)

def draw_tabla_modalidades_p14(canvas, data, titulo, x, y, col_widths):
    TABLE_WIDTH = sum(col_widths)
    header_style = ParagraphStyle(name="HeaderStyleP14", fontName=FONT_NAME_BOLD, fontSize=8, alignment=TA_CENTER, leading=10)
    distrito_style = ParagraphStyle(name="DistritoStyleP14", fontName=FONT_NAME_BOLD, fontSize=8, alignment=TA_LEFT, leading=9)
    cell_style = ParagraphStyle(name="CellStyleP14", fontName=FONT_NAME_REGULAR, fontSize=8, alignment=TA_CENTER, leading=9, wordWrap="CJK")
    
    table_data = []
    table_data.append([Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1))
    table_data.append([Paragraph(str(cell), header_style) for cell in data[0]])
    
    for row in data[1:]:
        nueva_fila = []
        for i, cell in enumerate(row):
            if i == 0:
                nueva_fila.append(Paragraph(str(cell), distrito_style))
            else:
                nueva_fila.append(Paragraph(str(cell), cell_style))
        table_data.append(nueva_fila)
        
    table = Table(table_data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARIO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E2FDD9")),
        ("TEXTCOLOR", (0, 1), (-1, 1), COLOR_TEXTO_OSCURO),
        ("BACKGROUND", (0, 2), (0, -1), colors.HexColor("#E2FDD9")),
        ("TEXTCOLOR", (0, 2), (0, -1), COLOR_TEXTO_OSCURO),
        ("BACKGROUND", (1, 2), (-1, -1), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("ALIGN", (1, 2), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    table.wrapOn(canvas, TABLE_WIDTH, 2000)
    table.drawOn(canvas, x, y - table._height)

def draw_tabla_dias_distritos_p15(canvas, data, titulo, x, y, col_widths):
    TABLE_WIDTH = sum(col_widths)
    header_style = ParagraphStyle(name="HeaderStyleP15", fontName=FONT_NAME_BOLD, fontSize=8, alignment=TA_CENTER, leading=10)
    distrito_style = ParagraphStyle(name="DistritoStyleP15", fontName=FONT_NAME_BOLD, fontSize=7, alignment=TA_LEFT, leading=9)
    cell_style = ParagraphStyle(name="CellStyleP15", fontName=FONT_NAME_REGULAR, fontSize=7, alignment=TA_CENTER, leading=9, wordWrap="CJK")
    
    table_data = []
    table_data.append([Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1))
    table_data.append([Paragraph(str(cell), header_style) for cell in data[0]])
    
    for row in data[1:]:
        nueva_fila = []
        for i, cell in enumerate(row):
            if i == 0:
                nueva_fila.append(Paragraph(str(cell), distrito_style))
            else:
                nueva_fila.append(Paragraph(str(cell), cell_style))
        table_data.append(nueva_fila)
        
    table = Table(table_data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARIO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E2FDD9")),
        ("TEXTCOLOR", (0, 1), (-1, 1), COLOR_TEXTO_OSCURO),
        ("BACKGROUND", (0, 2), (0, -1), colors.HexColor("#E2FDD9")),
        ("TEXTCOLOR", (0, 2), (0, -1), COLOR_TEXTO_OSCURO),
        ("BACKGROUND", (1, 2), (-1, -1), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("ALIGN", (1, 2), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    table.wrapOn(canvas, TABLE_WIDTH, 2000)
    table.drawOn(canvas, x, y - table._height)


# ================= LÍNEAS DE ACCIÓN =================
def normalizar_nombre(texto):
    texto = texto.lower()
    texto = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = texto.replace(" ", "_")
    texto = re.sub(r"[^a-z0-9_()]", "", texto)
    return texto

def construir_tabla_dinamica(titulo, items, ancho_total, estilo_header_color):
    total_items = len(items)
    if total_items <= 10:
        columnas, font_size, leading_size = 1, 10, 12
    elif total_items <= 18:
        columnas, font_size, leading_size = 2, 9, 11
    else:
        columnas, font_size, leading_size = 3, 8, 10

    CELDA_STYLE = ParagraphStyle(name="CeldaTabla", fontName=FONT_NAME_REGULAR, fontSize=font_size, leading=leading_size, alignment=TA_LEFT, wordWrap="CJK")
    
    data = [[titulo] + [""] * (columnas - 1) if columnas > 1 else [titulo]]
    
    if columnas == 1:
        for item in items:
            data.append([Paragraph(item[0], CELDA_STYLE)])
        col_widths = [ancho_total]
        
    elif columnas == 2:
        mitad = (len(items) + 1) // 2
        col1 = items[:mitad]
        col2 = items[mitad:]
        while len(col2) < len(col1):
            col2.append([""])
        for i in range(len(col1)):
            p1 = Paragraph(col1[i][0], CELDA_STYLE) if col1[i][0] else ""
            p2 = Paragraph(col2[i][0], CELDA_STYLE) if col2[i][0] else ""
            data.append([p1, p2])
        col_widths = [ancho_total / 2, ancho_total / 2]
        
    else:
        tercio = (len(items) + 2) // 3
        col1 = items[:tercio]
        col2 = items[tercio:tercio * 2]
        col3 = items[tercio * 2:]
        while len(col2) < len(col1): col2.append([""])
        while len(col3) < len(col1): col3.append([""])
        for i in range(len(col1)):
            p1 = Paragraph(col1[i][0], CELDA_STYLE) if col1[i][0] else ""
            p2 = Paragraph(col2[i][0], CELDA_STYLE) if col2[i][0] else ""
            p3 = Paragraph(col3[i][0], CELDA_STYLE) if col3[i][0] else ""
            data.append([p1, p2, p3])
        col_widths = [ancho_total / 3, ancho_total / 3, ancho_total / 3]

    tabla = Table(data, colWidths=col_widths)
    tabla.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), estilo_header_color),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
        ("FONTSIZE", (0, 0), (-1, 0), 12),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("TEXTCOLOR", (0, 1), (-1, -1), COLOR_TEXTO_OSCURO),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COLOR_FILA_ALTERNA]),
    ]))
    return tabla


def draw_pagina_linea_accion(canvas, doc, linea, vectores_path="assets/vectores"):
    page_width, page_height = A4
    SUBTITULO_FONT = FONT_NAME_BOLD
    SUBTITULO_SIZE = 14
    SUBTITULO_COLOR = COLOR_SECUNDARIO
    TITULO_FONT = FONT_NAME_BOLD
    TITULO_SIZE = 13
    TITULO_COLOR = COLOR_SECUNDARIO

    canvas.setFont(SUBTITULO_FONT, SUBTITULO_SIZE)
    canvas.setFillColor(SUBTITULO_COLOR)
    canvas.drawString(80, page_height - 70, "Líneas de Acción")
    
    titulo_style = ParagraphStyle(
        name="TituloLineaInterna",
        fontName=TITULO_FONT,
        fontSize=TITULO_SIZE,
        leading=TITULO_SIZE + 2,
        textColor=TITULO_COLOR,
        alignment=0,
        wordWrap="CJK"
    )
    
    texto_titulo = "<br/>".join(linea["problematicas"])
    p = Paragraph(texto_titulo, titulo_style)
    w, h = p.wrap(page_width * 0.90, 150)
    
    VECTOR_WIDTH = 150
    VECTOR_HEIGHT = 150
    VECTOR_Y = page_height - 280
    VECTOR_SPACING = 30
    total_vectores = len(linea["problematicas"])
    
    if total_vectores > 0:
        total_width = total_vectores * VECTOR_WIDTH + (total_vectores - 1) * VECTOR_SPACING
        start_x = (page_width - total_width) / 2
        for i, problema in enumerate(linea["problematicas"]):
            nombre = normalizar_nombre(problema) + ".png"
            ruta = os.path.join(vectores_path, nombre)
            if os.path.exists(ruta):
                canvas.drawImage(ruta, start_x + i * (VECTOR_WIDTH + VECTOR_SPACING), VECTOR_Y, width=VECTOR_WIDTH, height=VECTOR_HEIGHT, preserveAspectRatio=True, mask="auto")
                
    p.drawOn(canvas, 30, page_height - 78 - h)
    
    TOTAL_WIDTH = 90
    TOTAL_HEIGHT = 90
    TOTAL_X = page_width - TOTAL_WIDTH - 40
    TOTAL_Y = page_height - 170
    if os.path.exists("assets/total.png"):
        canvas.drawImage("assets/total.png", TOTAL_X, TOTAL_Y, width=TOTAL_WIDTH, height=TOTAL_HEIGHT, preserveAspectRatio=True, mask="auto")
        
    canvas.setFont(FONT_NAME_BOLD, 16)
    canvas.setFillColor(colors.white)
    texto_total = linea.get("total_porcentaje", "0.00%")
    text_width = canvas.stringWidth(texto_total, FONT_NAME_BOLD, 16)
    canvas.drawString(TOTAL_X + (TOTAL_WIDTH - text_width) / 2, TOTAL_Y + TOTAL_HEIGHT / 2 - 8 + 10, texto_total)

    tabla_c = construir_tabla_dinamica("Causas Socio Culturales y Estructurales", linea["causas"], page_width - 80, COLOR_PRIMARIO)
    tabla_c.wrapOn(canvas, page_width - 80, 400)
    
    tabla_p = construir_tabla_dinamica("Problemáticas Influyentes", linea["problemas_influyentes"], page_width - 80, COLOR_SECUNDARIO)
    tabla_p.wrapOn(canvas, page_width - 80, 400)
    
    y_tabla_c = page_height - 280 - tabla_c._height
    y_tabla_p = y_tabla_c - 20 - tabla_p._height
    tabla_c.drawOn(canvas, 40, y_tabla_c)
    tabla_p.drawOn(canvas, 40, y_tabla_p)


def draw_pagina_linea_accion_detalle(canvas, doc, linea):
    page_width, page_height = A4
    MARGEN_X = 40
    ANCHO_UTIL = page_width - 80
    OBJ_Y = page_height - 120
    ESPACIO_BLOQUES = 30
    
    canvas.setFont(FONT_NAME_BOLD, 14)
    canvas.setFillColor(COLOR_SECUNDARIO)
    canvas.drawString(MARGEN_X, OBJ_Y, "Objetivo general de la intervención:")
    
    objetivo_texto = "Optimizar la identificación y respuesta a reincidentes en situación de calle, mejorando la eficacia policial y reduciendo los delitos para aumentar la seguridad comunitaria."
    estilo_obj = ParagraphStyle(name="objetivo", fontName=FONT_NAME_REGULAR, fontSize=11, leading=15, alignment=TA_LEFT)
    p_obj = Paragraph(objetivo_texto, estilo_obj)
    w, h = p_obj.wrap(ANCHO_UTIL, 200)
    p_obj.drawOn(canvas, MARGEN_X, OBJ_Y - 20 - h)
    
    current_y = OBJ_Y - 20 - h - ESPACIO_BLOQUES
    
    if linea["lider_estrategico"]:
        estilo_lider = ParagraphStyle(name="LiderCell", fontName=FONT_NAME_BOLD, fontSize=16, leading=18, alignment=TA_LEFT, wordWrap="CJK")
        data_lider = [[Paragraph("Líder Estratégico", estilo_lider), Paragraph(linea["lider_estrategico"], estilo_lider)]]
        tabla_lider = Table(data_lider, colWidths=[ANCHO_UTIL * 0.4, ANCHO_UTIL * 0.6])
        tabla_lider.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), COLOR_PRIMARIO),
            ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#E2FDD9")),
            ("GRID", (0, 0), (-1, -1), 0.5, COLOR_SECUNDARIO),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        tabla_lider.wrapOn(canvas, ANCHO_UTIL, 200)
        tabla_lider.drawOn(canvas, MARGEN_X, current_y - tabla_lider._height)
        current_y = current_y - tabla_lider._height - ESPACIO_BLOQUES
        
    if linea["acciones"]:
        estilo_header = ParagraphStyle(name="HeaderAcciones", fontName=FONT_NAME_BOLD, fontSize=11, textColor=colors.white)
        estilo_celda = ParagraphStyle(name="CeldaAcciones", fontName=FONT_NAME_REGULAR, fontSize=10, leading=13, alignment=TA_LEFT, wordWrap="CJK")
        acciones_data = [[Paragraph("Acciones estratégicas", estilo_header)]]
        for a in linea["acciones"]:
            acciones_data.append([Paragraph("• " + a, estilo_celda)])
            
        tabla_acciones = Table(acciones_data, colWidths=[ANCHO_UTIL])
        tabla_acciones.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), COLOR_PRIMARIO),
            ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        tabla_acciones.wrapOn(canvas, ANCHO_UTIL, 400)
        tabla_acciones.drawOn(canvas, MARGEN_X, current_y - tabla_acciones._height)
        current_y = current_y - tabla_acciones._height - ESPACIO_BLOQUES
        
    if linea["cogestores"]:
        mitad = len(linea["cogestores"]) // 2
        col1 = linea["cogestores"][:mitad]
        col2 = linea["cogestores"][mitad:]
        filas = [["Cogestores", ""]]
        max_len = max(len(col1), len(col2))
        for i in range(max_len):
            c1 = "• " + str(col1[i]) if i < len(col1) else ""
            c2 = "• " + str(col2[i]) if i < len(col2) else ""
            filas.append([c1, c2])
            
        tabla_cog = Table(filas, colWidths=[ANCHO_UTIL / 2, ANCHO_UTIL / 2])
        tabla_cog.setStyle(TableStyle([
            ("SPAN", (0, 0), (-1, 0)),
            ("BACKGROUND", (0, 0), (-1, 0), COLOR_SECUNDARIO),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
            ("FONTSIZE", (0, 0), (-1, 0), 11),
            ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        tabla_cog.wrapOn(canvas, ANCHO_UTIL, 400)
        tabla_cog.drawOn(canvas, MARGEN_X, current_y - tabla_cog._height)


# ================= PERCEPCIÓN CIUDADANA =================
def draw_pagina_percepcion_1(canvas, doc, grafico_actual_path, grafico_comparacion_path, tabla_percepcion):
    page_width, page_height = A4
    canvas.setFont(FONT_NAME_BOLD, 18)
    canvas.setFillColor(COLOR_SECUNDARIO)
    canvas.drawString(40, page_height - 110, "¿Se siente seguro en su comunidad?")
    
    GRAFICO_Y = page_height - 380
    GRAFICO_WIDTH = 230
    
    canvas.setFont(FONT_NAME_BOLD, 14)
    canvas.drawString(40, GRAFICO_Y + 230 + 10, "Actualmente")
    canvas.drawImage(grafico_actual_path, 50, GRAFICO_Y + 10, width=250, height=250, preserveAspectRatio=True, mask="auto")
    
    canvas.drawString(page_width - GRAFICO_WIDTH - 40, GRAFICO_Y + 230 + 10, "Comparación año anterior")
    canvas.drawImage(grafico_comparacion_path, page_width - GRAFICO_WIDTH - 40, GRAFICO_Y, width=GRAFICO_WIDTH, height=230, preserveAspectRatio=True, mask="auto")
    
    canvas.setFont(FONT_NAME_BOLD, 16)
    canvas.drawString(40, 420, "Comparativo Percepción Ciudadana por Zonas")
    
    header_style = ParagraphStyle(name="PercepcionHeader", fontName=FONT_NAME_BOLD, fontSize=9, alignment=TA_CENTER, textColor=colors.white, wordWrap="CJK")
    cell_style = ParagraphStyle(name="PercepcionCell", fontName=FONT_NAME_REGULAR, fontSize=9, alignment=TA_CENTER, wordWrap="CJK")
    
    tabla_render = []
    for row_idx, row in enumerate(tabla_percepcion):
        nueva = [Paragraph(str(cell), header_style) if row_idx <= 1 else Paragraph(str(cell), cell_style) for cell in row]
        tabla_render.append(nueva)
        
    tabla = Table(tabla_render, colWidths=[130, 55, 55, 55, 55, 55, 55], repeatRows=2)
    tabla.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("BACKGROUND", (0, 0), (-1, 1), COLOR_PRIMARIO),
        ("BACKGROUND", (0, 2), (-1, -1), colors.white),
        ("BACKGROUND", (0, 2), (0, -1), colors.HexColor("#E2F0D9")),
        ("FONTNAME", (0, 2), (0, -1), FONT_NAME_BOLD),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    
    if len(tabla_percepcion) >= 2:
        tabla_percepcion[0][1] = "Seguro"
        tabla_percepcion[0][2] = "Inseguro"
        tabla_percepcion[0][3] = "No existe"
        tabla_percepcion[1][1] = ""
        tabla_percepcion[1][2] = ""
        tabla_percepcion[1][3] = ""
        
    tabla.wrapOn(canvas, page_width - 80, 500)
    tabla.drawOn(canvas, 40, 420 - tabla._height - 10)


def draw_pagina_percepcion_2(canvas, doc, grafico_victimizacion_path, grafico_no_denuncia_path, tabla_no_denuncia, motivo_principal, total_omitidas):
    page_width, page_height = A4
    canvas.setFont(FONT_NAME_BOLD, 20)
    canvas.setFillColor(COLOR_SECUNDARIO)
    canvas.drawString(40, page_height - 100, "Victimización Ciudadana")
    
    canvas.setFont(FONT_NAME_BOLD, 14)
    canvas.drawString(40, page_height - 140, "¿Usted ha sido víctima de algún delito en los últimos 12 meses?")
    canvas.drawString(40, page_height - 160, "¿Denunció ante el O.I.J?")
    canvas.drawImage(grafico_victimizacion_path, 40, page_height - 420, width=page_width - 80, height=250, preserveAspectRatio=True, mask="auto")
    
    canvas.setFont(FONT_NAME_BOLD, 14)
    canvas.drawString(40, page_height - 460, "¿Por qué la población que ha sido víctima no denuncia?")
    canvas.drawImage(grafico_no_denuncia_path, 40, page_height - 700, width=page_width - 80, height=230, preserveAspectRatio=True, mask="auto")
    
    draw_tabla_victimizacion(canvas, tabla_no_denuncia, "Detalle motivos de no denuncia", 60, 120, [170, 70], COLOR_PRIMARIO)
    
    texto = f"La mayor cantidad de los encuestados que fueron víctimas de algún delito, señalan que no denuncian, debido a {motivo_principal}. Total respuestas omitidas: {total_omitidas}."
    estilo = ParagraphStyle(name="textoVict", fontName=FONT_NAME_REGULAR, fontSize=11, leading=14, alignment=TA_JUSTIFY)
    p = Paragraph(texto, estilo)
    p.wrapOn(canvas, page_width * 0.4, 200)
    p.drawOn(canvas, 300, 70)


# ================= GENERADOR PRINCIPAL (Mantiene comportamiento exacto) =================
def generar_pdf(
    portada_path, grafico_relacion_path, grafico_edad_path, grafico_escolaridad_path,
    grafico_genero_path, delegacion, codigo, tabla_participacion, tabla_edad,
    tabla_escolaridad, tabla_genero, tabla_encuesta_comunidad, tabla_otras_encuestas,
    datos_pagina_8, datos_pagina_9, tabla_delitos, tabla_riesgos, porcentaje_delitos,
    porcentaje_riesgos, cantidad_delitos, cantidad_riesgos, micmac_poder,
    micmac_conflicto, micmac_autonomas, micmac_resultados, tabla_riesgos_micmac2,
    tabla_delitos_micmac2, cantidad_problematicas, riesgos_total, delitos_total,
    causas_identificadas, factores_micmac, triangulo_directa, triangulo_sociocultural,
    triangulo_estructural, tabla_instituciones, grafico_denuncias_path,
    tabla_denuncias, total_denuncias, grafico_horario_path, tabla_horario, total_am,
    total_pm, tabla_horario_distrito, grafico_p14_path, tabla_p14, grafico_p15_path,
    tabla_p15, total_lineas, lineas_municipalidad, lineas_fp, lineas_mixtas,
    logo_muni_path, lineas_accion_data, grafico_percepcion_actual_path,
    grafico_percepcion_comparacion_path, tabla_percepcion, grafico_victimizacion_path,
    grafico_no_denuncia_path, tabla_no_denuncia, motivo_principal, total_omitidas,
    grafico_horarios_percepcion, grafico_armas_percepcion, tabla_horarios_percepcion,
    tabla_armas, horario_mayor, metodo_mas_usado, omitidas_aportes,
    grafico_servicio_policial, grafico_servicio_anual, grafico_conoce_policia,
    grafico_atencion, tabla_servicio, tabla_servicio_anual, tabla_conoce,
    tabla_atencion, omitidas_servicio, total_respuestas_servicio,
    grafico_comercio_seguridad, grafico_comercio_programa, grafico_comercio_inscrito,
    grafico_comercio_contacto,
):
    buffer = BytesIO()
    styles = getSampleStyleSheet()
    styles["Heading1"].textColor = COLOR_SECUNDARIO
    
    styles.add(ParagraphStyle(name="NormalJustificado", parent=styles["Normal"], alignment=TA_JUSTIFY, leading=14))
    styles.add(ParagraphStyle(name="TituloGrande", fontName=FONT_NAME_REGULAR, fontSize=26, textColor=colors.white, leading=30, spaceAfter=2, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="TituloDelta", fontName=FONT_NAME_BOLD, fontSize=45, textColor=COLOR_PRIMARIO, leading=40, spaceAfter=10, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="TituloD2", fontName=FONT_NAME_BOLD, fontSize=60, textColor=colors.white, leading=60, spaceAfter=10, alignment=TA_LEFT))

    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    story = []

    # -- PORTADA (Página 1) --
    story.append(PageBreak())
    story.append(Spacer(1, 80))
    story.append(Paragraph("DELEGACIÓN POLICIAL", styles["TituloGrande"]))
    story.append(Paragraph(delegacion, styles["TituloDelta"]))
    story.append(Paragraph(codigo, styles["TituloD2"]))

    # -- PÁGINA 2 --
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Introducción", styles["Heading1"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "Desde el año 2022, el Ministerio de Seguridad Pública ha implementado en todo el territorio nacional el Modelo Preventivo de Gestión Policial, una iniciativa estratégica destinada a fortalecer la seguridad pública a través de un enfoque proactivo y colaborativo. Una parte integral de este modelo es la Estrategia Integral de Prevención para la Seguridad Pública, conocida como Sembremos Seguridad, que se centra en la contextualización de las dinámicas delincuenciales y sociales que afectan a nuestras comunidades.",
        styles["NormalJustificado"]
    ))
    story.append(Spacer(1, 20))
    story.append(Paragraph(
        "El presente informe, elaborado para el territorio asignado, surge como una herramienta esencial para la toma efectiva de decisiones. Este informe se concibe como un instrumento dinámico y orientado hacia el futuro, diseñado para proporcionar información clave y un plan de trabajo estructurado que permita abordar las problemáticas prioritarias identificadas en el ámbito de la seguridad pública.",
        styles["NormalJustificado"]
    ))
    story.append(Spacer(1, 40))
    if os.path.exists("assets/conformacion.png"):
        story.append(Image("assets/conformacion.png", width=500, height=333))

    # -- PÁGINAS FIJAS 3-7 --
    story.append(PageBreak())
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Datos de Participación", styles["Heading1"]))
    story.append(Spacer(1, 20))
    story.append(Paragraph("Participación por Distrito", styles["Heading1"]))
    story.append(Spacer(1, 10))
    
    tabla_part = Table(tabla_participacion, colWidths=[170, 170, 120])
    tabla_part.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_SECUNDARIO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
        ("FONTSIZE", (0, 0), (-1, 0), 12),
        ("FONTSIZE", (0, 1), (-1, -1), 11),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("BACKGROUND", (1, 1), (-1, -1), colors.HexColor("#E2FDD9"))
    ]))
    story.append(tabla_part)
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Datos de Participación", styles["Heading1"]))
    story.append(Spacer(1, 410))
    story.append(PageBreak())
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Proceso Metodológico", styles["Heading1"]))
    story.append(Paragraph("Información demográfica según zona asignada a la Delegación Policial", styles["Normal"]))
    story.append(Spacer(1, 190))
    if os.path.exists("assets/netquest.png"):
        story.append(Image("assets/netquest.png", width=550, height=125))
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Diagrama de Pareto", styles["Heading1"]))
    story.append(Paragraph("(Aplicando el principio de 80/20 donde el 80% es lo trivial y el 20% es lo vital)", styles["Normal"]))
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("MICMAC", styles["Heading1"]))
    story.append(Paragraph("(Matriz de Impactos Cruzado - Multiplicación Aplicada a un Clasificación)", styles["Normal"]))
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Triángulo de las Violencias", styles["Heading1"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Es una metodología también llamada 'teoría de los conflictos' creada por el sociólogo y matemático Johan Galtung, que permite estudiar y transformar los conflictos mediante la identificación de variantes de la violencia (Directa, cultural y estructural), visualizando las causas generadoras de las problemáticas.", styles["NormalJustificado"]))
    story.append(Spacer(1, 200))
    story.append(Spacer(1, 20))
    story.append(Paragraph("Proceso Metodológico", styles["Heading1"]))
    story.append(Paragraph("Lista de Instituciones participantes en calificación de procesos: MIC-MAC y Triángulo de las Violencias", styles["Heading2"]))
    
    # -- PÁGINAS ESTADÍSTICAS --
    story.append(PageBreak())
    story.append(PageBreak())
    story.append(Spacer(1, 30))
    story.append(Paragraph("Denuncias por distrito", styles["Heading2"]))
    story.append(Spacer(1, 220))
    story.append(Paragraph("Denuncias por rango horario", styles["Heading2"]))
    story.append(PageBreak())
    story.append(Spacer(1, 30))
    story.append(Paragraph("Denuncias por Modalidad en el Cantón", styles["Heading2"]))
    story.append(Spacer(1, 280))
    story.append(PageBreak())
    story.append(Spacer(1, 30))
    story.append(Paragraph("Denuncias por días de la semana en el Cantón", styles["Heading2"]))
    story.append(Spacer(1, 280))
    story.append(PageBreak())
    story.append(PageBreak())
    
    # -- LÍNEAS DE ACCIÓN INTRO --
    story.append(Spacer(1, 40))
    story.append(Paragraph("Líneas de Acción", styles["Heading1"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "El procedimiento de construcción de líneas de acción en la estrategia SEMBREMOS SEGURIDAD es esencial en el ámbito local. Este proceso comienza con la recolección y análisis de datos específicos del territorio, utilizando herramientas metodológicas científicas para identificar y contextualizar las problemáticas locales. Posteriormente, un grupo de expertos de diversas instituciones evalúa las causas subyacentes y propone soluciones viables. Estos pasos se consolidan en el siguiente apartado con la finalidad de plasmar de manera transparente cuáles serán las acciones específicas para la atención de las problemáticas priorizadas.",
        styles["NormalJustificado"]
    ))
    story.append(Spacer(1, 20))
    story.append(Paragraph(
        "Los coordinadores estratégicos que desempeñan un papel fundamental en este procedimiento, son el Gobierno Local por su rol de autoridad local y Fuerza Pública como ente competente para la prevención del delito. Estos son responsables de garantizar la trazabilidad y el cumplimiento de los indicadores previamente consensuados por los actores sociales.",
        styles["NormalJustificado"]
    ))
    story.append(Spacer(1, 20))
    story.append(Paragraph(
        "Estos líderes estratégicos aseguran que cada etapa del plan se lleve a cabo de manera efectiva y que las intervenciones estén alineadas con los objetivos planteados. Su liderazgo y supervisión continua son cruciales para el éxito de la estrategia SEMBREMOS SEGURIDAD.",
        styles["NormalJustificado"]
    ))

    # Generador páginas dinámicas LA
    for i in range(total_lineas):
        story.append(PageBreak())
        story.append(Spacer(1, 1))
        story.append(PageBreak())
        story.append(Spacer(1, 1))
        story.append(PageBreak())
        story.append(Spacer(1, 1))

    # Generador páginas Percepción
    for _ in range(6):
        story.append(PageBreak())
        story.append(Spacer(1, 1))
    
    # Final
    story.append(PageBreak())
    story.append(Spacer(1, 1))

    # ================= LOGICA ON LATER PAGES =================
    def first_page(canvas, doc):
        FullImage(portada_path)(canvas, doc)

    def later_pages(canvas, doc):
        if doc.page == 2:
            if os.path.exists("assets/intro.png"): FullImage("assets/intro.png")(canvas, doc)
        elif doc.page == 3:
            header_footer(canvas, doc)
        elif doc.page == 4:
            if os.path.exists("assets/participacion.png"): FullImage("assets/participacion.png")(canvas, doc)
        elif doc.page == 5:
            header_footer(canvas, doc)
            draw_grafico_relacion(canvas, grafico_relacion_path)
        elif doc.page == 6:
            header_footer(canvas, doc)
            draw_grafico_edad(canvas, doc, grafico_edad_path)
            draw_tabla_edad(canvas, doc, tabla_edad)
            draw_grafico_escolaridad(canvas, grafico_escolaridad_path)
            draw_tabla_escolaridad(canvas, tabla_escolaridad)
            draw_grafico_genero(canvas, grafico_genero_path)
            draw_tabla_genero(canvas, tabla_genero)
        elif doc.page == 7:
            if os.path.exists("assets/metodologico.png"): FullImage("assets/metodologico.png")(canvas, doc)
        elif doc.page == 8:
            header_footer(canvas, doc)
            draw_tabla_simple(canvas, tabla_encuesta_comunidad, "Encuesta a Comunidad", 90, A4[1] - 150, [100, 100, 100, 100], COLOR_SECUNDARIO)
            draw_tabla_simple(canvas, tabla_otras_encuestas, "Otras encuestas", 90, A4[1] - 250, [100, 100, 100, 100], COLOR_SECUNDARIO)
            if os.path.exists("assets/datos.png"):
                canvas.drawImage("assets/datos.png", 0, 20, width=600, height=400, preserveAspectRatio=True, mask="auto")
            canvas.setFont(FONT_NAME_BOLD, 22)
            canvas.setFillColor(colors.white)
            canvas.drawString(110, 265, str(datos_pagina_8.get("encuesta_comunidad", "")))
            canvas.drawString(80, 140, str(datos_pagina_8.get("encuesta_policial", "")))
            canvas.drawString(400, 265, str(datos_pagina_8.get("encuesta_comercio", "")))
            canvas.drawString(430, 140, str(datos_pagina_8.get("estadistica", "")))
            canvas.setFont(FONT_NAME_BOLD, 28)
            canvas.drawString(270, 140, str(datos_pagina_8.get("total_datos", "")))
        elif doc.page == 9:
            header_footer(canvas, doc)
            img_width, img_height = 550, 300
            img_x, img_y = (A4[0] - img_width) / 2, A4[1] - img_height - 70
            if os.path.exists("assets/pareto.png"):
                canvas.drawImage("assets/pareto.png", img_x, img_y, width=img_width, height=img_height, preserveAspectRatio=True, mask="auto")
            canvas.setFont(FONT_NAME_BOLD, 20)
            canvas.setFillColor(colors.white)
            canvas.drawString(img_x + 150, img_y + img_height - 155, datos_pagina_9.get("lado_izquierdo", ""))
            canvas.drawString(img_x + img_width - 100, img_y + img_height - 115, datos_pagina_9.get("derecha_superior", ""))
            canvas.drawString(img_x + img_width - 100, img_y + 105, datos_pagina_9.get("derecha_inferior", ""))
            
            x_izq, x_der = 60, A4[0] / 2 + 20
            y_base = A4[1] - 380
            draw_porcentaje(canvas, porcentaje_delitos, x_izq + 90, y_base + 45)
            draw_tabla_pareto(canvas, "Delitos", tabla_delitos, x_izq, y_base, header_color=COLOR_PRIMARIO)
            draw_cantidad(canvas, f"Total: {cantidad_delitos}", x_izq + 90, y_base - 365)
            
            draw_porcentaje(canvas, porcentaje_riesgos, x_der + 90, y_base + 45)
            draw_tabla_pareto(canvas, "Riesgos Sociales", tabla_riesgos, x_der, y_base, header_color=COLOR_PRIMARIO)
            draw_cantidad(canvas, f"Total: {cantidad_riesgos}", x_der + 90, y_base - 365)
            
        elif doc.page == 10:
            header_footer(canvas, doc)
            img_width, img_height = 650, 420
            img_x, img_y = (A4[0] - img_width) / 2, A4[1] - img_height - 115
            if os.path.exists("assets/micmac.png"):
                canvas.drawImage("assets/micmac.png", img_x, img_y, width=img_width, height=img_height, preserveAspectRatio=True, mask="auto")
            
            quad_w = img_width / 2 - 80
            x_left, x_right = img_x + 90, img_x + img_width / 2 + 5
            y_top, y_bottom = img_y + img_height - 70, img_y + img_height / 2 - 20
            
            draw_micmac_lista(canvas, micmac_poder, x_left, y_top, quad_w)
            draw_micmac_lista(canvas, micmac_conflicto, x_right, y_top, quad_w)
            draw_micmac_lista(canvas, micmac_autonomas, x_left, y_bottom, quad_w)
            draw_micmac_lista(canvas, micmac_resultados, x_right, y_bottom, quad_w)
            
            # Parte 2 MICMAC
            img_width2, img_height2 = 520, 260
            img_x2, img_y2 = (A4[0] - img_width2) / 2, 50
            if os.path.exists("assets/micmac2.png"):
                canvas.drawImage("assets/micmac2.png", img_x2, img_y2, width=img_width2, height=img_height2, preserveAspectRatio=True, mask="auto")
            
            tabla_width, tabla_font = 130, 8
            x_tabla_riesgos = img_x2 + img_width2 * 0.55 - tabla_width / 2
            x_tabla_delitos = img_x2 + img_width2 * 0.85 - tabla_width / 2
            y_tablas = img_y2 + img_height2 - 80
            
            draw_tabla_overlay(canvas, tabla_riesgos_micmac2, x_tabla_riesgos, y_tablas, width=tabla_width, font_size=tabla_font)
            draw_tabla_overlay(canvas, tabla_delitos_micmac2, x_tabla_delitos, y_tablas, width=tabla_width, font_size=tabla_font)
            
            x_centro, y_centro = img_x2 + img_width2 / 4, img_y2 + img_height2 / 2
            draw_texto_overlay(canvas, cantidad_problematicas, x_centro - 30, y_centro - 15, size=30)
            draw_texto_overlay(canvas, riesgos_total, x_centro + 155, y_centro + 80, size=18)
            draw_texto_overlay(canvas, delitos_total, x_centro + 290, y_centro + 80, size=18)

        elif doc.page == 11:
            header_footer(canvas, doc)
            draw_texto_mixto(canvas, 45, A4[1] - 350, "Frente a lo anterior, esta metodología permitió la identificación de", causas_identificadas, "causas, directamente relacionadas con los", factores_micmac, "factores priorizados en la Mic-Mac.")
            
            img_width, img_height = 260, 260
            img_x, img_y = A4[0] / 2 + 10, A4[1] - img_height - 130
            if os.path.exists("assets/triangulo.png"):
                canvas.drawImage("assets/triangulo.png", img_x, img_y, width=img_width, height=img_height, preserveAspectRatio=True, mask="auto")
            
            img_width, img_height = 260, 260
            img_x, img_y = A4[0] / 2 + 10, A4[1] - img_height - 130
            if os.path.exists("assets/triangulo.png"):
                canvas.drawImage("assets/triangulo.png", img_x, img_y, width=img_width, height=img_height, preserveAspectRatio=True, mask="auto")
            
            canvas.setFont(FONT_NAME_BOLD, 18)
            canvas.setFillColor(colors.black)
            canvas.drawCentredString(img_x + img_width / 2, img_y + img_height - 50, str(triangulo_directa))
            canvas.drawCentredString(img_x + 40, img_y + 45, str(triangulo_sociocultural))
            canvas.drawCentredString(img_x + img_width - 40, img_y + 45, str(triangulo_estructural))
            
            draw_tabla_instituciones(canvas, tabla_instituciones, x=60, y=360, table_width=480, header_bg=COLOR_PRIMARIO)

        elif doc.page == 12:
            if os.path.exists("assets/estadistica.png"): FullImage("assets/estadistica.png")(canvas, doc)

        elif doc.page == 13:
            header_footer(canvas, doc)
            canvas.drawImage(grafico_denuncias_path, x=(A4[0] - 550) / 2, y=A4[1] - 330, width=250, height=250, preserveAspectRatio=True, mask="auto")
            draw_tabla_simple(canvas, tabla_denuncias, "Detalle de denuncias por distrito", 347, A4[1] - 80, [130], COLOR_SECUNDARIO)
            
            canvas.setFillColor(COLOR_SECUNDARIO)
            canvas.rect(A4[0] / 2 - 75, 510, 115, 50, fill=1, stroke=0)
            canvas.setFillColor(colors.white)
            canvas.setFont(FONT_NAME_BOLD, 11)
            canvas.drawCentredString(A4[0] / 2 - 18, 540, "Total denuncias")
            canvas.setFont(FONT_NAME_BOLD, 22)
            canvas.drawCentredString(A4[0] / 2 - 18, 518, str(total_denuncias))
            
            canvas.drawImage(grafico_horario_path, x=(A4[0] - 550) / 2, y=A4[1] - 600, width=250, height=250, preserveAspectRatio=True, mask="auto")
            draw_tabla_simple(canvas, tabla_horario, "Denuncias por horario", 420, A4[1] - 350, [90, 40], COLOR_SECUNDARIO)
            
            canvas.setFillColor(COLOR_SECUNDARIO)
            canvas.rect(280, 360, 100, 40, fill=1, stroke=0)
            canvas.setFillColor(colors.white)
            canvas.setFont(FONT_NAME_BOLD, 12)
            canvas.drawCentredString(330, 385, "AM")
            canvas.setFont(FONT_NAME_BOLD, 16)
            canvas.drawCentredString(330, 368, str(total_am))
            
            canvas.setFillColor(COLOR_SECUNDARIO)
            canvas.rect(280, 300, 100, 40, fill=1, stroke=0)
            canvas.setFillColor(colors.white)
            canvas.setFont(FONT_NAME_BOLD, 12)
            canvas.drawCentredString(330, 325, "PM")
            canvas.setFont(FONT_NAME_BOLD, 16)
            canvas.drawCentredString(330, 308, str(total_pm))
            
            total_columnas = len(tabla_horario_distrito[0])
            draw_tabla_horario_distrito(canvas, tabla_horario_distrito, "DCLP según horario, por distrito", 40, 260, [(A4[0] - 80) / total_columnas] * total_columnas)
            
            if os.path.exists("assets/horas.png"):
                canvas.drawImage("assets/horas.png", 295, 410, 70, 70)

        elif doc.page == 14:
            header_footer(canvas, doc)
            POS_X, POS_Y = (A4[0] - 500) / 2, A4[1] - 300 - 110
            canvas.drawImage(grafico_p14_path, POS_X, POS_Y, width=500, height=300, preserveAspectRatio=True, mask="auto")
            
            total_col = len(tabla_p14[0])
            draw_tabla_modalidades_p14(canvas, tabla_p14, "Frecuencia de modalidades por distrito", 30, POS_Y - 30, [(A4[0] - 60) / total_col] * total_col)

        elif doc.page == 15:
            header_footer(canvas, doc)
            POS_X, POS_Y = (A4[0] - 450) / 2, A4[1] - 300 - 100
            canvas.drawImage(grafico_p15_path, POS_X, POS_Y, width=450, height=300, preserveAspectRatio=True, mask="auto")
            
            total_col = len(tabla_p15[0])
            draw_tabla_dias_distritos_p15(canvas, tabla_p15, "Frecuencia por distrito según día", 30, POS_Y - 50, [(A4[0] - 60) / total_col] * total_col)

        elif doc.page == 16:
            if os.path.exists("assets/lineas.png"): FullImage("assets/lineas.png")(canvas, doc)

        elif doc.page == 17:
            header_footer(canvas, doc)
            if os.path.exists("assets/lins.png"):
                canvas.drawImage("assets/lins.png", (A4[0] - 700) / 2, 100, width=700, height=400, preserveAspectRatio=True, mask="auto")
            
            canvas.setFont(FONT_NAME_BOLD, 60)
            canvas.setFillColor(colors.white)
            canvas.drawCentredString(130, 185, str(total_lineas))
            
            if os.path.exists(logo_muni_path):
                canvas.drawImage(logo_muni_path, 400, 320, width=175, height=175, preserveAspectRatio=True, mask="auto")
                
            canvas.setFont(FONT_NAME_BOLD, 40)
            canvas.setFillColor(COLOR_PRIMARIO)
            canvas.drawString(330, 370, f"{lineas_municipalidad}")
            canvas.setFillColor(colors.white)
            canvas.drawString(330, 170, f"{lineas_fp}")
            
            if lineas_mixtas is not None:
                canvas.drawString(375, 280, f"Mixtas: {lineas_mixtas}")

        else:
            lineas_inicio = 18
            percepcion_inicio = lineas_inicio + (len(lineas_accion_data) * 3)
            
            # Dinámicas L.A.
            if lineas_inicio <= doc.page < percepcion_inicio:
                index = (doc.page - lineas_inicio) // 3
                posicion = (doc.page - lineas_inicio) % 3
                if index < len(lineas_accion_data):
                    linea = lineas_accion_data[index]
                    
                    if posicion == 0:
                        if os.path.exists("assets/la.png"):
                            canvas.drawImage("assets/la.png", 0, 0, width=A4[0], height=A4[1], preserveAspectRatio=True, mask="auto")
                        cor = str(linea.get("corresponsable", "")).lower().strip()
                        LOGO_Y = A4[1] - 700
                        
                        if "municipal" in cor and "mixta" not in cor:
                            if os.path.exists(logo_muni_path):
                                canvas.drawImage(logo_muni_path, A4[0] - 160 - 40, LOGO_Y, width=160, height=160, preserveAspectRatio=True, mask="auto")
                        elif "fuerza" in cor and "mixta" not in cor:
                            if os.path.exists("assets/fp.png"):
                                canvas.drawImage("assets/fp.png", A4[0] - 160 - 40, LOGO_Y, width=160, height=160, preserveAspectRatio=True, mask="auto")
                        elif "mixta" in cor:
                            if os.path.exists(logo_muni_path):
                                canvas.drawImage(logo_muni_path, A4[0] - 160 - 40, LOGO_Y, width=160, height=160, preserveAspectRatio=True, mask="auto")
                            if os.path.exists("assets/fp.png"):
                                canvas.drawImage("assets/fp.png", 40, LOGO_Y, width=160, height=160, preserveAspectRatio=True, mask="auto")
                                
                        canvas.setFont(FONT_NAME_BOLD, 150)
                        canvas.setFillColor(colors.white)
                        canvas.drawString(400, A4[1] - 300, f"{linea['numero']}")
                        
                        titulo_style = ParagraphStyle(name="TituloL", fontName=FONT_NAME_BOLD, fontSize=16, leading=20, textColor=colors.white, alignment=TA_CENTER)
                        texto_titulo = "<br/>".join(linea["problematicas"])
                        p = Paragraph(texto_titulo, titulo_style)
                        w, h = p.wrap(A4[0] * 0.7, 500)
                        p.drawOn(canvas, (A4[0] - (A4[0] * 0.7)) / 2, A4[1] - 760 - (h / 2))
                        
                    elif posicion == 1:
                        header_footer(canvas, doc)
                        draw_pagina_linea_accion(canvas, doc, linea)
                    elif posicion == 2:
                        header_footer(canvas, doc)
                        draw_pagina_linea_accion_detalle(canvas, doc, linea)

            # Percepción (6 páginas)
            elif doc.page == percepcion_inicio:
                if os.path.exists("assets/percepcion.png"): FullImage("assets/percepcion.png")(canvas, doc)
            elif doc.page == percepcion_inicio + 1:
                header_footer(canvas, doc)
                draw_pagina_percepcion_1(canvas, doc, grafico_percepcion_actual_path, grafico_percepcion_comparacion_path, tabla_percepcion)
            elif doc.page == percepcion_inicio + 2:
                header_footer(canvas, doc)
                draw_pagina_percepcion_2(canvas, doc, grafico_victimizacion_path, grafico_no_denuncia_path, tabla_no_denuncia, motivo_principal, total_omitidas)
            elif doc.page == percepcion_inicio + 3:
                header_footer(canvas, doc)
                canvas.setFont(FONT_NAME_BOLD, 16)
                canvas.setFillColor(COLOR_SECUNDARIO)
                canvas.drawString(70, 730, "Información Relevante de la Comunidad")
                
                canvas.setFont(FONT_NAME_BOLD, 12)
                canvas.drawString(70, 480 + 200 + 10, "Horario y método delictivo según percepción")
                canvas.drawImage(grafico_horarios_percepcion, 70, 480, 200, 200)
                
                tabla1 = Table(tabla_horarios_percepcion)
                tabla1.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, -1), FONT_NAME_REGULAR),
                    ("FONTSIZE", (0, 0), (-1, -1), 10),
                    ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
                    ("ALIGN", (1, 0), (1, -1), "CENTER"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]))
                tabla1.wrapOn(canvas, 0, 0)
                tabla1.drawOn(canvas, 430, 490)
                
                texto1 = f"La mayor parte de las personas que fueron víctimas de algún delito, consideran que las horas con mayor incidencia delincuencial se ubican entre las {horario_mayor} horas.<br/><br/>Total respuestas omitidas: {omitidas_aportes}."
                p1 = Paragraph(texto1, ParagraphStyle("texto", fontName=FONT_NAME_REGULAR, fontSize=11, leading=14, alignment=TA_JUSTIFY))
                p1.wrapOn(canvas, 230, 200)
                p1.drawOn(canvas, 330, 410)
                
                canvas.setFont(FONT_NAME_BOLD, 12)
                canvas.setFillColor(COLOR_SECUNDARIO)
                canvas.drawString(340, 150 + 200 + 10, "Armas utilizadas en hechos delictivos")
                canvas.drawImage(grafico_armas_percepcion, 340, 150, 200, 200)
                
                tabla2 = Table(tabla_armas)
                tabla2.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, -1), FONT_NAME_REGULAR),
                    ("FONTSIZE", (0, 0), (-1, -1), 10),
                    ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
                    ("ALIGN", (1, 0), (1, -1), "CENTER"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]))
                tabla2.wrapOn(canvas, 0, 0)
                tabla2.drawOn(canvas, 65, 180)
                
                texto2 = f"La mayor parte de las personas que fueron víctimas de algún delito consideran que el método más utilizado para la ejecución de ilícitos es el {metodo_mas_usado}.<br/><br/>Total respuestas omitidas: {omitidas_aportes}."
                p2 = Paragraph(texto2, ParagraphStyle("texto", fontName=FONT_NAME_REGULAR, fontSize=11, leading=14, alignment=TA_JUSTIFY))
                p2.wrapOn(canvas, 230, 200)
                p2.drawOn(canvas, 65, 90)

            elif doc.page == percepcion_inicio + 4:
                header_footer(canvas, doc)
                canvas.setFont(FONT_NAME_BOLD, 18)
                canvas.setFillColor(COLOR_SECUNDARIO)
                canvas.drawString(60, 740, "Percepción del Servicio Policial")
                
                canvas.drawImage(grafico_servicio_policial, 55, 515, width=360, height=220, preserveAspectRatio=True, mask='auto')
                
                # Tablas con estilo estandarizado
                def draw_mini_table(canvas, data, x, y, col_widths, num_colores):
                    tabla = Table(data, colWidths=col_widths)
                    estilo = [
                        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDE),
                        ("FONTNAME", (0, 0), (-1, -1), FONT_NAME_REGULAR),
                        ("FONTSIZE", (0, 0), (-1, -1), 9),
                        ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                        ("ALIGN", (1, 0), (1, -1), "CENTER"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ]
                    for i, color in enumerate(P3_PALETA_GRAFICO[:num_colores]):
                        estilo.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor(color)))
                    tabla.setStyle(TableStyle(estilo))
                    tabla.wrapOn(canvas, 0, 0)
                    tabla.drawOn(canvas, x, y)
                
                draw_mini_table(canvas, tabla_servicio, 420, 585, [110, 45], 5)
                
                # Gráficos e Info Inferior
                PIE_SIZE = 145
                canvas.setFont(FONT_NAME_BOLD, 10)
                canvas.setFillColor(COLOR_SECUNDARIO)
                canvas.drawString(90, 330 + PIE_SIZE + 18, "Calificación del servicio policial")
                canvas.drawString(90, 330 + PIE_SIZE + 6, "de los últimos dos años")
                canvas.drawImage(grafico_servicio_anual, 90, 330, width=PIE_SIZE, height=PIE_SIZE, preserveAspectRatio=True, mask='auto')
                draw_mini_table(canvas, [[Paragraph(str(c), ParagraphStyle("s", fontName=FONT_NAME_REGULAR, fontSize=9, textColor=colors.white))] for c in [r[0] for r in tabla_servicio_anual]], 80, 275, [90], 3)
                
                canvas.setFont(FONT_NAME_BOLD, 10)
                canvas.setFillColor(COLOR_SECUNDARIO)
                canvas.drawString(360, 330 + PIE_SIZE + 18, "¿Conoce usted a los policías?")
                canvas.drawImage(grafico_conoce_policia, 360, 330, width=PIE_SIZE, height=PIE_SIZE, preserveAspectRatio=True, mask='auto')
                draw_mini_table(canvas, [[Paragraph(str(c), ParagraphStyle("s", fontName=FONT_NAME_REGULAR, fontSize=9, textColor=colors.white))] for c in [r[0] for r in tabla_conoce]], 350, 275, [90], 3)
                
                canvas.setFont(FONT_NAME_BOLD, 11)
                canvas.drawString(40, 95 + 150 + 12, "Qué tipo de atención ha recibido")
                canvas.drawImage(grafico_atencion, 40, 95, width=320, height=150, preserveAspectRatio=True, mask='auto')
                draw_mini_table(canvas, [[Paragraph(str(c), ParagraphStyle("s", fontName=FONT_NAME_REGULAR, fontSize=9, textColor=colors.white))] for c in [r[0] for r in tabla_atencion]], 380, 115, [130], len(tabla_atencion))
                
                canvas.setFont(FONT_NAME_BOLD, 11)
                canvas.setFillColor(colors.black)
                canvas.drawString(470, 60 + 18, "Respuestas omitidas")
                canvas.setFont(FONT_NAME_REGULAR, 11)
                canvas.drawString(470, 60, str(omitidas_servicio))

            elif doc.page == percepcion_inicio + 5:
                header_footer(canvas, doc)
                canvas.setFont(FONT_NAME_BOLD, 18)
                canvas.setFillColor(COLOR_SECUNDARIO)
                canvas.drawString(60, 740, "Percepción Sector Comercial")
                
                if os.path.exists("assets/comer.png"):
                    canvas.drawImage("assets/comer.png", 258, 381, 80, 80)
                
                canvas.setFont(FONT_NAME_BOLD, 11)
                canvas.drawString(70, 470 + 200 + 15 + 12, "¿Se siente seguro en")
                canvas.drawString(70, 470 + 200 + 15, "su establecimiento comercial?")
                canvas.drawImage(grafico_comercio_seguridad, 70, 470, 200, 200)
                
                canvas.drawString(330, 470 + 200 + 15 + 12, "¿Conoce el programa de Seguridad Comercial")
                canvas.drawString(330, 470 + 200 + 15, "que imparte Fuerza Pública?")
                canvas.drawImage(grafico_comercio_programa, 330, 470, 200, 200)
                
                canvas.drawString(70, 130 + 200 + 15 + 12, "¿Está inscrito en el programa")
                canvas.drawString(70, 130 + 200 + 15, "de Seguridad Comercial?")
                canvas.drawImage(grafico_comercio_inscrito, 70, 130, 200, 200)
                
                canvas.drawString(330, 130 + 200 + 15 + 12, "¿Le gustaría que se le contacte")
                canvas.drawString(330, 130 + 200 + 15, "para formar parte del programa?")
                canvas.drawImage(grafico_comercio_contacto, 330, 130, 200, 200)
                
            elif doc.page == percepcion_inicio + 6:
                if os.path.exists("assets/final.png"): FullImage("assets/final.png")(canvas, doc)

    doc.build(
        story,
        onFirstPage=first_page,
        onLaterPages=later_pages
    )
    buffer.seek(0)
    return buffer
