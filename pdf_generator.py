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
    TableStyle
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from io import BytesIO
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT

# ===== ESTILO GLOBAL PARA TABLAS PARETO =====
pareto_cell_style = ParagraphStyle(
    name="ParetoCell",
    fontName="Helvetica",
    fontSize=11,
    leading=13,
    alignment=TA_LEFT,
    wordWrap="CJK"
)

# ================= UTILIDAD FULL PAGE =================
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


# ================= HEADER / FOOTER =================
def header_footer(canvas, doc):
    page_width, page_height = A4

    canvas.drawImage(
        "assets/header.png",
        0,
        page_height - 80,
        width=page_width,
        height=60,
        mask="auto"
    )

    canvas.drawImage(
        "assets/footer.png",
        0,
        0,
        width=page_width,
        height=60,
        mask="auto"
    )


# ================= GRAFICO DE EDAD =================
def draw_grafico_edad(canvas, doc, grafico_edad_path):
    page_width, page_height = A4

    img_width = 220
    img = ImageReader(grafico_edad_path)
    img_w, img_h = img.getSize()
    img_height = img_width * img_h / img_w

    x = 40
    y = page_height - img_height - 120

    canvas.drawImage(
        grafico_edad_path,
        x,
        y,
        width=img_width,
        height=img_height,
        preserveAspectRatio=True,
        mask="auto"
    )


# ================= GRAFICO ESCOLARIDAD =================
def draw_grafico_escolaridad(canvas, grafico_path):
    page_width, page_height = A4

    img_width = 220
    img = ImageReader(grafico_path)
    img_w, img_h = img.getSize()
    img_height = img_width * img_h / img_w

    x = page_width / 2 + 10
    y = page_height - img_height - 340

    canvas.drawImage(
        grafico_path,
        x,
        y,
        width=img_width,
        height=img_height,
        preserveAspectRatio=True,
        mask="auto"
    )


# ================= GRAFICO GENERO =================
def draw_grafico_genero(canvas, grafico_path):
    page_width, page_height = A4

    img_width = 220
    img = ImageReader(grafico_path)
    img_w, img_h = img.getSize()
    img_height = img_width * img_h / img_w

    x = 40
    y = page_height - img_height - 540

    canvas.drawImage(
        grafico_path,
        x,
        y,
        width=img_width,
        height=img_height,
        preserveAspectRatio=True,
        mask="auto"
    )


# ================= TABLA EDAD =================
def draw_tabla_edad(canvas, doc, tabla_edad):
    page_width, page_height = A4

    TABLE_WIDTH = 220
    FONT_SIZE_HEADER = 12
    FONT_SIZE_BODY = 11
    X = page_width / 2 + 10
    Y = page_height - 60

    data = [["Participación por Edad", ""]]
    data.extend(tabla_edad)

    table = Table(
        data,
        colWidths=[TABLE_WIDTH * 0.6, TABLE_WIDTH * 0.4]
    )

    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (0, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DEEBF7")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(0x000000)),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), FONT_SIZE_HEADER),
        ("GRID", (0, 1), (-1, -1), 0.5, colors.white),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 1), (-1, -1), FONT_SIZE_BODY),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFFFFF")),
    ]))

    colores_filas = [
        colors.HexColor("#5B9BD5"),
        colors.HexColor("#A5A5A5"),
        colors.HexColor("#4472C4"),
        colors.HexColor("#255E91"),
        colors.HexColor("#B7B7B7")
    ]

    for i, color in enumerate(colores_filas, start=1):
        table.setStyle([
            ("BACKGROUND", (0, i), (-1, i), color),
            ("TEXTCOLOR", (0, i), (-1, i), colors.white)
        ])

    table.wrapOn(canvas, TABLE_WIDTH, 200)
    table.drawOn(canvas, X, Y - 200)


# ================= TABLA ESCOLARIDAD =================
def draw_tabla_escolaridad(canvas, tabla_escolaridad):
    page_width, page_height = A4

    TABLE_WIDTH = 220
    FONT_SIZE_HEADER = 12
    FONT_SIZE_BODY = 11

    x = 20
    y = page_height - 340

    data = [["Participación por Escolaridad", ""]]
    data.extend(tabla_escolaridad)

    table = Table(
        data,
        colWidths=[TABLE_WIDTH * 0.6, TABLE_WIDTH * 0.4]
    )

    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DEEBF7")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), FONT_SIZE_HEADER),
        ("GRID", (0, 1), (-1, -1), 0.5, colors.white),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 1), (-1, -1), FONT_SIZE_BODY),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFFFFF")),
    ]))

    colores_filas = [
        colors.HexColor("#5B9BD5"),
        colors.HexColor("#A5A5A5"),
        colors.HexColor("#4472C4"),
        colors.HexColor("#255E91"),
        colors.HexColor("#B7B7B7"),
        colors.HexColor("#9DC3E6"),
        colors.HexColor("#8FAADC"),
        colors.HexColor("#D9E1F2")
    ]

    for i, color in enumerate(colores_filas, start=1):
        table.setStyle([
            ("BACKGROUND", (0, i), (-1, i), color),
            ("TEXTCOLOR", (0, i), (-1, i), colors.white)
        ])

    table.wrapOn(canvas, TABLE_WIDTH, 200)
    table.drawOn(canvas, x, y - 200)


# ================= TABLA GENERO =================
def draw_tabla_genero(canvas, tabla_genero):
    page_width, page_height = A4

    TABLE_WIDTH = 220
    FONT_SIZE_HEADER = 12
    FONT_SIZE_BODY = 11

    x = page_width / 2 + 10
    y = page_height - 510

    data = [["Participación por Género", ""]]
    data.extend(tabla_genero)

    table = Table(
        data,
        colWidths=[TABLE_WIDTH * 0.6, TABLE_WIDTH * 0.4]
    )

    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DEEBF7")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), FONT_SIZE_HEADER),
        ("GRID", (0, 1), (-1, -1), 0.5, colors.white),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 1), (-1, -1), FONT_SIZE_BODY),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFFFFF")),
    ]))

    colores_filas = [
        colors.HexColor("#5B9BD5"),
        colors.HexColor("#A5A5A5"),
        colors.HexColor("#4472C4")
    ]

    for i, color in enumerate(colores_filas, start=1):
        table.setStyle([
            ("BACKGROUND", (0, i), (-1, i), color),
            ("TEXTCOLOR", (0, i), (-1, i), colors.white)
        ])

    table.wrapOn(canvas, TABLE_WIDTH, 200)
    table.drawOn(canvas, x, y - 200)


#-------------*-*-*-*-*-*-*-*-*-*-Tablas de encuestas
def draw_tabla_simple(
    canvas,
    data,
    titulo,
    x,
    y,
    col_widths,
    header_color=colors.HexColor("#013051"),
    body_color=colors.white,
    border_color=colors.black,
    font_size_header=12,
    font_size_body=11
):

    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle

    TABLE_WIDTH = sum(col_widths)

    # ===== ESTILO CELDAS =====
    cell_style = ParagraphStyle(
        name="CellStyle",
        fontName="Helvetica",
        fontSize=font_size_body,
        leading=font_size_body + 2,
        alignment=TA_LEFT,
        wordWrap="CJK"
    )

    header_style = ParagraphStyle(
        name="HeaderStyle",
        fontName="Helvetica-Bold",
        fontSize=font_size_header,
        alignment=TA_LEFT
    )

    # ===== CONSTRUIR TABLA =====
    table_data = [
        [Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1)
    ]

    for row in data:
        nueva_fila = []
        for cell in row:
            nueva_fila.append(Paragraph(str(cell), cell_style))
        table_data.append(nueva_fila)

    table = Table(table_data, colWidths=col_widths)

    style = [
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 1, border_color),
        ("BACKGROUND", (0, 1), (-1, -1), body_color),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]

    table.setStyle(TableStyle(style))
    table.wrapOn(canvas, TABLE_WIDTH, 400)
    table.drawOn(canvas, x, y)

#=========================================Tabla Pareto==================

def draw_tabla_pareto(
    canvas,
    titulo,
    data,
    x,
    y,
    table_width=180,
    font_title="Helvetica-Bold",
    font_size_title=14,
    header_color=colors.HexColor("#4471C4"),
    body_color=colors.white,
    text_color=colors.black
):
    # -------- FONDO DEL TÍTULO --------
    canvas.setFillColor(header_color)
    canvas.rect(
        x,
        y + 8,
        table_width,
        24,
        stroke=0,
        fill=1
    )
    
    # -------- TEXTO DEL TÍTULO --------
    canvas.setFont(font_title, font_size_title)
    canvas.setFillColor(text_color)
    canvas.drawCentredString(
        x + table_width / 2,
        y + 14,
        titulo
    )
    

    # -------- PREPARAR DATOS CON PARAGRAPH --------
    wrapped_data = [
        [Paragraph(str(row[0]), pareto_cell_style)]
        for row in data
    ]

    table = Table(
        wrapped_data,
        colWidths=[table_width],
        rowHeights=None  # 👈 auto height
    )

    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, -1), body_color),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    table.wrapOn(canvas, table_width, 400)
    table.drawOn(canvas, x, y - table._height)


def draw_porcentaje(
    canvas,
    texto,
    x,
    y,
    font="Helvetica-Bold",
    size=18,
    color=colors.black
):
    canvas.setFont(font, size)
    canvas.setFillColor(color)
    canvas.drawCentredString(x, y, texto)

def draw_cantidad(
    canvas,
    texto,
    x,
    y,
    font="Helvetica-Bold",
    size=14,
    color=colors.black
):
    canvas.setFont(font, size)
    canvas.setFillColor(color)
    canvas.drawCentredString(x, y, texto)

##=============MIC-MAC_________________________

def draw_micmac_lista(
    canvas,
    data,
    x,
    y,
    width=200,
    max_height=180,
    font_size=8
):
    style = ParagraphStyle(
        name="MicmacCell",
        fontName="Helvetica",
        fontSize=font_size,
        leading=9,
        alignment=TA_LEFT,
        wordWrap="CJK"
    )

    # Convertir a Paragraph y dividir en 2 columnas
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

    table = Table(
        table_data,
        colWidths=[width / 2, width / 2],
        rowHeights=None
    )

    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))

    table.wrapOn(canvas, width, max_height)
    table.drawOn(canvas, x, y - table._height)

#==================MICMAC2=========================

def draw_tabla_overlay(
    canvas,
    data,
    x,
    y,
    width=120,
    font_size=8,
    max_height=160
):
    style = ParagraphStyle(
        name="OverlayCell",
        fontName="Helvetica",
        fontSize=font_size,
        leading=font_size + 2,
        alignment=TA_LEFT,
        wordWrap="CJK"
    )

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

#_____texto______________________
    
def draw_texto_overlay(
    canvas,
    texto,
    x,
    y,
    font="Helvetica-Bold",
    size=20,
    color=colors.white
):
    canvas.setFont(font, size)
    canvas.setFillColor(color)
    canvas.drawCentredString(x, y, str(texto))

#_______________________Triangulo de violes-----------------------

def draw_texto_mixto(
    canvas,
    x,
    y,
    texto_antes,
    valor_1,
    texto_medio,
    valor_2,
    texto_despues,
    width=250,
    font_size=11,
    valor_size=15
):
    normal = ParagraphStyle(
        "normal",
        fontName="Helvetica",
        fontSize=font_size,
        alignment=TA_JUSTIFY
    )

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

#_________________________________TABLA DE INSTIS_____________________________________


def draw_tabla_instituciones(
    canvas,
    data,
    titulo="Lista de Instituciones",
    x=60,
    y=250,
    table_width=480,
    header_bg=colors.HexColor("#013051"),
    header_text_color=colors.white,
    body_bg=colors.whitesmoke,
    border_color=colors.black,
    font_header=12,
    font_body=10
):
    if not data:
        return

    # Construir encabezado
    table_data = [[titulo, ""]]
    table_data.extend(data)

    col_widths = [table_width * 0.6, table_width * 0.4]

    table = Table(table_data, colWidths=col_widths)

    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), header_text_color),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), font_header),

        ("GRID", (0, 1), (-1, -1), 0.5, border_color),
        ("BACKGROUND", (0, 1), (-1, -1), body_bg),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), font_body),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    table.wrapOn(canvas, table_width, 300)
    table.drawOn(canvas, x, y - table._height)

#====================TABLA DE HORARIOS PAGINA 13===================

def draw_tabla_horario_distrito(
    canvas,
    data,
    titulo,
    x,
    y,
    col_widths
):

    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER

    TABLE_WIDTH = sum(col_widths)

    header_style = ParagraphStyle(
        name="HorarioHeader",
        fontName="Helvetica-Bold",
        fontSize=9,
        alignment=TA_CENTER,
        leading=11
    )

    cell_style = ParagraphStyle(
        name="HorarioCell",
        fontName="Helvetica",
        fontSize=8,
        alignment=TA_CENTER,
        leading=9,
        wordWrap="CJK"
    )

    table_data = []

    # Título
    table_data.append(
        [Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1)
    )

    # Encabezados + datos
    for row in data:
        nueva_fila = []
        for cell in row:
            nueva_fila.append(Paragraph(str(cell), cell_style))
        table_data.append(nueva_fila)

    table = Table(table_data, colWidths=col_widths)

    style = [
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#30a907")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),

        # Encabezado real
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E2F0D9")),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),

        # Bordes
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),

        # Centrado
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        # Compactar altura
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LEFTPADDING", (0, 0), (-1, -1), 1),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
    ]

    table.setStyle(TableStyle(style))

    table.wrapOn(canvas, TABLE_WIDTH, 2000)
    table.drawOn(canvas, x, y - table._height)

#=======================TABLA PAGINA 14============
    
def draw_tabla_modalidades_p14(
    canvas,
    data,
    titulo,
    x,
    y,
    col_widths
):

    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    TABLE_WIDTH = sum(col_widths)

    # ===== VARIABLES EDITABLES =====
    COLOR_TITULO = colors.HexColor("#30a907")
    COLOR_HEADER = colors.HexColor("#E2Fdd9")
    COLOR_BODY = colors.white
    COLOR_BORDER = colors.black

    FONT_HEADER = 8
    FONT_BODY = 8

    # ===============================

    header_style = ParagraphStyle(
        name="HeaderStyleP14",
        fontName="Helvetica-Bold",
        fontSize=FONT_HEADER,
        alignment=TA_CENTER,
        leading=10
    )

    distrito_style = ParagraphStyle(
        name="DistritoStyleP14",
        fontName="Helvetica-Bold",
        fontSize=FONT_BODY,
        alignment=TA_LEFT,
        leading=9
    )

    cell_style = ParagraphStyle(
        name="CellStyleP14",
        fontName="Helvetica",
        fontSize=FONT_BODY,
        alignment=TA_CENTER,
        leading=9,
        wordWrap="CJK"
    )

    table_data = []

    # Título
    table_data.append(
        [Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1)
    )

    # Encabezados
    header_row = []
    for cell in data[0]:
        header_row.append(Paragraph(str(cell), header_style))
    table_data.append(header_row)

    # Datos
    for row in data[1:]:
        nueva_fila = []
        for i, cell in enumerate(row):
            if i == 0:
                nueva_fila.append(Paragraph(str(cell), distrito_style))
            else:
                nueva_fila.append(Paragraph(str(cell), cell_style))
        table_data.append(nueva_fila)

    table = Table(table_data, colWidths=col_widths)

    style = [
        # Título
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_TITULO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),

        # Encabezado modalidades
        ("BACKGROUND", (0, 1), (-1, 1), COLOR_HEADER),
        ("TEXTCOLOR", (0, 1), (-1, 1), colors.white),

        # Columna Distritos
        ("BACKGROUND", (0, 2), (0, -1), COLOR_HEADER),
        ("TEXTCOLOR", (0, 2), (0, -1), colors.white),

        # Cuerpo
        ("BACKGROUND", (1, 2), (-1, -1), COLOR_BODY),

        # Bordes
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDER),

        # Centrado general
        ("ALIGN", (1, 2), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        # Compacto
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]

    table.setStyle(TableStyle(style))

    table.wrapOn(canvas, TABLE_WIDTH, 2000)
    table.drawOn(canvas, x, y - table._height)

#======================TABLA P15...................

def draw_tabla_dias_distritos_p15(
    canvas,
    data,
    titulo,
    x,
    y,
    col_widths
):

    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    TABLE_WIDTH = sum(col_widths)

    # ===== VARIABLES EDITABLES =====
    COLOR_TITULO = colors.HexColor("#30a907")
    COLOR_HEADER = colors.HexColor("#E2Fdd9")
    COLOR_BODY = colors.white
    COLOR_BORDER = colors.black

    FONT_HEADER = 8
    FONT_BODY = 7
    # =================================

    header_style = ParagraphStyle(
        name="HeaderStyleP15",
        fontName="Helvetica-Bold",
        fontSize=FONT_HEADER,
        alignment=TA_CENTER,
        leading=10
    )

    distrito_style = ParagraphStyle(
        name="DistritoStyleP15",
        fontName="Helvetica-Bold",
        fontSize=FONT_BODY,
        alignment=TA_LEFT,
        leading=9
    )

    cell_style = ParagraphStyle(
        name="CellStyleP15",
        fontName="Helvetica",
        fontSize=FONT_BODY,
        alignment=TA_CENTER,
        leading=9,
        wordWrap="CJK"
    )

    table_data = []

    # Título
    table_data.append(
        [Paragraph(titulo, header_style)] + [""] * (len(data[0]) - 1)
    )

    # Encabezado (dias)
    header_row = []
    for cell in data[0]:
        header_row.append(Paragraph(str(cell), header_style))
    table_data.append(header_row)

    # Datos
    for row in data[1:]:

        nueva_fila = []

        for i, cell in enumerate(row):

            if i == 0:
                nueva_fila.append(Paragraph(str(cell), distrito_style))
            else:
                nueva_fila.append(Paragraph(str(cell), cell_style))

        table_data.append(nueva_fila)

    table = Table(table_data, colWidths=col_widths)

    style = [

        # Título
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_TITULO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),

        # Encabezado dias
        ("BACKGROUND", (0, 1), (-1, 1), COLOR_HEADER),
        ("TEXTCOLOR", (0, 1), (-1, 1), colors.white),

        # Columna distrito
        ("BACKGROUND", (0, 2), (0, -1), COLOR_HEADER),
        ("TEXTCOLOR", (0, 2), (0, -1), colors.white),

        # Cuerpo
        ("BACKGROUND", (1, 2), (-1, -1), COLOR_BODY),

        # Bordes
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDER),

        # Centrado
        ("ALIGN", (1, 2), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        # Compacto
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]

    table.setStyle(TableStyle(style))

    table.wrapOn(canvas, TABLE_WIDTH, 2000)
    table.drawOn(canvas, x, y - table._height)

#==============================DEFINICION; LINEAS ACCION==========

def normalizar_nombre(texto):
    texto = texto.lower()

    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

    texto = texto.replace(" ", "_")

    # 🔥 IMPORTANTE: NO quitamos paréntesis ahora
    texto = re.sub(r"[^a-z0-9_()]", "", texto)

    return texto

def draw_pagina_linea_accion(
    canvas,
    doc,
    linea,
    vectores_path="assets/vectores"
):
    page_width, page_height = A4

    # ================= CONFIGURABLES =================
    SUBTITULO_FONT = "Helvetica-Bold"
    SUBTITULO_SIZE = 14
    SUBTITULO_COLOR = colors.HexColor("#013051")
    SUBTITULO_X = 40
    SUBTITULO_Y = page_height - 110

    TITULO_FONT = "Helvetica-Bold"
    TITULO_SIZE = 18
    TITULO_COLOR = colors.HexColor("#013051")
    TITULO_WIDTH = page_width * 0.8
    TITULO_X = (page_width - TITULO_WIDTH) / 2
    TITULO_Y = page_height - 130

    VECTOR_WIDTH = 150
    VECTOR_HEIGHT = 150
    VECTOR_Y = page_height - 300
    VECTOR_SPACING = 30

    TABLA_C_X = 40
    TABLA_C_Y = page_height - 310
    TABLA_C_WIDTH = page_width * 0.45

    TABLA_P_X = page_width / 2 + 10
    TABLA_P_Y = page_height - 310
    TABLA_P_WIDTH = page_width * 0.45
    # =================================================

    # ===== SUBTITULO =====
    canvas.setFont(SUBTITULO_FONT, SUBTITULO_SIZE)
    canvas.setFillColor(SUBTITULO_COLOR)
    canvas.drawString(SUBTITULO_X, SUBTITULO_Y, "Líneas de Acción")

    # ===== TITULO =====
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER

    titulo_style = ParagraphStyle(
        name="TituloLineaInterna",
        fontName=TITULO_FONT,
        fontSize=TITULO_SIZE,
        textColor=TITULO_COLOR,
        alignment=TA_CENTER
    )

    texto_titulo = "<br/>".join(linea["problematicas"])
    p = Paragraph(texto_titulo, titulo_style)
    w, h = p.wrap(TITULO_WIDTH, 200)
    p.drawOn(canvas, TITULO_X, TITULO_Y)

    # ===== VECTORES =====
    total_vectores = len(linea["problematicas"])
    if total_vectores > 0:
        total_width = total_vectores * VECTOR_WIDTH + (total_vectores - 1) * VECTOR_SPACING
        start_x = (page_width - total_width) / 2

        for i, problema in enumerate(linea["problematicas"]):
            nombre = normalizar_nombre(problema) + ".png"
            ruta = os.path.join(vectores_path, nombre)

            if os.path.exists(ruta):
                canvas.drawImage(
                    ruta,
                    start_x + i * (VECTOR_WIDTH + VECTOR_SPACING),
                    VECTOR_Y,
                    width=VECTOR_WIDTH,
                    height=VECTOR_HEIGHT,
                    preserveAspectRatio=True,
                    mask="auto"
                )

    #FORMATO DE CELDAS
    CELDA_STYLE = ParagraphStyle(
        name="CeldaTabla",
        fontName="Helvetica",
        fontSize=9,
        leading=11,   # controla interlineado
        alignment=TA_LEFT
    )
    
    # ===== TABLA CAUSAS =====
    from reportlab.platypus import Table, TableStyle

    tabla_c_data = [["Causas Socio Culturales y Estructurales"]]
    
    for fila in linea["causas"]:
        texto = Paragraph(fila[0], CELDA_STYLE)
        tabla_c_data.append([texto])
    
    tabla_c = Table(
        tabla_c_data,
        colWidths=[TABLA_C_WIDTH]
    )

    tabla_c.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#30A907")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))

    tabla_c.wrapOn(canvas, TABLA_C_WIDTH, 400)
    tabla_c.drawOn(canvas, TABLA_C_X, TABLA_C_Y - tabla_c._height)

    # ===== TABLA PROBLEMAS =====
    tabla_p_data = [["Problematicas Influyentes"]]

    for fila in linea["problemas_influyentes"]:
        texto = Paragraph(fila[0], CELDA_STYLE)
        tabla_p_data.append([texto])
    
    tabla_p = Table(
        tabla_p_data,
        colWidths=[TABLA_P_WIDTH]
    )

    tabla_p.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#013051")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))

    tabla_p.wrapOn(canvas, TABLA_P_WIDTH, 400)
    tabla_p.drawOn(canvas, TABLA_P_X, TABLA_P_Y - tabla_p._height)

# ================= GENERADOR PDF =================
def generar_pdf(
    portada_path,
    grafico_relacion_path,
    grafico_edad_path,
    grafico_escolaridad_path,
    grafico_genero_path,
    delegacion,
    codigo,
    tabla_participacion,
    tabla_edad,
    tabla_escolaridad,
    tabla_genero,
    tabla_encuesta_comunidad, 
    tabla_otras_encuestas,
    datos_pagina_8,
    datos_pagina_9,
    tabla_delitos,
    tabla_riesgos,
    porcentaje_delitos,
    porcentaje_riesgos,
    cantidad_delitos,
    cantidad_riesgos,
    micmac_poder,
    micmac_conflicto,
    micmac_autonomas,
    micmac_resultados,
    tabla_riesgos_micmac2,
    tabla_delitos_micmac2,
    cantidad_problematicas,
    riesgos_total,
    delitos_total,
    causas_identificadas,
    factores_micmac,
    triangulo_directa,
    triangulo_sociocultural,
    triangulo_estructural,
    tabla_instituciones,
    grafico_denuncias_path,
    tabla_denuncias,
    total_denuncias,
    grafico_horario_path,
    tabla_horario,
    total_am,
    total_pm,
    tabla_horario_distrito,
    grafico_p14_path,
    tabla_p14,
    grafico_p15_path,
    tabla_p15,
    total_lineas,
    lineas_municipalidad,
    lineas_fp,
    lineas_mixtas,
    logo_muni_path,
    lineas_accion_data,    
):

    buffer = BytesIO()
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="NormalJustificado",
        parent=styles["Normal"],
        alignment=TA_JUSTIFY
    ))

    styles.add(ParagraphStyle(
        name="TituloGrande",
        fontName="Helvetica",
        fontSize=26,
        textColor=colors.white,
        leading=80,
        spaceAfter=10,
        alignment=TA_LEFT
    ))

    styles.add(ParagraphStyle(
        name="TituloDelta",
        fontName="Helvetica-Bold",
        fontSize=45,
        textColor=colors.HexColor("#30a907"),
        leading=30,
        spaceAfter=10,
        alignment=TA_LEFT
    ))

    styles.add(ParagraphStyle(
        name="TituloD2",
        fontName="Helvetica-Bold",
        fontSize=60,
        textColor=colors.white,
        leading=30,
        spaceAfter=10,
        alignment=TA_LEFT
    ))

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    story = []

    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("DELEGACIÓN POLICIAL", styles["TituloGrande"]))
    story.append(Paragraph(delegacion, styles["TituloDelta"]))
    story.append(Paragraph(codigo, styles["TituloD2"]))

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
        "El presente informe, elaborado para el territorio que comprende la Delegación Policial de San Ramón, surge como una herramienta esencial para la toma efectiva de decisiones. Este informe se concibe como un instrumento dinámico y orientado hacia el futuro, diseñado para proporcionar información clave y un plan de trabajo estructurado que permita abordar las problemáticas prioritarias identificadas en el ámbito de la seguridad pública.",
        styles["NormalJustificado"]
    ))

    story.append(Spacer(1, 20))
    story.append(Image("assets/conformacion.png", width=600, height=400))

    story.append(PageBreak())
    story.append(PageBreak())

    story.append(Spacer(1, 40))
    story.append(Paragraph("Datos de Participación", styles["Heading1"]))
    story.append(Spacer(1, 20))
    story.append(Paragraph("Participación por Distrito", styles["Heading2"]))
    story.append(Spacer(1, 10))

    tabla = Table(tabla_participacion, colWidths=[180, 180, 120])
    tabla.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#013051")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("BACKGROUND", (1, 1), (-1, -1), colors.HexColor("#E2FDD9"))
    ]))

    story.append(tabla)

    story.append(Spacer(1, 25))
    story.append(Paragraph("Relación por distrito", styles["Heading2"]))
    story.append(Spacer(1, 15))
    story.append(Image(grafico_relacion_path, width=400, height=250))

    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Datos de Participación", styles["Heading1"]))
    story.append(Spacer(1, 200))
    story.append(Paragraph("__________________________________________________________________________________________"))
    story.append(Spacer(1, 210))
    story.append(Paragraph("__________________________________________________________________________________________"))

    story.append(PageBreak())
    story.append(PageBreak())

    story.append(Spacer(1, 40))
    story.append(Paragraph("Proceso Metodológico", styles["Heading1"]))
    story.append(Paragraph(
    "Información demográfica según zona asignada a la Delegación Policial",
    styles["Normal"]
    ))

    story.append(Spacer(1, 190))    
    story.append(Image("assets/netquest.png", width=550, height=125))

    story.append(Spacer(1, 1))
    story.append(Paragraph("__________________________________________________________________________________________"))
    
    story.append(PageBreak())

    story.append(Spacer(1, 40))
    story.append(Paragraph("Diagrama de Pareto", styles["Heading1"]))
    story.append(Paragraph(
    "(Aplicando el principio de 80/20 donde el 80% es lo trivial y el 20% es lo vital)",
    styles["Normal"]
    ))    
    story.append(Spacer(1, 170))
    story.append(Paragraph("__________________________________________________________________________________________"))

    story.append(PageBreak())

    story.append(Spacer(1, 40))
    story.append(Paragraph("MICMAC", styles["Heading1"]))
    story.append(Paragraph(
    "((Matriz de Impactos Cruzado – Multiplicación Aplicada a un Clasificación))",
    styles["Normal"]
    ))   
    story.append(Spacer(1, 390))
    story.append(Paragraph("__________________________________________________________________________________________"))

    story.append(PageBreak())

    story.append(Spacer(1, 40))
    story.append(Paragraph("Triangulo de las Violencias", styles["Heading1"]))
    story.append(Spacer(1, 10))  
    story.append(Paragraph(("Es una metodología también llamada “teoría de los conflictos” creada por el sociólogo y matemático Johan Galtung, que permite estudiar y transformar los conflictos mediante la identificación de variantes de la violencia, (Directa, cultural y estructural), visualizando las causas generadoras de las problemáticas. "),
    styles["NormalJustificado"]
    ))  
    story.append(Spacer(1, 200))
    story.append(Paragraph("__________________________________________________________________________________________"))
    story.append(Spacer(1, 20))
    story.append(Paragraph("Proceso Metodológico", styles["Heading1"]))
    story.append(Paragraph("Lista de Instituciones participantes en calificacion de procesos: MIC-MAC y Triangulo de las Violencias", styles["Heading2"]))

    story.append(PageBreak())
    story.append(PageBreak())

    story.append(Spacer(1, 30))
    story.append(Paragraph("Denuncias por distrito", styles["Heading2"]))
    story.append(Spacer(1, 220))
    story.append(Paragraph("__________________________________________________________________________________________"))
    story.append(Paragraph("Denuncias por rango horario", styles["Heading2"]))

    story.append(PageBreak())   
    story.append(Spacer(1, 30))
    story.append(Paragraph("Denuncias por Modalidad en el Cantón", styles["Heading2"]))
    story.append(Spacer(1, 280))
    story.append(Paragraph("__________________________________________________________________________________________"))

    story.append(PageBreak())   
    story.append(Spacer(1, 30))
    story.append(Paragraph("Denuncias por dias de la semana en el Cantón", styles["Heading2"]))
    story.append(Spacer(1, 280))
    story.append(Paragraph("__________________________________________________________________________________________"))

    story.append(PageBreak())
    story.append(PageBreak())

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

   # =========================================
    # CREAR PAGINAS DINAMICAS DE PORTADAS
    # =========================================
    
    for i in range(len(lineas_accion_data) * 2):
        story.append(PageBreak())
        story.append(Spacer(1, 1))
    ##----------------------------------Bloque de funciones 
                 
    def first_page(canvas, doc):
        FullImage(portada_path)(canvas, doc)

    def later_pages(canvas, doc):
        if doc.page == 2:
            FullImage("assets/intro.png")(canvas, doc)
        elif doc.page == 4:
            FullImage("assets/participacion.png")(canvas, doc)
        elif doc.page == 6:
            header_footer(canvas, doc)
            draw_grafico_edad(canvas, doc, grafico_edad_path)
            draw_tabla_edad(canvas, doc, tabla_edad)
            draw_grafico_escolaridad(canvas, grafico_escolaridad_path)
            draw_tabla_escolaridad(canvas, tabla_escolaridad)
            draw_grafico_genero(canvas, grafico_genero_path)
            draw_tabla_genero(canvas, tabla_genero)
        elif doc.page == 7:
            FullImage("assets/metodologico.png")(canvas, doc)
        elif doc.page == 8:
            header_footer(canvas, doc)
            
            page_width, page_height = A4

            # ================= TABLA 1 =================
            draw_tabla_simple(
                canvas=canvas,
                data=tabla_encuesta_comunidad,
                titulo="Encuesta a Comunidad",
                x=90,                  # 👈 posición horizontal (editable)
                y=page_height - 190,   # 👈 posición vertical (editable)
                col_widths=[100, 100, 100, 100],  # 👈 ancho columnas
                header_color=colors.HexColor("#4471C4")
            )

            # ================= TABLA 2 =================
            draw_tabla_simple(
                canvas=canvas,
                data=tabla_otras_encuestas,
                titulo="Otras encuestas",
                x=90,                  # 👈 centrada respecto a la tabla 1
                y=page_height - 300,   # 👈 distancia estética hacia abajo
                col_widths=[100, 100, 100, 100],
                header_color=colors.HexColor("#4471C4")
            )

            # ================= IMAGEN DATOS =================
            img_x = 0            # posición horizontal (izquierda → derecha)
            img_y = 20           # posición vertical (abajo → arriba)
            img_width = 600       # ancho de la imagen
            img_height = 400     # alto de la imagen

            canvas.drawImage(
                "assets/datos.png",
                img_x,
                img_y,
                width=img_width,
                height=img_height,
                preserveAspectRatio=True,
                mask="auto"
            )

            # ================= DATOS SOBRE LA IMAGEN =================
            canvas.setFont("Helvetica-Bold", 22)
            canvas.setFillColor(colors.white)

            # ENCUESTA COMUNIDAD
            canvas.drawString(110, 265, str(datos_pagina_8["encuesta_comunidad"]))
            
            # ENCUESTA comercio
            canvas.drawString(80,140, str(datos_pagina_8["encuesta_policial"]))
            
            # ENCUESTA policial
            canvas.drawString(400, 265, str(datos_pagina_8["encuesta_comercio"]))
            
            # ESTADÍSTICA REGISTRADA
            canvas.drawString(430, 140, str(datos_pagina_8["estadistica"]))
            
            # TOTAL DE DATOS (más grande)
            canvas.setFont("Helvetica-Bold", 28)
            canvas.setFillColor(colors.HexColor("#FFFFFF"))
            canvas.drawString(260, 150, str(datos_pagina_8["total_datos"]))
            
        elif doc.page == 9:
            header_footer(canvas, doc)
            img_width = 550     # ancho de la imagen
            img_height = 300    # alto de la imagen

            img_x = (A4[0] - img_width) / 2   # centrado horizontal
            img_y = A4[1] - img_height - 70  # debajo del header

            canvas.drawImage(
                "assets/pareto.png",
                img_x,
                img_y,
                width=img_width,
                height=img_height,
                preserveAspectRatio=True,
                mask="auto"
            )

            #___datos pareto______
            canvas.setFont("Helvetica-Bold", 22)
            canvas.setFillColor(colors.white)

            canvas.drawString(
                img_x + 150,              # x (izquierda)
                img_y + img_height - 155, # y (parte superior)
                datos_pagina_9["lado_izquierdo"]
            )

            canvas.drawString(
                img_x + img_width - 105, # x (derecha)
                img_y + img_height - 115, # y (superior)arriba abajo
                datos_pagina_9["derecha_superior"]
            )

            canvas.drawString(
                img_x + img_width - 105, # x (derecha)
                img_y + 105,              # y (inferior)
                datos_pagina_9["derecha_inferior"]
            )

            #===========listas paretos
            page_width, page_height = A4

            x_izq = 60
            x_der = page_width / 2 + 20
            y_base = page_height - 380

                       #=======DELITOS===
            draw_porcentaje(
                canvas,
                porcentaje_delitos,
                x_izq + 90,
                y_base + 45
            )
            
            color_titulo_riesgos = colors.HexColor("#FFFFFF")
            fondo_titulo_riesgos = colors.HexColor("#30a907")
            fondo_filas_riesgos  = colors.HexColor("#F2F2F2")
            
            draw_tabla_pareto(
                canvas,
                "Delitos",
                tabla_delitos,
                x_izq,
                y_base,
                text_color=color_titulo_riesgos,
                header_color=fondo_titulo_riesgos,
                body_color=fondo_filas_riesgos
            )
            
            draw_cantidad(
                canvas,
                f"Total: {cantidad_delitos}",
                x_izq + 90,
                y_base - 365
            )
            
            #=======Riesgos============
            draw_porcentaje(
                canvas,
                porcentaje_riesgos,
                x_der + 90,
                y_base + 45
            )
            
            color_titulo_riesgos = colors.HexColor("#FFFFFF")
            fondo_titulo_riesgos = colors.HexColor("#30a907")
            fondo_filas_riesgos  = colors.HexColor("#F2F2F2")
            
            draw_tabla_pareto(
                canvas,
                "Riesgos Sociales",
                tabla_riesgos,
                x_der,
                y_base,
                text_color=color_titulo_riesgos,
                header_color=fondo_titulo_riesgos,
                body_color=fondo_filas_riesgos
            )
            
            draw_cantidad(
                canvas,
                f"Total: {cantidad_riesgos}",
                x_der + 90,
                y_base - 365
            )

        elif doc.page == 10:
            header_footer(canvas, doc)
        
            page_width, page_height = A4
        
            # ===== IMAGEN MICMAC =====
            img_width = 650
            img_height = 420
            img_x = (page_width - img_width) / 2
            img_y = page_height - img_height - 115
        
            canvas.drawImage(
                "assets/micmac.png",
                img_x,
                img_y,
                width=img_width,
                height=img_height,
                preserveAspectRatio=True,
                mask="auto"
            )

            # ===== COORDENADAS DE CUADRANTES =====
            quad_w = img_width / 2 - 80 #-- aqui puedo mover a la derecha los datos micmac
            quad_h = img_height / 2 - 20
        
            x_left  = img_x + 90 #---- aqui los hago para el centro
            x_right = img_x + img_width / 2 + 5
        
            y_top    = img_y + img_height - 70 #----- aqui para abajo entre mas grnade mas abajo
            y_bottom = img_y + img_height / 2 - 20
        
            # ===== PODER =====
            draw_micmac_lista(canvas, micmac_poder, x_left, y_top, quad_w)
        
            # ===== CONFLICTO =====
            draw_micmac_lista(canvas, micmac_conflicto, x_right, y_top, quad_w)
        
            # ===== AUTÓNOMAS =====
            draw_micmac_lista(canvas, micmac_autonomas, x_left, y_bottom, quad_w)
        
            # ===== RESULTADOS =====
            draw_micmac_lista(canvas, micmac_resultados, x_right, y_bottom, quad_w)

            #--------------------------micmac2-----------------------------------------

            header_footer(canvas, doc)

            page_width, page_height = A4

            # ===== IMAGEN MICMAC2 =====
            img_width = 520
            img_height = 260
            img_x = (page_width - img_width) / 2
            img_y = 50   # respeta footer

            canvas.drawImage(
                "assets/micmac2.png",
                img_x,
                img_y,
                width=img_width,
                height=img_height,
                preserveAspectRatio=True,
                mask="auto"
            )

            # ===== VARIABLES DE POSICIÓN =====
            tabla_width = 130
            tabla_font = 8

            x_tabla_riesgos = img_x + img_width * 0.55 - tabla_width / 2
            x_tabla_delitos = img_x + img_width * 0.85 - tabla_width / 2
            y_tablas = img_y + img_height - 80

            draw_tabla_overlay(
                canvas,
                tabla_riesgos_micmac2,
                x_tabla_riesgos,
                y_tablas,
                width=tabla_width,
                font_size=tabla_font
            )

            draw_tabla_overlay(
                canvas,
                tabla_delitos_micmac2,
                x_tabla_delitos,
                y_tablas,
                width=tabla_width,
                font_size=tabla_font
            )

            # ===== DATOS SUELTOS =====
            x_centro = img_x + img_width / 4
            y_centro = img_y + img_height / 2

            #____________P.Priorizadas
            
            draw_texto_overlay(
                canvas,
                cantidad_problematicas,
                x_centro - 30 ,
                y_centro - 15,
                size=30,
                color=colors.white
            )

            #--------------R.sociales
            draw_texto_overlay(
                canvas,
                riesgos_total,
                x_centro + 155,
                y_centro + 80,
                size=18,
                color=colors.white
            )

            #---------------delitos
            draw_texto_overlay(
                canvas,
                delitos_total,
                x_centro + 290,
                y_centro + 80,
                size=18,
                color=colors.white
            )

        elif doc.page == 11:
            header_footer(canvas, doc)
            page_width, page_height = A4
        
            # ===== TEXTO IZQUIERDA =====
            draw_texto_mixto(
                canvas,
                x=45,                                # ← mueve derecha/izquierda
                y=page_height - 210,                # ← sube/baja
                texto_antes="Frente a lo anterior, esta metodología permitió la identificación de",
                valor_1=causas_identificadas,
                texto_medio="causas, directamente relacionadas con los",
                valor_2=factores_micmac,
                texto_despues="factores priorizados en la Mic-Mac.",
                width=260,
                font_size=10,
                valor_size=16
            )
        
            # ===== IMAGEN TRIÁNGULO =====
            img_width = 260
            img_height = 260
        
            img_x = page_width / 2 + 10         # mueve derecha/izquierda
            img_y = page_height - img_height - 130  # mueve arriba/abajo
        
            canvas.drawImage(
                "assets/triangulo.png",
                img_x,
                img_y,
                width=img_width,
                height=img_height,
                preserveAspectRatio=True,
                mask="auto"
            )
        
            canvas.setFont("Helvetica-Bold", 18)
            canvas.setFillColor(colors.black)
        
            # DIRECTA (arriba)
            canvas.drawCentredString(
                img_x + img_width / 2,
                img_y + img_height - 50,
                str(triangulo_directa)
            )
        
            # SOCIOCULTURAL (abajo izquierda)
            canvas.drawCentredString(
                img_x + 40,
                img_y + 45,
                str(triangulo_sociocultural)
            )
        
            # ESTRUCTURAL (abajo derecha)
            canvas.drawCentredString(
                img_x + img_width - 40,
                img_y + 45,
                str(triangulo_estructural)
            )

            # ===== TABLA LISTA DE INSTITUCIONES =====
            draw_tabla_instituciones(
                canvas,
                tabla_instituciones,
                x=60,        # izquierda / derecha
                y=360,       # mitad inferior de la página
                table_width=480,
                header_bg=colors.HexColor("#30a907"),
                body_bg=colors.HexColor("#F2F2F2"),
                font_header=13,
                font_body=10
            )

        elif doc.page == 12:
            FullImage("assets/estadistica.png")(canvas, doc)

        elif doc.page == 13:
            header_footer(canvas, doc)
            page_width, page_height = A4        
       
            # ================== GRÁFICO CIRCULAR ==================
            canvas.drawImage(
                grafico_denuncias_path,
                x=(page_width - 550)/ 2,
                y=page_height - 330, #altura, a - mas altura
                width=250,
                height=250,
                preserveAspectRatio=True,
                mask="auto"
            )
        
            # ================== TABLA ==================
            draw_tabla_simple(
                canvas=canvas,
                data=tabla_denuncias,
                titulo="Detalle de denuncias por distrito",
                x=347,
                y=page_height - 310,
                col_widths=[100],
                header_color=colors.HexColor("#4472C4"),
                font_size_header=12,
                font_size_body=10
            )
        
            # ================== CUADRO TOTAL ==================
            canvas.setFillColor(colors.HexColor("#013051"))
            canvas.rect(
                page_width / 2 - 75,
                531,  #Altura + es mas
                115,
                50,
                fill=1,
                stroke=0
            )
        
            canvas.setFillColor(colors.white)
            canvas.setFont("Helvetica-Bold", 12)
            canvas.drawCentredString(
                page_width / 2 - 20,
                540,
                "Total de denuncias"
            )
        
            canvas.setFont("Helvetica-Bold", 22)
            canvas.drawCentredString(
                page_width / 2 - 20,
                555, #todos estos los aumente de 675
                str(total_denuncias)
            )

        ##______________________________GRAFICO CON HORARIOS Y TABLA_______________
            # ===== GRAFICO HORARIO =====
            canvas.drawImage(
                grafico_horario_path,
                x=(page_width - 550)/ 2,
                y=page_height - 590, # altura a - mas
                width=250,
                height=250,
                preserveAspectRatio=True,
                mask="auto"
            )

           #tabla
            draw_tabla_simple(
                canvas=canvas,
                data=tabla_horario,
                titulo="Denuncias por horario",
                x=420,
                y=page_height - 550,
                col_widths=[90, 40],
                header_color=colors.HexColor("#4472C4"),
                font_size_header=12,
                font_size_body=10
            ) 

            # ===== CUADRO AM =====
            canvas.setFillColor(colors.HexColor("#013051"))
            canvas.rect(280, 360, 100, 40, fill=1, stroke=0)
            
            canvas.setFillColor(colors.white)
            canvas.setFont("Helvetica-Bold", 12)
            canvas.drawCentredString(330, 385, "AM")
            
            canvas.setFont("Helvetica-Bold", 18)
            canvas.drawCentredString(330, 368, str(total_am))
            
            
            # ===== CUADRO PM =====
            canvas.setFillColor(colors.HexColor("#013051"))
            canvas.rect(280, 300, 100, 40, fill=1, stroke=0)
            
            canvas.setFillColor(colors.white)
            canvas.setFont("Helvetica-Bold", 12)
            canvas.drawCentredString(330, 325, "PM")
            
            canvas.setFont("Helvetica-Bold", 18)
            canvas.drawCentredString(330, 308, str(total_pm))

            ##tabla grande
            # ===== TABLA GRANDE REORDENADA =====
            tabla_data = tabla_horario_distrito[1:]  # datos sin encabezado
            encabezado_tabla = tabla_horario_distrito[0]
            
            # Ajuste dinámico de ancho
            total_columnas = len(encabezado_tabla)
            ancho_total = page_width - 80
            ancho_columna = ancho_total / total_columnas
            
            draw_tabla_horario_distrito(
                canvas=canvas,
                data=tabla_horario_distrito,
                titulo="DCLP según horario, por distrito",
                x=40,
                y=275,
                col_widths=[ancho_columna] * total_columnas
            )

        elif doc.page == 14:
            header_footer(canvas, doc)
            
            page_width, page_height = A4
            
            # ===== VARIABLES DE POSICION =====
            IMG_WIDTH = 500
            IMG_HEIGHT = 300
            
            POS_X = (page_width - IMG_WIDTH) / 2
            POS_Y = page_height - IMG_HEIGHT - 110
            # ================================
            
            canvas.drawImage(
                grafico_p14_path,
                POS_X,
                POS_Y,
                width=IMG_WIDTH,
                height=IMG_HEIGHT,
                preserveAspectRatio=True,
                mask="auto"
            )

       # ===== TABLA MODALIDADES =====
            TOTAL_COLUMNAS = len(tabla_p14[0])
                
            ANCHO_TOTAL = page_width - 60
            ANCHO_COLUMNA = ANCHO_TOTAL / TOTAL_COLUMNAS
                
            draw_tabla_modalidades_p14(
                canvas=canvas,
                data=tabla_p14,
                titulo="Frecuencia de modalidades por distrito",
                x=30,
                y=POS_Y - 40,   # 👈 Ajustable para subir/bajar
                col_widths=[ANCHO_COLUMNA] * TOTAL_COLUMNAS
            )

        elif doc.page == 15:

            header_footer(canvas, doc)
        
            page_width, page_height = A4
        
            # ===== VARIABLES CONFIGURABLES =====
            IMG_WIDTH = 450
            IMG_HEIGHT = 300
        
            POS_X = (page_width - IMG_WIDTH) / 2
            POS_Y = page_height - IMG_HEIGHT - 100
            # ===================================
        
            canvas.drawImage(
                grafico_p15_path,
                POS_X,
                POS_Y,
                width=IMG_WIDTH,
                height=IMG_HEIGHT,
                preserveAspectRatio=True,
                mask="auto"
            )

            #tabla
        # ===== TABLA DIAS VS DISTRITOS =====

            TOTAL_COLUMNAS = len(tabla_p15[0])
            
            ANCHO_TOTAL = page_width - 60
            ANCHO_COLUMNA = ANCHO_TOTAL / TOTAL_COLUMNAS
            
            draw_tabla_dias_distritos_p15(
                canvas=canvas,
                data=tabla_p15,
                titulo="Frecuencia por distrito según día",
                x=30,
                y=POS_Y - 60,  # 👈 Ajustable
                col_widths=[ANCHO_COLUMNA] * TOTAL_COLUMNAS
            )
            

        elif doc.page == 16:
            FullImage("assets/lineas.png")(canvas, doc)  

        elif doc.page == 17:

            header_footer(canvas, doc)
        
            page_width, page_height = A4
        
            # ================= VARIABLES CONFIGURABLES =================
        
            # Imagen base
            IMG_PATH = "assets/lins.png"
            IMG_WIDTH = 700
            IMG_HEIGHT = 400
            IMG_X = (page_width - IMG_WIDTH) / 2
            IMG_Y = 100
        
            # TOTAL LINEAS
            TOTAL_FONT = "Helvetica-Bold"
            TOTAL_SIZE = 60
            TOTAL_COLOR = colors.white
            TOTAL_X = 130 #bien
            TOTAL_Y = 185
        
            # Logo municipalidad
            LOGO_WIDTH = 175
            LOGO_HEIGHT = 175
            LOGO_X = 400
            LOGO_Y = 320 #355
        
            # Texto lineas municipales
            TEXT_FONT = "Helvetica-Bold"
            TEXT_SIZE = 40
            
            COLOR_MUNICIPAL = colors.HexColor("#30A907")  # Verde
            COLOR_FP = colors.white                       # Blanco
            COLOR_MIXTAS = colors.white     # blanco
        
            MUNICIPAL_X = 330 #380
            MUNICIPAL_Y = 370
            
            FP_X = 330
            FP_Y = 170 # bien
            
            MIXTAS_X = 375
            MIXTAS_Y = 280 #280
        
            # ===========================================================
        
            # ===== DIBUJAR IMAGEN BASE =====
            canvas.drawImage(
                IMG_PATH,
                IMG_X,
                IMG_Y,
                width=IMG_WIDTH,
                height=IMG_HEIGHT,
                preserveAspectRatio=True,
                mask="auto"
            )
        
            # ===== TOTAL LINEAS =====
            canvas.setFont(TOTAL_FONT, TOTAL_SIZE)
            canvas.setFillColor(colors.white)
            canvas.drawCentredString(
                TOTAL_X,
                TOTAL_Y,
                str(total_lineas)
            )
        
            # ===== LOGO MUNICIPALIDAD =====
            canvas.drawImage(
                logo_muni_path,
                LOGO_X,
                LOGO_Y,
                width=LOGO_WIDTH,
                height=LOGO_HEIGHT,
                preserveAspectRatio=True,
                mask="auto"
            )
        
            # ===== LINEAS MUNICIPALIDAD =====
            canvas.setFont(TEXT_FONT, TEXT_SIZE)
            canvas.setFillColor(COLOR_MUNICIPAL)
            
            canvas.drawString(
                MUNICIPAL_X,
                MUNICIPAL_Y,
                f"{lineas_municipalidad}"
            )
        
            # ===== LINEAS FP =====
            canvas.drawString(
                FP_X,
                FP_Y,
                f"{lineas_fp}"
            )
        
            # ===== LINEAS MIXTAS (CONDICIONAL) =====
            if lineas_mixtas is not None:
                canvas.setFillColor(COLOR_MIXTAS)
                canvas.drawString(
                    MIXTAS_X,
                    MIXTAS_Y,
                    f"Mixtas: {lineas_mixtas}"
                )

        elif doc.page >= 18:

            index = (doc.page - 18) // 2
        
            if index < len(lineas_accion_data):
                            
                linea = lineas_accion_data[index]
                page_width, page_height = A4
        
                # =====================================
                # PAGINA PAR → PORTADA
                # =====================================
                if (doc.page - 18) % 2 == 0:
        
                    canvas.drawImage(
                        "assets/la.png",
                        0,
                        0,
                        width=page_width,
                        height=page_height,
                        preserveAspectRatio=True,
                        mask="auto"
                    )
        
                    # NUMERO
                    canvas.setFont("Helvetica-Bold", 150)
                    canvas.setFillColor(colors.white)
                    canvas.drawString(400, page_height - 300, f"{linea['numero']}")
        
                    # TITULO
                    titulo_style = ParagraphStyle(
                        name="TituloLinea",
                        fontName="Helvetica-Bold",
                        fontSize=18,
                        textColor=colors.white,
                        alignment=TA_CENTER
                    )
        
                    TITULO_WIDTH = page_width * 0.7
                    TITULO_X = (page_width - TITULO_WIDTH) / 2
                    TITULO_Y = page_height - 765
        
                    texto_titulo = "<br/>".join(linea["problematicas"])
                    p = Paragraph(texto_titulo, titulo_style)
                    w, h = p.wrap(TITULO_WIDTH, 500)
        
                    p.drawOn(canvas, TITULO_X, TITULO_Y - (h / 2))
        
                    # LOGOS
                    if linea["corresponsable"] == "Fuerza Publica":
        
                        canvas.drawImage(
                            "assets/fp.png",
                            380,
                            page_height - 675,
                            width=200,
                            height=200,
                            mask="auto"
                        )
        
                    elif linea["corresponsable"] == "Municipalidad":
        
                        canvas.drawImage(
                            logo_muni_path,
                            380,
                            page_height - 675,
                            width=200,
                            height=200,
                            mask="auto"
                        )
        
                    elif linea["corresponsable"] == "Mixta":
        
                        canvas.drawImage(
                            "assets/fp.png",
                            380,
                            page_height - 675,
                            width=200,
                            height=200,
                            mask="auto"
                        )
        
                        canvas.drawImage(
                            logo_muni_path,
                            200,
                            page_height - 675,
                            width=200,
                            height=200,
                            mask="auto"
                        )
        
                # =====================================
                # PAGINA IMPAR → INTERNA
                # =====================================
                else:
        
                    header_footer(canvas, doc)
                    draw_pagina_linea_accion(canvas, doc, linea)

        else:
            header_footer(canvas, doc)

    doc.build(
        story,
        onFirstPage=first_page,
        onLaterPages=later_pages
    )

    buffer.seek(0)
    return buffer
