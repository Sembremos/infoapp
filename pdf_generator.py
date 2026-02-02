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
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY
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

    TABLE_WIDTH = sum(col_widths)

    table_data = [[titulo] + [""] * (len(data[0]) - 1)]
    table_data.extend(data)

    table = Table(table_data, colWidths=col_widths)

    style = [
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#013051")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), font_size_header),
        ("GRID", (0, 0), (-1, -1), 1, border_color),
        ("BACKGROUND", (0, 1), (-1, -1), body_color),
        ("FONTSIZE", (0, 1), (-1, -1), font_size_body),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]

    table.setStyle(TableStyle(style))
    table.wrapOn(canvas, TABLE_WIDTH, 200)
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
    titulo="Lista de instituciones",
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
    tabla_instituciones
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
                y=220,       # mitad inferior de la página
                table_width=480,
                header_bg=colors.HexColor("#30a907"),
                body_bg=colors.HexColor("#F2F2F2"),
                font_header=13,
                font_body=10
            )
        


        else:
            header_footer(canvas, doc)

    doc.build(
        story,
        onFirstPage=first_page,
        onLaterPages=later_pages
    )

    buffer.seek(0)
    return buffer
