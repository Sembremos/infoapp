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
    header_color=colors.HexColor("#4471C4"),
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
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
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


#infografia de datos

generar_infografia_datos(
    template_path=str(ASSETS_DIR / "datos.png"),
    output_path=str(infografia_datos_path),
    datos=datos_fuentes,
    config=config_infografia
)

# ================= GENERADOR PDF =================
def generar_pdf(
    portada_path,
    grafico_relacion_path,
    grafico_edad_path,
    grafico_escolaridad_path,
    grafico_genero_path,
    infografia_datos_path,
    delegacion,
    codigo,
    tabla_participacion,
    tabla_edad,
    tabla_escolaridad,
    tabla_genero,
    tabla_encuesta_comunidad, 
    tabla_otras_encuestas        
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
    story.append(Image("assets/netquest.png", width=600, height=175))

    story.append(Spacer(1, 5))
    story.append(Paragraph("__________________________________________________________________________________________"))
    

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
            #===============infografia datos========================
            canvas.drawImage(
                str(infografia_datos_path),
                50,                    # margen izquierdo
                page_height - 550,     # mitad inferior
                width=500,
                height=280,
                mask="auto"
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
