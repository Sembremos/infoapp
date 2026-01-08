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
from reportlab.lib.enums import TA_LEFT
from reportlab.lib import colors
from io import BytesIO
from reportlab.lib.enums import TA_JUSTIFY

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

### GRAFICO DE EDAD ##

def draw_grafico_edad(canvas, doc, grafico_edad_path):
    page_width, page_height = A4

    # Tamaño del gráfico (más pequeño)
    img_width = 220
    img_height = 220

    # Posición: tercio superior izquierdo
    x = 40  # margen izquierdo
    y = page_height - img_height - 120  # baja un poco desde arriba

    canvas.drawImage(
        grafico_edad_path,
        x,
        y,
        width=img_width,
        preserveAspectRatio=True,
        mask="auto"
    )


# ================= GENERADOR PDF =================
def generar_pdf(portada_path, grafico_relacion_path, grafico_edad_path, delegacion, codigo, tabla_participacion):
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

    # ================= PÁGINA 2 =================
    story.append(PageBreak())
    story.append(Paragraph("DELEGACIÓN POLICIAL", styles["TituloGrande"]))
    story.append(Paragraph(delegacion, styles["TituloDelta"]))
    story.append(Paragraph(codigo, styles["TituloD2"]))

    # ================= INTRODUCCIÓN =================
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

    # ================= PARTICIPACIÓN =================
    story.append(PageBreak()) ## esto lo HAGO PARA QUE QUEDE LA IMAGEN SOLITA 
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
#grafico de relacion
    story.append(Spacer(1, 25))
    story.append(Paragraph("Relación por distrito", styles["Heading2"]))
    story.append(Spacer(1, 15))
    story.append(Image(grafico_relacion_path, width=400, height=250))

    # ================= PARTICIPACIÓN POR EDAD =================
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Participación por Edad", styles["Heading1"]))

    # ================= CONSTRUCCIÓN =================
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

        else:
            header_footer(canvas, doc)

    doc.build(
        story,
        onFirstPage=first_page,
        onLaterPages=later_pages
    )

    buffer.seek(0)
    return buffer
