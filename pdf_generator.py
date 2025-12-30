from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Spacer
from io import BytesIO
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_LEFT


# ================= UTILIDAD FULL PAGE =================
def FullImage(path):
    def draw(canvas, doc):
        w, h = A4
        canvas.drawImage(
            path,
            0, 0,
            width=w,
            height=h,
            preserveAspectRatio=True,
            mask="auto"
        )
    return draw

# HF

def header_footer(canvas, doc):
    page_width, page_height = A4

    # HEADER (
    canvas.drawImage(
        "assets/header.png",
        0,
        page_height - 80,
        width=page_width,
        height=60,
        mask="auto"
    )

    # FOOTER 
    canvas.drawImage(
        "assets/footer.png",
        0,
        0,
        width=page_width,
        height=60,
        mask="auto"
    )



# ================= GENERADOR PDF =================
def generar_pdf(portada_path, grafico_path, delegacion, codigo):
    buffer = BytesIO()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
    name="TituloGrande",
    fontName="Helvetica",
    fontSize=26,
    textColor=colors.HexColor("#FFFFFF"),
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
    textColor=colors.HexColor("#FFFFFF"),
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

    # Página 2 (intro)
    story.append(PageBreak())
    story.append(Paragraph(
        "DELEGACIÓN POLICIAL",
        styles["TituloGrande"]
))

    story.append(Paragraph(
        delegacion,
        styles["TituloDelta"]
))
    story.append(Paragraph(
        codigo,
        styles["TituloD2"]
))

    

    # Página 3 en adelante (contenido)
    story.append(PageBreak())
    story.append(Spacer(1,40))
    story.append(Paragraph("Introducción", styles["Heading1"]))
    story.append(Spacer(1,12))

    story.append(Paragraph(
        """Desde el año 2022, el Ministerio de Seguridad Pública ha implementado en todo el territorio nacional el Modelo Preventivo de Gestión Policial, una iniciativa estratégica destinada a fortalecer la seguridad pública a través de un enfoque proactivo y colaborativo. Una parte integral de este modelo es la Estrategia Integral de Prevención para la Seguridad Pública, conocida como "Sembremos Seguridad", que se centra en la contextualización de las dinámicas delincuenciales y sociales que afectan a nuestras comunidades.""",
        styles["Normal"]
    ))
    story.append(Spacer(1,20))

    
    story.append(Paragraph(
        """El presente informe, elaborado para el territorio que comprende la Delegación Policial de San Ramón, surge como una herramienta esencial para la toma efectiva de decisiones. Este informe se concibe como un instrumento dinámico y orientado hacia el futuro, diseñado para proporcionar información clave y un plan de trabajo estructurado que permita abordar las problemáticas prioritarias identificadas en el ámbito de la seguridad pública.""",
        styles["Normal"]
    ))
    story.append(Spacer(1,20))

    story.append(Image(
                "assets/conformacion.png",
                width=600,
                height=400
    ))

    story.append(PageBreak())

    # Índices
    story.append(PageBreak())
    story.append(Paragraph("Hola", styles["Heading1"]))

    story.append(PageBreak())
    story.append(Paragraph("Datos de participación", styles["Heading1"]))
    story.append(Image(grafico_path, width=400, height=300))

    # ===== CONTROL DE PÁGINAS =====
    def first_page(canvas, doc):
        FullImage(portada_path)(canvas, doc)

    def later_pages(canvas, doc):
        if doc.page == 2:
            FullImage("assets/intro.png")(canvas, doc)
        elif doc.page == 4:
            FullImage("assets/participacion.png")(canvas, doc)
        else:
            header_footer(canvas, doc)

    doc.build(
        story,
        onFirstPage=first_page,
        onLaterPages=later_pages
    )

    buffer.seek(0)
    return buffer
