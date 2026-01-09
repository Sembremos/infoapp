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
from reportlab.lib.utils import ImageReader

# ================= CONSTANTES DE LAYOUT =================
MARGEN_IZQUIERDO = 40
MARGEN_SUPERIOR_GRAFICO = 120
GRAFICO_EDAD_ANCHO = 220

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


def draw_grafico_edad(canvas, doc, grafico_edad_path):
    page_width, page_height = A4

    img = ImageReader(grafico_edad_path)
    img_w, img_h = img.getSize()
    img_height = GRAFICO_EDAD_ANCHO * img_h / img_w

    x = MARGEN_IZQUIERDO
    y = page_height - img_height - MARGEN_SUPERIOR_GRAFICO

    canvas.drawImage(
        grafico_edad_path,
        x,
        y,
        width=GRAFICO_EDAD_ANCHO,
        height=img_height,
        preserveAspectRatio=True,
        mask="auto"
    )


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

    story.append(PageBreak())
    story.append(Paragraph("DELEGACIÓN POLICIAL", styles["TituloGrande"]))
    story.append(Paragraph(delegacion, styles["TituloDelta"]))
    story.append(Paragraph(codigo, styles["TituloD2"]))

    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Introducción", styles["Heading1"]))
    story.append(Spacer(1, 12))

    story.append(Paragraph(
        "Desde el año 2022, el Ministerio de Seguridad Pública ha implementado en todo el territorio nacional el Modelo Preventivo de Gestión Policial, una iniciativa estratégica destinada a fortalecer la seguridad pública a través de un enfoque proactivo y colaborativo.",
        styles["NormalJustificado"]
    ))

    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Datos de Participación", styles["Heading1"]))
    story.append(Spacer(1, 20))

    tabla = Table(tabla_participacion, colWidths=[180, 180, 120])
    tabla.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#013051")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))

    story.append(tabla)
    story.append(Spacer(1, 25))
    story.append(Paragraph("Relación por distrito", styles["Heading2"]))
    story.append(Spacer(1, 15))
    story.append(Image(grafico_relacion_path, width=400, height=250))

    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Datos de Participación", styles["Heading1"]))

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
