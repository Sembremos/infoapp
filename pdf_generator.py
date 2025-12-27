from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO

def generar_pdf(portada_path, grafico_path):
    buffer = BytesIO()
    styles = getSampleStyleSheet()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    story = []

    # PORTADA (escalado seguro)
    portada = Image(portada_path)
    portada._restrictSize(doc.width, doc.height)
    story.append(portada)
    story.append(PageBreak())

    # CONTENIDO
    story.append(Paragraph("Datos de participación", styles["Heading1"]))

    grafico = Image(grafico_path)
    grafico._restrictSize(400, 300)
    story.append(grafico)

    doc.build(story)

    buffer.seek(0)
    return buffer
