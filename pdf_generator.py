from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

def generar_pdf(ruta_pdf, portada_path, grafico_path):
    styles = getSampleStyleSheet()

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    story = []

    # PORTADA (AJUSTADA AL FRAME REAL)
    story.append(
        Image(
            portada_path,
            width=doc.width - 10,
            height=doc.height - 10
        )
    )
    story.append(PageBreak())

    # CONTENIDO
    story.append(Paragraph("Datos de participación", styles["Heading1"]))
    story.append(
        Image(
            grafico_path,
            width=400,
            height=300
        )
    )

    doc.build(story)
