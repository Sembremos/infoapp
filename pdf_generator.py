from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from pathlib import Path

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

    # ===== PORTADA (AJUSTADA AL FRAME) =====
    story.append(
        Image(
            portada_path,
            width=doc.width,
            height=doc.height
        )
    )

    story.append(Spacer(1, 24))
    story.append(Paragraph("Datos de participación", styles["Heading1"]))
    story.append(Spacer(1, 12))

    story.append(
        Image(
            grafico_path,
            width=400,
            height=300
        )
    )

    doc.build(story)
