from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Image, Spacer, PageBreak
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

def generar_pdf(
    ruta_pdf,
    data,
    graficos,
    assets_path="assets"
):
    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    story = []

    # ========= PORTADA =========
    story.append(Image(f"{assets_path}/portada.png", width=595, height=842))
    story.append(PageBreak())

    # ========= INTRO =========
    story.append(Image(f"{assets_path}/intro.png", width=595, height=842))
    story.append(PageBreak())

    story.append(Paragraph("Introducción", styles["Heading1"]))
    story.append(Paragraph(data["introduccion"], styles["Normal"]))
    story.append(Spacer(1, 12))

    story.append(Image(f"{assets_path}/conformacion.png", width=400, height=300))
    story.append(PageBreak())

    # ========= PARTICIPACIÓN =========
    story.append(Image(f"{assets_path}/participacion.png", width=595, height=842))
    story.append(PageBreak())

    story.append(Paragraph("Datos de participación", styles["Heading1"]))
    story.append(Image(graficos["relacion"], width=400, height=300))
    story.append(Spacer(1, 12))
    story.append(Image(graficos["edad"], width=400, height=300))
    story.append(PageBreak())

    # ========= PERCEPCIÓN =========
    story.append(Image(f"{assets_path}/percepcion.png", width=595, height=842))
    story.append(PageBreak())

    story.append(Paragraph("Percepción ciudadana", styles["Heading1"]))
    story.append(Image(graficos["seguridad"], width=400, height=300))
    story.append(PageBreak())

    # ========= FINAL =========
    story.append(Image(f"{assets_path}/final.png", width=595, height=842))

    doc.build(story)
