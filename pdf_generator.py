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

    # CONTENIDO (empieza en página 2)
    story.append(PageBreak())
    story.append(Paragraph("Datos de participación", styles["Heading1"]))

    grafico = Image(grafico_path, width=400, height=300)
    story.append(grafico)

    # ---------- PORTADA FULL PAGE ----------
    def portada_full(canvas, doc):
        width, height = A4
        canvas.drawImage(
            portada_path,
            0,
            0,
            width=width,
            height=height,
            preserveAspectRatio=True,
            mask='auto'
        )

    doc.build(
        story,
        onFirstPage=portada_full
    )

    buffer.seek(0)
    return buffer
