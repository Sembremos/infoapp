from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
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

# ================= GENERADOR PDF =================
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

    # ===== CONTENIDO (empieza después de la portada) =====
    story.append(PageBreak())

    # Introducción
    story.append(Paragraph("Introducción", styles["Heading1"]))
    story.append(Paragraph("Esto es la intro", styles["Normal"]))

    story.append(PageBreak())

    story.append(Paragraph("HOla", styles["Heading1"]))

    story.append(PageBreak())

    # Gráficos
    story.append(Paragraph("Datos de participación", styles["Heading1"]))
    grafico = Image(grafico_path, width=400, height=300)
    story.append(grafico)

    # ===== BUILD =====
    doc.build(
        story,
        onFirstPage=FullImage(portada_path)
    )

    buffer.seek(0)
    return buffer


