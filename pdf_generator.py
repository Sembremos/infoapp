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
            0, 0,
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

    # Página 2 (intro)
    story.append(PageBreak())

    # Página 3 en adelante (contenido)
    story.append(PageBreak())
    story.append(Paragraph("Introducción", styles["Heading1"]))
    story.append(Paragraph("Esto es la intro", styles["Normal"]))

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

    doc.build(
        story,
        onFirstPage=first_page,
        onLaterPages=later_pages
    )

    buffer.seek(0)
    return buffer



