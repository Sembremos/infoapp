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
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
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


# ================= GRAFICO DE EDAD =================
def draw_grafico_edad(canvas, doc, grafico_edad_path):
    page_width, page_height = A4

    # Tamaño deseado (ancho)
    img_width = 220

    # Leer imagen real
    img = ImageReader(grafico_edad_path)
    img_w, img_h = img.getSize()

    # Calcular alto proporcional
    img_height = img_width * img_h / img_w

    # Posición: tercio superior izquierdo
    x = 40
    y = page_height - img_height - 120

    canvas.drawImage(
        grafico_edad_path,
        x,
        y,
        width=img_width,
        height=img_height,
        preserveAspectRatio=True,
        mask="auto"
    )


#TABLA DE EDAD================================================///////////
def draw_tabla_edad(canvas, doc, tabla_edad):
    page_width, page_height = A4

    # ================= CONFIGURACIÓN =================
    TABLE_WIDTH = 220        # mismo ancho visual que el gráfico
    FONT_SIZE_HEADER = 12    # 👈 puedes cambiar tamaño aquí
    FONT_SIZE_BODY = 11      # 👈 y aquí
    X = page_width / 2 + 10  # lado derecho
    Y = page_height - 60    # alineado con gráfico

    # ================= CONTENIDO =================
    data = [["Participación por Edad", ""]]
    data.extend(tabla_edad)

    table = Table(
        data,
        colWidths=[TABLE_WIDTH * 0.6, TABLE_WIDTH * 0.4]
    )

    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (0, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DEEBF7")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(000000)),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), FONT_SIZE_HEADER),

        ("GRID", (0, 1), (-1, -1), 0.5, colors.white),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 1), (-1, -1), FONT_SIZE_BODY),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFFFFF")),
    ]))

    table.wrapOn(canvas, TABLE_WIDTH, 200)
    table.drawOn(canvas, X, Y - 200)

#________---------___________----------__________BLOQUE DE ESCOLARIDAD**********************************
    # Intervalos / niveles (A39:A46)
escolaridad_labels = df.iloc[38:46, 0].astype(str)

# Porcentajes (B39:B46)
escolaridad_percent_values = (
    df.iloc[38:46, 1]
    .astype(str)
    .str.replace("%", "", regex=False)
    .str.replace(",", ".", regex=False)
)

escolaridad_percent_values = pd.to_numeric(escolaridad_percent_values, errors="coerce")

# Filtrar filas válidas
mask = escolaridad_percent_values.notna()
escolaridad_labels = escolaridad_labels[mask]
escolaridad_percent_values = escolaridad_percent_values[mask]

# -------- GRÁFICO --------
fig_esco, ax_esco = plt.subplots(figsize=(7, 7))

colores_esco = [
    "#5B9BD5",
    "#A5A5A5",
    "#4472C4",
    "#255E91",
    "#B7B7B7",
    "#9DC3E6",
    "#8FAADC",
    "#D9E1F2"
]

wedges, texts, autotexts = ax_esco.pie(
    escolaridad_percent_values,
    labels=escolaridad_labels,
    autopct=lambda p: f"{p:.0f}%",
    pctdistance=0.6,
    labeldistance=1.35,
    startangle=90,
    colors=colores_esco,
    textprops={"fontsize": 20}
)

ax_esco.axis("equal")

for text in texts:
    text.set_fontsize(20)

for autotext in autotexts:
    autotext.set_fontsize(25)

ax_esco.set_facecolor("none")
fig_esco.patch.set_alpha(0)

buf_esco = BytesIO()
fig_esco.savefig(
    buf_esco,
    format="png",
    bbox_inches="tight",
    pad_inches=0,
    dpi=200,
    transparent=True
)

plt.close(fig_esco)
buf_esco.seek(0)

grafico_escolaridad_path = BASE_DIR / "grafico_participacion_escolaridad.png"
with open(grafico_escolaridad_path, "wb") as f:
    f.write(buf_esco.getbuffer())

# -------- TABLA --------
tabla_escolaridad = df.iloc[38:46, 0:2].fillna("").values.tolist()

#"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
#___________________________________________________________________________________________________________________________
#"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""GENERADOR DE PDF""""""""""""""""""""""""""""""""""""""""""""""
def generar_pdf(
    portada_path,
    grafico_relacion_path,
    grafico_edad_path,
    grafico_escolaridad_path=str(grafico_escolaridad_path), 
    delegacion,
    codigo,
    tabla_participacion,
    tabla_edad,
    tabla_escolaridad=tabla_escolaridad 
):
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
    story.append(PageBreak())  # imagen sola
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

    # grafico de relacion
    story.append(Spacer(1, 25))
    story.append(Paragraph("Relación por distrito", styles["Heading2"]))
    story.append(Spacer(1, 15))
    story.append(Image(grafico_relacion_path, width=400, height=250))

    # ================= PARTICIPACIÓN POR EDAD =================
    story.append(PageBreak())
    story.append(Spacer(1, 40))
    story.append(Paragraph("Datos de Participación", styles["Heading1"]))

    # ================= CONSTRUCCIÓN later pages =================
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
            draw_tabla_edad(canvas, doc, tabla_edad)

        else:
            header_footer(canvas, doc)

    doc.build(
        story,
        onFirstPage=first_page,
        onLaterPages=later_pages
    )

    buffer.seek(0)
    return buffer
