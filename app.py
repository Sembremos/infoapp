import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Image, Spacer, PageBreak
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

# ================= CONFIG =================
st.set_page_config(page_title="Generador de Informes PDF", layout="centered")
st.title("Generador de Informes PDF")

# ================= UTILIDADES =================
def crear_grafico_barras(labels, values, titulo):
    fig, ax = plt.subplots()
    ax.bar(labels, values)
    ax.set_title(titulo)
    ax.set_ylim(0, 100)
    plt.xticks(rotation=20)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf

def crear_grafico_pie(labels, values, titulo):
    fig, ax = plt.subplots()
    ax.pie(values, labels=labels, autopct="%1.0f%%", startangle=90)
    ax.set_title(titulo)
    ax.axis("equal")
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf

def generar_pdf(ruta_pdf, data, graficos):
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

    # ===== PORTADA =====
    story.append(Image("assets/portada.png", width=595, height=842))
    story.append(PageBreak())

    # ===== INTRO =====
    story.append(Image("assets/intro.png", width=595, height=842))
    story.append(PageBreak())

    story.append(Paragraph("Introducción", styles["Heading1"]))
    story.append(Paragraph(data["introduccion"], styles["Normal"]))
    story.append(Spacer(1, 12))
    story.append(Image("assets/conformacion.png", width=400, height=300))
    story.append(PageBreak())

    # ===== PARTICIPACIÓN =====
    story.append(Image("assets/participacion.png", width=595, height=842))
    story.append(PageBreak())

    story.append(Paragraph("Datos de participación", styles["Heading1"]))
    story.append(Image(graficos["relacion"], width=400, height=300))
    story.append(Spacer(1, 12))
    story.append(Image(graficos["edad"], width=400, height=300))
    story.append(PageBreak())

    # ===== PERCEPCIÓN =====
    story.append(Image("assets/percepcion.png", width=595, height=842))
    story.append(PageBreak())

    story.append(Paragraph("Percepción ciudadana", styles["Heading1"]))
    story.append(Image(graficos["seguridad"], width=400, height=300))
    story.append(PageBreak())

    # ===== FINAL =====
    story.append(Image("assets/final.png", width=595, height=842))

    doc.build(story)

# ================= APP =================
archivo = st.file_uploader("Subir matriz Excel", type=["xlsx"])

if archivo:
    try:
        part = pd.read_excel(
            archivo,
            sheet_name="participacion",
            header=None,
            engine="openpyxl"
        )

        rel = pd.read_excel(
            archivo,
            sheet_name="relevante",
            header=None,
            engine="openpyxl"
        )

        graficos = {
            "relacion": crear_grafico_barras(
                ["Comunidad", "Comercio", "Fuerza Pública"],
                part.iloc[33,4:7] * 100,
                "Relación de participación"
            ),
            "edad": crear_grafico_pie(
                part.iloc[1:6,1],
                part.iloc[1:6,2] * 100,
                "Participación por edad"
            ),
            "seguridad": crear_grafico_pie(
                rel.iloc[1:4,0],
                rel.iloc[1:4,1],
                "¿Se siente seguro en su comunidad?"
            )
        }

        data = {
            "introduccion": (
                "El presente informe tiene como objetivo analizar la "
                "participación comunitaria y la percepción ciudadana en "
                "materia de seguridad pública, a partir de la información "
                "recopilada en el territorio."
            )
        }

        if st.button("Generar PDF"):
            ruta_pdf = "Informe_Territorial.pdf"
            generar_pdf(ruta_pdf, data, graficos)

            with open(ruta_pdf, "rb") as f:
                st.download_button(
                    "⬇️ Descargar PDF",
                    f,
                    file_name="Informe_Territorial.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
