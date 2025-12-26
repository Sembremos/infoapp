import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import base64
from io import BytesIO
from pathlib import Path

# ================== CONFIG ==================
st.set_page_config(page_title="Generador de Informes", layout="wide")
st.title("Generador de Informes")

# ================== FUNCIONES BASE ==================
def img_to_base64(path: Path) -> str:
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()

def fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=200)
    plt.close(fig)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()

def limpiar_df(labels, values):
    df = pd.DataFrame({"label": labels, "value": values})
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna()
    return df

def grafico_pie(labels, values, title):
    df = limpiar_df(labels, values)
    if df.empty or df["value"].sum() == 0:
        return ""
    fig, ax = plt.subplots()
    ax.pie(df["value"], labels=df["label"], autopct="%1.0f%%", startangle=90)
    ax.axis("equal")
    ax.set_title(title)
    return fig_to_base64(fig)

def grafico_barras(labels, values, title):
    df = limpiar_df(labels, values)
    if df.empty:
        return ""
    fig, ax = plt.subplots()
    ax.bar(df["label"], df["value"])
    ax.set_ylim(0, 100)
    ax.set_ylabel("Porcentaje")
    ax.set_title(title)
    plt.xticks(rotation=20)
    return fig_to_base64(fig)

# ================== CARGA HTML / CSS ==================
html = Path("Informe.html").read_text(encoding="utf-8")
css = Path("estilos.css").read_text(encoding="utf-8")

# ================== IMÁGENES ESTÁTICAS ==================
imagenes = [
    "portada.png", "intro.png", "conformacion.png", "participacion.png",
    "metodologico.png", "estadistica.png", "lineas.png",
    "percepcion.png", "final.png"
]

for img in imagenes:
    p = Path(img)
    if p.exists():
        html = html.replace(f'src="{img}"', f'src="{img_to_base64(p)}"')

# ================== SUBIR EXCEL ==================
archivo_excel = st.file_uploader("Subir matriz Excel", type=["xlsx"])

if archivo_excel:
    part = pd.read_excel(archivo_excel, sheet_name="participacion", header=None)
    rel = pd.read_excel(archivo_excel, sheet_name="relevante", header=None)

    # ================== PARTICIPACIÓN ==================
    # Tabla distrito
    html = html.replace(
        "{{tabla_distritos}}",
        f"""
        <tr>
          <td>{part.iloc[33,1]}</td>
          <td>{float(part.iloc[33,2])*100:.0f}%</td>
          <td>{int(part.iloc[33,3])}</td>
        </tr>
        """
    )

    # Relación
    html = html.replace(
        "{{grafico_relacion}}",
        grafico_barras(
            ["Comunidad", "Comercio", "Fuerza Pública"],
            part.iloc[33,4:7] * 100,
            "Relación de participación"
        )
    )

    # Edad
    html = html.replace(
        "{{grafico_edad}}",
        grafico_pie(
            part.iloc[1:6,1],
            part.iloc[1:6,2] * 100,
            "Participación por edad"
        )
    )

    # Escolaridad
    html = html.replace(
        "{{grafico_escolaridad}}",
        grafico_pie(
            part.iloc[11:18,1],
            part.iloc[11:18,2] * 100,
            "Participación por escolaridad"
        )
    )

    # Género
    html = html.replace(
        "{{grafico_genero}}",
        grafico_pie(
            part.iloc[23:26,1],
            part.iloc[23:26,2] * 100,
            "Participación por género"
        )
    )

    # ================== PERCEPCIÓN CIUDADANA ==================
    html = html.replace(
        "{{grafico_seguridad_comunidad}}",
        grafico_pie(
            rel.iloc[1:4,0],
            rel.iloc[1:4,1],
            "¿Se siente seguro en su comunidad?"
        )
    )

    html = html.replace(
        "{{grafico_seguridad_barrio}}",
        grafico_barras(
            rel.iloc[6:9,0],
            rel.iloc[6:9,1],
            "Seguridad en el barrio este año"
        )
    )

    tabla_seguridad = ""
    for i in range(1,4):
        tabla_seguridad += f"""
        <tr>
          <td>{rel.iloc[i,0]}</td>
          <td>{rel.iloc[i,1]}%</td>
        </tr>
        """
    html = html.replace("{{tabla_seguridad}}", tabla_seguridad)

    # ================== VICTIMIZACIÓN ==================
    html = html.replace(
        "{{grafico_victima_denuncia}}",
        grafico_barras(
            rel.iloc[12:14,0],
            rel.iloc[12:14,1],
            "Victimización y denuncia"
        )
    )

    html = html.replace(
        "{{grafico_no_denuncia}}",
        grafico_barras(
            rel.iloc[15:18,0],
            rel.iloc[15:18,1],
            "Motivos de no denuncia"
        )
    )

    # ================== DINÁMICA DELICTIVA ==================
    html = html.replace(
        "{{grafico_horario}}",
        grafico_pie(
            rel.iloc[20:24,0],
            rel.iloc[20:24,1],
            "Horario delictivo"
        )
    )

    html = html.replace(
        "{{grafico_armas}}",
        grafico_pie(
            rel.iloc[25:29,0],
            rel.iloc[25:29,1],
            "Uso de armas en hechos delictivos"
        )
    )

    # ================== SERVICIO POLICIAL ==================
    html = html.replace(
        "{{grafico_servicio_policial}}",
        grafico_barras(
            rel.iloc[31:35,0],
            rel.iloc[31:35,1],
            "Percepción del servicio policial"
        )
    )

    html = html.replace(
        "{{grafico_calificacion}}",
        grafico_pie(
            rel.iloc[36:40,0],
            rel.iloc[36:40,1],
            "Calificación del servicio policial"
        )
    )

    html = html.replace(
        "{{grafico_conoce_policias}}",
        grafico_pie(
            rel.iloc[41:43,0],
            rel.iloc[41:43,1],
            "¿Conoce a los policías de su comunidad?"
        )
    )

    html = html.replace(
        "{{grafico_conversado}}",
        grafico_pie(
            rel.iloc[44:46,0],
            rel.iloc[44:46,1],
            "¿Ha conversado con ellos?"
        )
    )

    # ================== SECTOR COMERCIAL ==================
    html = html.replace(
        "{{grafico_seguridad_local}}",
        grafico_pie(
            rel.iloc[48:50,0],
            rel.iloc[48:50,1],
            "¿Se siente seguro en su local comercial?"
        )
    )

    html = html.replace(
        "{{grafico_conoce_programa}}",
        grafico_pie(
            rel.iloc[51:53,0],
            rel.iloc[51:53,1],
            "¿Conoce el programa Seguridad Comercial?"
        )
    )

    html = html.replace(
        "{{grafico_inscrito_programa}}",
        grafico_pie(
            rel.iloc[54:56,0],
            rel.iloc[54:56,1],
            "¿Está inscrito en el programa?"
        )
    )

    html = html.replace(
        "{{grafico_contacto_programa}}",
        grafico_pie(
            rel.iloc[57:59,0],
            rel.iloc[57:59,1],
            "¿Le gustaría ser contactado?"
        )
    )

# ================== INYECTAR CSS ==================
html = html.replace("</head>", f"<style>{css}</style></head>")

# ================== PREVIEW ==================
st.subheader("Vista previa del informe")
st.components.v1.html(html, height=1000, scrolling=True)

# ================== DESCARGA ==================
st.download_button(
    "⬇️ Descargar HTML (listo para PDF)",
    html,
    "Informe.html",
    "text/html"
)
