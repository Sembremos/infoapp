import streamlit as st
import pandas as pd

st.set_page_config(page_title="Generador de Informes")

st.title("Generador de Informes PDF")

archivo = st.file_uploader("Subir matriz Excel", type=["xlsx"])

if archivo:
    df = pd.read_excel(archivo)
    st.success("Excel cargado correctamente")
    st.dataframe(df.head())
