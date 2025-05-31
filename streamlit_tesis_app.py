
import streamlit as st
import pandas as pd
import numpy as np
import os
import io
from sentence_transformers import SentenceTransformer, util
from PIL import Image

# Verificación de librería xlsxwriter
try:
    import xlsxwriter
except ImportError:
    st.error("La librería 'xlsxwriter' no está instalada. Ejecute 'pip install xlsxwriter' en su entorno.")
    st.stop()

st.set_page_config(page_title="Clasificador de Tesis", layout="centered")

st.markdown("""
## Clasificador de Capítulos de Tesis  
**Creado por la arquitecta María José Duarte Torres**

Esta aplicación permite analizar y clasificar los títulos o capítulos de una tesis según su relación con los **objetivos**, la **metodología** y el **marco teórico**. También detecta si un título **se repite en otro capítulo** o **puede resumirse o eliminarse**. Finalmente, exporta un archivo Excel con los resultados y colores que te ayudarán a decidir qué contenido conservar.

""")

# Carga de archivos
uploaded_file = st.file_uploader("📤 Sube tu archivo Excel con los títulos", type=["xlsx"])
objetivo = st.text_area("🎯 Pega el texto de los objetivos de tu tesis")
metodologia = st.text_area("🛠️ Pega el texto de la metodología")
marco = st.text_area("📚 Pega el texto del marco teórico")

modelo = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")

def obtener_similitud(titulo, texto_referencia):
    emb_titulo = modelo.encode(titulo, convert_to_tensor=True)
    emb_ref = modelo.encode(texto_referencia, convert_to_tensor=True)
    similitud = util.cos_sim(emb_titulo, emb_ref).item()
    return round(similitud, 3)

def convertir_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Clasificación")
        workbook = writer.book
        worksheet = writer.sheets["Clasificación"]

        formato_si = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        formato_no = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        columnas = ["¿Relaciona un objetivo?", "¿Es clave para entender metodología/resultados?", "¿Aporta al marco teórico?",
                    "¿Se repite en otro capítulo?", "¿Puede resumirse/eliminarse?"]

        for col in columnas:
            if col in df.columns:
                col_idx = df.columns.get_loc(col)
                worksheet.conditional_format(1, col_idx, len(df), col_idx, {"type": "text", "criteria": "containing", "value": "Sí", "format": formato_si})
                worksheet.conditional_format(1, col_idx, len(df), col_idx, {"type": "text", "criteria": "containing", "value": "No", "format": formato_no})

    output.seek(0)
    return output

if uploaded_file and objetivo and metodologia and marco:
    df = pd.read_excel(uploaded_file)
    if "Capítulo o título" not in df.columns:
        st.error("Tu archivo debe tener una columna llamada exactamente: Capítulo o título")
        st.stop()

    df["Capítulo o título"] = df["Capítulo o título"].astype(str)

    df["¿Relaciona un objetivo?"] = df["Capítulo o título"].apply(lambda x: "Sí" if obtener_similitud(x, objetivo) > 0.25 else "No")
    df["¿Es clave para entender metodología/resultados?"] = df["Capítulo o título"].apply(lambda x: "Sí" if obtener_similitud(x, metodologia) > 0.25 else "No")
    df["¿Aporta al marco teórico?"] = df["Capítulo o título"].apply(lambda x: "Sí" if obtener_similitud(x, marco) > 0.25 else "No")

    df["¿Se repite en otro capítulo?"] = df["Capítulo o título"].duplicated(keep=False).map({True: "Sí", False: "No"})

    def puede_eliminarse(row):
        return "Sí" if (row["¿Relaciona un objetivo?"] == "No" and
                        row["¿Es clave para entender metodología/resultados?"] == "No" and
                        row["¿Aporta al marco teórico?"] == "No") else "No"

    df["¿Puede resumirse/eliminarse?"] = df.apply(puede_eliminarse, axis=1)

    st.success("✅ Clasificación completada. Revisa los resultados abajo.")
    st.dataframe(df, use_container_width=True)

    archivo_excel = convertir_excel(df)
    st.download_button("📥 Descargar Excel Clasificado", data=archivo_excel, file_name="contenido_clasificado.xlsx")

    st.markdown("---")
    st.markdown("### Ejemplo de salida exportada:")
    st.image("icono_app.png", caption="Ejemplo visual de clasificación", use_container_width=True)
