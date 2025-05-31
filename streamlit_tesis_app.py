import streamlit as st
import pandas as pd
import numpy as np
from sentence_transformers import SentenceTransformer, util
from io import BytesIO
from PIL import Image

st.set_page_config(page_title="Clasificador de Tesis", layout="centered")

st.title("📘 Clasificador de Capítulos de Tesis")
st.subheader("Creado por arquitecta María José Duarte Torres")

st.markdown("""
Esta aplicación permite clasificar automáticamente los capítulos o títulos de tu tesis según su relación con los **objetivos**, la **metodología** y el **marco teórico**. También detecta repeticiones entre capítulos y sugiere si pueden **resumirse o eliminarse**. El resultado incluye una clasificación automática de relevancia (A = muy relevante, E = poco relevante).

**📥 Instrucciones:**
1. Sube tu archivo de Excel (.xlsx) con una columna llamada `Capítulo o título`.
2. Escribe el texto completo de tus objetivos, metodología y marco teórico.
3. Descarga el archivo clasificado con colores y sugerencias.
""")

# Subir el archivo Excel
archivo_excel = st.file_uploader("📂 Sube tu archivo Excel con los títulos", type=["xlsx"])

# Entradas de texto
texto_objetivos = st.text_area("🎯 Objetivos de tu tesis", height=200)
texto_metodologia = st.text_area("🛠️ Metodología", height=200)
texto_marco = st.text_area("📚 Marco teórico", height=200)

if archivo_excel and texto_objetivos and texto_metodologia and texto_marco:
    df = pd.read_excel(archivo_excel)
    if "Capítulo o título" not in df.columns:
        st.error("⚠️ Asegúrate de que el archivo tenga una columna llamada 'Capítulo o título'")
    else:
        model = SentenceTransformer("all-MiniLM-L6-v2")
        titulos = df["Capítulo o título"].astype(str).tolist()
        embeddings = model.encode(titulos + [texto_objetivos, texto_metodologia, texto_marco], convert_to_tensor=True)

        emb_titulos = embeddings[:len(titulos)]
        emb_obj = embeddings[len(titulos)]
        emb_met = embeddings[len(titulos)+1]
        emb_marco = embeddings[len(titulos)+2]

        # Similitudes
        sim_obj = util.cos_sim(emb_titulos, emb_obj).cpu().numpy().flatten()
        sim_met = util.cos_sim(emb_titulos, emb_met).cpu().numpy().flatten()
        sim_marco = util.cos_sim(emb_titulos, emb_marco).cpu().numpy().flatten()

        def binarizar(sim, umbral=0.5):
            return ["Sí" if s >= umbral else "No" for s in sim]

        df["¿Relaciona un objetivo?"] = binarizar(sim_obj, 0.5)
        df["¿Es clave para entender metodología/resultados?"] = binarizar(sim_met, 0.5)
        df["¿Aporta al marco teórico?"] = binarizar(sim_marco, 0.5)

        # NUEVAS COLUMNAS
        df["¿Se repite en otro capítulo?"] = df["Capítulo o título"].duplicated(keep=False).map({True: "Sí", False: "No"})
        def puede_eliminarse(row):
            return "Sí" if (row["¿Relaciona un objetivo?"] == "No" and
                            row["¿Es clave para entender metodología/resultados?"] == "No" and
                            row["¿Aporta al marco teórico?"] == "No") else "No"
        df["¿Puede resumirse/eliminarse?"] = df.apply(puede_eliminarse, axis=1)

        # Clasificación A-E
        def clasificar(row):
            score = sum([
                row["¿Relaciona un objetivo?"] == "Sí",
                row["¿Es clave para entender metodología/resultados?"] == "Sí",
                row["¿Aporta al marco teórico?"] == "Sí"
            ])
            return ["E", "D", "C", "B", "A"][score]

        df["Categoría final (A-E)"] = df.apply(clasificar, axis=1)

        # Mostrar tabla con colores
        def colorear(val):
            if val == "Sí":
                return "background-color: #d1ffd6"
            elif val == "No":
                return "background-color: #ffd1d1"
            elif val == "A":
                return "background-color: #b3ffd1"
            elif val == "E":
                return "background-color: #ffb3b3"
            else:
                return ""

        st.success("✅ Clasificación completada. Aquí está el resultado:")
        st.dataframe(df.style.applymap(colorear), use_container_width=True)

        # Botón de descarga
        def convertir_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Clasificación")
            return output.getvalue()

        st.download_button(
            label="⬇️ Descargar archivo clasificado",
            data=convertir_excel(df),
            file_name="contenido_tesis_clasificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Mostrar imagen de ejemplo
        st.markdown("---")
        st.image("https://i.imgur.com/W0iZBp0.png", caption="Ejemplo de archivo exportado", use_container_width=True)
