import streamlit as st
import pandas as pd
import io
import base64
from sentence_transformers import SentenceTransformer, util
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
import xlsxwriter

st.set_page_config(page_title="Clasificador de Tesis", layout="wide")

st.title("🧠 Clasificador de Capítulos de Tesis")
st.subheader("Creado por la arquitecta María José Duarte Torres")

st.markdown("""
Esta aplicación permite analizar automáticamente la relevancia de capítulos o títulos dentro de una tesis en función de su afinidad con los **objetivos**, **metodología** y **marco teórico**. 
El sistema clasifica el contenido en categorías (A-E) según su nivel de pertinencia y permite descargar una tabla coloreada con los resultados.

👉 **Recomendaciones para cargar los archivos:**
- La tabla Excel debe tener una columna llamada exactamente: `Capítulo o título`
- Sube los textos de **objetivos**, **metodología** y **marco teórico** en formato `.txt`
""")

# Archivos
excel_file = st.file_uploader("📂 Sube el archivo Excel con los títulos", type=["xlsx"])
objetivos_txt = st.file_uploader("🎯 Sube el archivo de Objetivos (.txt)", type=["txt"])
metodologia_txt = st.file_uploader("🔧 Sube el archivo de Metodología (.txt)", type=["txt"])
marco_teorico_txt = st.file_uploader("📚 Sube el archivo de Marco Teórico (.txt)", type=["txt"])

# Botón de acción
if st.button("📊 Generar tabla de clasificación"):

    if not all([excel_file, objetivos_txt, metodologia_txt, marco_teorico_txt]):
        st.error("🚨 Por favor, sube todos los archivos antes de continuar.")
    else:
        # Modelo
        model = SentenceTransformer("all-MiniLM-L6-v2")

        # Cargar textos
        objetivos = objetivos_txt.read().decode("utf-8")
        metodologia = metodologia_txt.read().decode("utf-8")
        marco_teorico = marco_teorico_txt.read().decode("utf-8")

        # Leer Excel
        df = pd.read_excel(excel_file)
        if "Capítulo o título" not in df.columns:
            st.error("❌ La tabla debe contener la columna 'Capítulo o título'.")
        else:
            titulos = df["Capítulo o título"].astype(str).tolist()

            # Embeddings
            emb_titulos = model.encode(titulos, convert_to_tensor=True)
            emb_objetivos = model.encode(objetivos, convert_to_tensor=True)
            emb_metodologia = model.encode(metodologia, convert_to_tensor=True)
            emb_marco = model.encode(marco_teorico, convert_to_tensor=True)

            # Calcular similitudes
            similitudes_obj = [float(util.cos_sim(t, emb_objetivos)) for t in emb_titulos]
            similitudes_met = [float(util.cos_sim(t, emb_metodologia)) for t in emb_titulos]
            similitudes_marco = [float(util.cos_sim(t, emb_marco)) for t in emb_titulos]

            # Clasificación
            df["¿Relaciona un objetivo?"] = ["Sí" if s > 0.45 else "No" for s in similitudes_obj]
            df["¿Es clave para entender metodología/resultados?"] = ["Sí" if s > 0.45 else "No" for s in similitudes_met]
            df["¿Aporta al marco teórico?"] = ["Sí" if s > 0.45 else "No" for s in similitudes_marco]

            # Repeticiones
            df["¿Se repite en otro capítulo?"] = df.duplicated(subset=["Capítulo o título"], keep=False).map({True: "Sí", False: "No"})

            # Sugerencia si eliminar
            df["¿Puede resumirse/eliminarse?"] = ["Sí" if (a == "No" and b == "No" and c == "No") else "No"
                                                  for a, b, c in zip(df["¿Relaciona un objetivo?"],
                                                                     df["¿Es clave para entender metodología/resultados?"],
                                                                     df["¿Aporta al marco teórico?"])]

            # Categorización
            def categorizar(row):
                puntos = sum([row["¿Relaciona un objetivo?"] == "Sí",
                              row["¿Es clave para entender metodología/resultados?"] == "Sí",
                              row["¿Aporta al marco teórico?"] == "Sí"])
                if puntos == 3:
                    return "A"
                elif puntos == 2:
                    return "B"
                elif puntos == 1:
                    return "C"
                elif row["¿Puede resumirse/eliminarse?"] == "Sí":
                    return "E"
                else:
                    return "D"

            df["Categoría final (A-E)"] = df.apply(categorizar, axis=1)

            # Guardar Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Clasificación")
                workbook = writer.book
                worksheet = writer.sheets["Clasificación"]

                # Colores para Sí/No
                color_si = workbook.add_format({"bg_color": "#D4EFDF"})
                color_no = workbook.add_format({"bg_color": "#F5B7B1"})

                # Colorear celdas
                columnas_si_no = [
                    "¿Relaciona un objetivo?",
                    "¿Es clave para entender metodología/resultados?",
                    "¿Aporta al marco teórico?",
                    "¿Se repite en otro capítulo?",
                    "¿Puede resumirse/eliminarse?"
                ]

                for col_idx, col in enumerate(df.columns):
                    if col in columnas_si_no:
                        for row_idx, valor in enumerate(df[col]):
                            formato = color_si if valor == "Sí" else color_no
                            worksheet.write(row_idx + 1, col_idx, valor, formato)

            # Descargar
            st.success("✅ Clasificación completada. Descarga disponible:")
            st.download_button("📥 Descargar archivo Excel", data=output.getvalue(),
                               file_name="contenido_clasificado.xlsx", mime="application/vnd.ms-excel")

# Imagen decorativa
st.markdown("---")
st.image("https://i.imgur.com/0uX4A0z.png", caption="Ejemplo del resultado exportado", use_container_width=True)
