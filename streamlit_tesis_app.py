import streamlit as st
import pandas as pd
from sentence_transformers import SentenceTransformer, util
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Cargar modelo de embeddings
modelo = SentenceTransformer('all-MiniLM-L6-v2')

st.set_page_config(page_title="Clasificador de capítulos", layout="centered")

st.markdown("## Clasificador de capítulos de tesis")
st.markdown("**Creado por la arquitecta María José Duarte Torres**")
st.markdown("""
Esta herramienta permite analizar los títulos o capítulos de una tesis y determinar si están relacionados con los **objetivos**, la **metodología** y el **marco teórico**.
Los resultados se clasifican con colores y una etiqueta final (de A a E), para ayudarte a decidir qué conservar, resumir o eliminar del documento.
""")

# 📥 Subida de archivos
archivo_excel = st.file_uploader("📎 Sube el archivo Excel con los capítulos", type=["xlsx"])

objetivos = st.text_area("✏️ Pega aquí el texto de los objetivos")
metodologia = st.text_area("✏️ Pega aquí el texto de la metodología")
marco_teorico = st.text_area("✏️ Pega aquí el texto del marco teórico")

if st.button("📊 Analizar y clasificar"):
    if archivo_excel and objetivos and metodologia and marco_teorico:
        df = pd.read_excel(archivo_excel)

        # Verificar nombre de columna
        if "Capítulo o título" not in df.columns:
            st.error("❗ La tabla debe tener una columna llamada 'Capítulo o título'")
        else:
            titulos = df["Capítulo o título"].astype(str).tolist()

            # Embeddings
            emb_titulos = modelo.encode(titulos, convert_to_tensor=True)
            emb_obj = modelo.encode(objetivos, convert_to_tensor=True)
            emb_met = modelo.encode(metodologia, convert_to_tensor=True)
            emb_marco = modelo.encode(marco_teorico, convert_to_tensor=True)

            # Comparaciones
            def comparar(texto_emb):
                return [float(util.cos_sim(t, texto_emb)) for t in emb_titulos]

            sims_obj = comparar(emb_obj)
            sims_met = comparar(emb_met)
            sims_marco = comparar(emb_marco)

            # Detectar repeticiones
            repetidos = df["Capítulo o título"].duplicated(keep=False)

            # Decisiones
            df["¿Relaciona un objetivo?"] = ["Sí" if s > 0.3 else "No" for s in sims_obj]
            df["¿Es clave para entender metodología/resultados?"] = ["Sí" if s > 0.3 else "No" for s in sims_met]
            df["¿Aporta al marco teórico?"] = ["Sí" if s > 0.3 else "No" for s in sims_marco]
            df["¿Se repite en otro capítulo?"] = ["Sí" if r else "No" for r in repetidos]

            # Clasificación final
            def categorizar(fila):
                puntaje = sum([
                    fila["¿Relaciona un objetivo?"] == "Sí",
                    fila["¿Es clave para entender metodología/resultados?"] == "Sí",
                    fila["¿Aporta al marco teórico?"] == "Sí"
                ])
                if puntaje == 3:
                    return "A"
                elif puntaje == 2:
                    return "B"
                elif puntaje == 1:
                    return "C"
                elif fila["¿Se repite en otro capítulo?"] == "Sí":
                    return "D"
                else:
                    return "E"

            df["Categoría final (A-E)"] = df.apply(categorizar, axis=1)

            # Guardar con color
            salida = BytesIO()
            df.to_excel(salida, index=False)
            salida.seek(0)

            wb = load_workbook(salida)
            ws = wb.active

            color_si = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Verde
            color_no = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Rojo

            for row in ws.iter_rows(min_row=2, min_col=2, max_col=6):
                for cell in row:
                    if cell.value == "Sí":
                        cell.fill = color_si
                    elif cell.value == "No":
                        cell.fill = color_no

            # Descargar
            final_salida = BytesIO()
            wb.save(final_salida)
            final_salida.seek(0)

            st.success("✅ Clasificación completada con éxito.")
            st.download_button("📥 Descargar resultados", data=final_salida,
                               file_name="contenido_tesis_clasificado.xlsx", mime="application/vnd.ms-excel")

            st.image("logo.png", use_container_width=True)
    else:
        st.warning("🔍 Por favor, sube el Excel y escribe los tres textos.")
