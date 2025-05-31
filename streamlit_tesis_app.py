# streamlit_tesis_app.py
import streamlit as st
import pandas as pd
from sentence_transformers import SentenceTransformer, util
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os

# Inicializar modelo
@st.cache_resource
def cargar_modelo():
    return SentenceTransformer("all-MiniLM-L6-v2")

model = cargar_modelo()

# Función para aplicar colores en Excel
def aplicar_colores(nombre_archivo):
    wb = load_workbook(nombre_archivo)
    ws = wb.active

    color_si = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    color_no = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    categorias_color = {
        "A": PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
        "B": PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid"),
        "C": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
        "D": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
        "E": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    }

    headers = {cell.value: idx for idx, cell in enumerate(ws[1])}
    columnas_si_no = [
        "¿Relaciona un objetivo?",
        "¿Es clave para entender metodología/resultados?",
        "¿Aporta al marco teórico?",
        "¿Se repite en otro capítulo?",
        "¿Puede resumirse/eliminarse?"
    ]

    for fila in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for col in columnas_si_no:
            cell = fila[headers[col]]
            cell.fill = color_si if cell.value == "Sí" else color_no
        cat = fila[headers["Categoría final (A-E)"]].value
        if cat in categorias_color:
            fila[headers["Categoría final (A-E)"]].fill = categorias_color[cat]

    wb.save(nombre_archivo)

# Interfaz Streamlit
st.title("Clasificación Automática de Capítulos de Tesis")
st.markdown("Sube tus archivos para clasificar los títulos automáticamente.")

archivo_excel = st.file_uploader("📄 Sube el archivo .xlsx con los títulos", type="xlsx")
objetivos = st.text_area("🎯 Pega el texto de los objetivos")
metodologia = st.text_area("⚙️ Pega el texto de la metodología")
marco = st.text_area("📚 Pega el texto del marco teórico")

if st.button("Ejecutar Clasificación"):
    if archivo_excel and objetivos and metodologia and marco:
        df = pd.read_excel(archivo_excel)
        titulos = df["Título"].astype(str).tolist()

        emb_titulos = model.encode(titulos, convert_to_tensor=True).cpu()
        emb_obj = model.encode(objetivos, convert_to_tensor=True).cpu()
        emb_met = model.encode(metodologia, convert_to_tensor=True).cpu()
        emb_marco = model.encode(marco, convert_to_tensor=True).cpu()

        resultados = []
        for i, titulo in enumerate(titulos):
            e = emb_titulos[i]
            rel_obj = util.cos_sim(e, emb_obj).item() > 0.35
            rel_met = util.cos_sim(e, emb_met).item() > 0.35
            rel_marco = util.cos_sim(e, emb_marco).item() > 0.35
            repetido = any(i != j and util.cos_sim(e, emb).item() > 0.65 for j, emb in enumerate(emb_titulos))
            puede_resumirse = not (rel_obj or rel_met or rel_marco or repetido)

            total_si = sum([rel_obj, rel_met, rel_marco])
            cat = "A" if total_si == 3 else "B" if total_si == 2 else "C" if total_si == 1 else "D" if not puede_resumirse else "E"

            resultados.append({
                "Título": titulo,
                "¿Relaciona un objetivo?": "Sí" if rel_obj else "No",
                "¿Es clave para entender metodología/resultados?": "Sí" if rel_met else "No",
                "¿Aporta al marco teórico?": "Sí" if rel_marco else "No",
                "¿Se repite en otro capítulo?": "Sí" if repetido else "No",
                "¿Puede resumirse/eliminarse?": "Sí" if puede_resumirse else "No",
                "Categoría final (A-E)": cat
            })

        df_resultado = pd.DataFrame(resultados)
        nombre_salida = "contenido_tesis_clasificado_coloreado.xlsx"
        df_resultado.to_excel(nombre_salida, index=False)
        aplicar_colores(nombre_salida)

        with open(nombre_salida, "rb") as f:
            st.download_button("📥 Descargar archivo clasificado y coloreado", f, file_name=nombre_salida)

    else:
        st.error("Por favor, sube el archivo y pega los tres textos para continuar.")
