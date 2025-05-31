import streamlit as st
import pandas as pd
import torch
from sentence_transformers import SentenceTransformer, util
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Configurar página
st.set_page_config(
    page_title="Clasificador de Capítulos de Tesis",
    page_icon="📘",
    layout="centered"
)

# Cabecera visual
st.title("📘 Clasificador Inteligente de Capítulos de Tesis")
st.caption("Creado por la arquitecta María José Duarte Torres")

st.markdown("""
Este prototipo de aplicación permite analizar de forma automatizada el **contenido de los capítulos de una tesis**
para determinar su nivel de **relevancia** y relación con los **objetivos**, el **marco teórico** o la **metodología** del proyecto investigativo.

El sistema aplica un modelo de **inteligencia artificial basado en similitud semántica** para generar etiquetas como:

- ✅ Si el contenido está relacionado con los objetivos del proyecto
- ✅ Si aporta directamente a la metodología o resultados
- ✅ Si se relaciona con el marco teórico
- ⚠️ Si se repite en otro capítulo
- 🗑️ Si puede eliminarse o resumirse

Finalmente, el sistema asigna una **categoría de relevancia (A–E)** a cada capítulo, y genera una hoja de Excel de salida con colores y sugerencias automáticas.
""")

# Imagen de ejemplo
st.image("ejemplo_resultado.png", caption="Ejemplo de archivo clasificado exportado por la app", use_column_width=True)

st.markdown("---")

# Cargar archivos
st.header("📂 Subir archivos")

archivo_excel = st.file_uploader("Sube tu archivo de capítulos (.xlsx)", type=["xlsx"])
objetivos = st.file_uploader("Sube el archivo de texto con los objetivos", type=["txt"])
metodologia = st.file_uploader("Sube el archivo de texto con la metodología", type=["txt"])
marco = st.file_uploader("Sube el archivo de texto con el marco teórico", type=["txt"])

if archivo_excel and objetivos and metodologia and marco:
    st.success("✅ Archivos cargados correctamente")

    # Leer textos base
    objetivos_texto = objetivos.read().decode("utf-8")
    metodologia_texto = metodologia.read().decode("utf-8")
    marco_texto = marco.read().decode("utf-8")

    # Leer Excel
    df = pd.read_excel(archivo_excel)
    titulos = df["Capítulo o título"].astype(str).tolist()

    # Cargar modelo
    modelo = SentenceTransformer("paraphrase-MiniLM-L6-v2")

    # Obtener embeddings
    embeddings = modelo.encode(titulos + [objetivos_texto, metodologia_texto, marco_texto], convert_to_tensor=True)
    emb_titulos = embeddings[:len(titulos)]
    emb_obj = embeddings[len(titulos)]
    emb_met = embeddings[len(titulos)+1]
    emb_mar = embeddings[len(titulos)+2]

    # Clasificación
    data = []
    for i, titulo in enumerate(titulos):
        sim_obj = float(util.cos_sim(emb_titulos[i], emb_obj))
        sim_met = float(util.cos_sim(emb_titulos[i], emb_met))
        sim_mar = float(util.cos_sim(emb_titulos[i], emb_mar))

        relacion_objetivo = "Sí" if sim_obj > 0.45 else "No"
        clave_metodologia = "Sí" if sim_met > 0.45 else "No"
        aporta_marco = "Sí" if sim_mar > 0.45 else "No"
        se_repite = "Sí" if titulo in titulos[:i] + titulos[i+1:] else "No"
        puede_eliminarse = "Sí" if (sim_obj < 0.3 and sim_met < 0.3 and sim_mar < 0.3) else "No"

        # Categoría A–E según cantidad de "Sí"
        puntuacion = sum([relacion_objetivo, clave_metodologia, aporta_marco].count("Sí"))
        categoria = "A" if puntuacion == 3 else "B" if puntuacion == 2 else "C" if puntuacion == 1 else "D" if se_repite == "Sí" else "E"

        data.append([
            categoria, relacion_objetivo, clave_metodologia,
            aporta_marco, se_repite, puede_eliminarse
        ])

    columnas = [
        "Categoría final (A–E)",
        "¿Relaciona un objetivo?",
        "¿Es clave para entender metodología/resultados?",
        "¿Aporta al marco teórico?",
        "¿Se repite en otro capítulo?",
        "¿Puede resumirse/eliminarse?"
    ]

    df_resultado = pd.DataFrame(data, columns=columnas)
    df_final = pd.concat([df["Capítulo o título"], df_resultado], axis=1)

    # Guardar resultado en Excel
    archivo_salida = "resultado_clasificado.xlsx"
    df_final.to_excel(archivo_salida, index=False)

    # Aplicar colores
    wb = load_workbook(archivo_salida)
    ws = wb.active
    color_si = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Verde
    color_no = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Rojo

    for row in ws.iter_rows(min_row=2, min_col=2, max_col=6):
        for cell in row:
            if cell.value == "Sí":
                cell.fill = color_si
            elif cell.value == "No":
                cell.fill = color_no

    wb.save(archivo_salida)

    # Descargar
    with open(archivo_salida, "rb") as f:
        st.download_button(
            label="📥 Descargar archivo clasificado",
            data=f,
            file_name="clasificacion_tesis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
