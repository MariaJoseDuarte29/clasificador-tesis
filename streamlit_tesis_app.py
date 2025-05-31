import streamlit as st
import pandas as pd
import torch
from sentence_transformers import SentenceTransformer, util
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

# Configurar la página con un estilo futurista
st.set_page_config(
    page_title="Clasificador de Capítulos de Tesis",
    page_icon="🧠",
    layout="centered"
)

st.markdown("""
<style>
body {
    background-color: #0d1117;
    color: #c9d1d9;
}

h1, h2, h3 {
    color: #58a6ff;
}

.css-18e3th9 {
    background-color: #161b22;
    border: 1px solid #30363d;
    border-radius: 10px;
    padding: 1rem;
}
</style>
""", unsafe_allow_html=True)

st.title("🧠 Clasificador Inteligente de Capítulos de Tesis")
st.caption("Creado por la arquitecta María José Duarte Torres")

st.markdown("""
Este prototipo de aplicación permite analizar automáticamente el **contenido de los capítulos de una tesis**
para determinar su nivel de **relevancia** y relación con los **objetivos**, el **marco teórico** o la **metodología**.

La app usa un modelo de **IA semántica** para clasificar cada título según:
- Si se relaciona con los **objetivos**
- Si es clave para la **metodología/resultados**
- Si aporta al **marco teórico**
- Si se **repite**
- Si puede **resumirse/eliminarse**

Cada capítulo recibe una categoría (A–E) y se genera un Excel con colores automáticos. 

**Recomendaciones para subir tu archivo Excel:**
- La hoja debe tener una columna llamada exactamente: **"Capítulo o título"** (respetar mayúsculas, minúsculas y acentos).
- No debe contener filas vacías.
""")

st.image("ejemplo_resultado.png", caption="Ejemplo de archivo clasificado exportado por la app", use_container_width=True)

st.markdown("---")

st.header("✍️ Escribe tus textos base")
objetivos_texto = st.text_area("Pega aquí el texto de los objetivos")
metodologia_texto = st.text_area("Pega aquí el texto de la metodología")
marco_texto = st.text_area("Pega aquí el texto del marco teórico")

archivo_excel = st.file_uploader("Sube tu archivo Excel con la columna 'Capítulo o título'", type=["xlsx"])

if st.button("🚀 Ejecutar clasificación") and archivo_excel and objetivos_texto and metodologia_texto and marco_texto:
    try:
        df = pd.read_excel(archivo_excel)

        if "Capítulo o título" not in df.columns:
            st.error("El archivo debe tener una columna llamada exactamente: 'Capítulo o título'")
        else:
            titulos = df["Capítulo o título"].astype(str).tolist()

            modelo = SentenceTransformer("paraphrase-MiniLM-L6-v2")
            embeddings = modelo.encode(titulos + [objetivos_texto, metodologia_texto, marco_texto], convert_to_tensor=True)
            emb_titulos = embeddings[:len(titulos)]
            emb_obj = embeddings[len(titulos)]
            emb_met = embeddings[len(titulos)+1]
            emb_mar = embeddings[len(titulos)+2]

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

            salida = io.BytesIO()
            with pd.ExcelWriter(salida, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Clasificado")
                workbook = writer.book
                ws = workbook.active
                verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

                try:
                    for row in ws.iter_rows(min_row=2, min_col=2, max_col=6):
                        for cell in row:
                            if isinstance(cell.value, str):
                                if cell.value.strip() == "Sí":
                                    cell.fill = verde
                                elif cell.value.strip() == "No":
                                    cell.fill = rojo
                except Exception as e:
                    st.warning(f"No se pudo aplicar colores: {e}")

            st.success("🌟 Clasificación completada")
            st.download_button(
                "📅 Descargar archivo clasificado",
                data=salida.getvalue(),
                file_name="clasificacion_tesis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocurrió un error inesperado: {e}")

# Imagen de marca personal al final
st.markdown("---")
st.image("logo_autora.png", caption="Desarrollado por María José Duarte Torres", use_container_width=True)
 use_container_width=True)

