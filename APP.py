import streamlit as st
import tempfile
from pathlib import Path

from pdf_a_excel_v1 import pdf_a_excel as simple
from pdf_a_excel_v2 import pdf_a_excel as completo

st.set_page_config(page_title="PDF a Excel", layout="centered")

st.title("📄 → 📊 PDF a Excel")

tipo = st.selectbox(
    "Elegir formato",
    [
        "Simple (tabla plana)",
        "Completo (con secciones y resumen)"
    ]
)

uploaded_file = st.file_uploader("Subir PDF", type=["pdf"])

if uploaded_file is not None and st.button("Convertir"):

    with st.spinner("Procesando..."):
        try:
            # guardar pdf temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(uploaded_file.read())
                pdf_path = tmp.name

            output_path = pdf_path.replace(".pdf", ".xlsx")

            # elegir función
            if tipo == "Simple (tabla plana)":
                result = simple(pdf_path, output_path)
            else:
                result = completo(pdf_path, output_path)

            with open(result, "rb") as f:
                data = f.read()

            st.success("✅ Excel generado")

            st.download_button(
                "Descargar Excel",
                data=data,
                file_name=Path(result).name
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
