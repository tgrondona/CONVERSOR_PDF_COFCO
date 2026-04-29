import streamlit as st
import tempfile
from pathlib import Path

from pdf_a_excel import pdf_a_excel

st.set_page_config(page_title="PDF a Excel", layout="centered")

st.title("📄 → 📊 PDF a Excel")
st.write("Subí un estado de cuenta en PDF y descargá el Excel generado.")

uploaded_file = st.file_uploader("Seleccionar PDF", type=["pdf"])

if uploaded_file is not None:
    if st.button("Convertir"):

        with st.spinner("Procesando PDF..."):
            try:
                # Guardar PDF temporal
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(uploaded_file.read())
                    pdf_path = tmp_pdf.name

                # Archivo de salida
                output_path = pdf_path.replace(".pdf", ".xlsx")

                # Ejecutar tu función
                result = pdf_a_excel(pdf_path, output_path)

                # Leer Excel para descarga
                with open(result, "rb") as f:
                    excel_bytes = f.read()

                st.success("✅ Archivo generado")

                st.download_button(
                    label="📥 Descargar Excel",
                    data=excel_bytes,
                    file_name=Path(result).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error: {e}")
