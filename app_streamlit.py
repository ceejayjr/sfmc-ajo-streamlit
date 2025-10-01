import streamlit as st
from replacer import process

st.set_page_config(page_title="SFMC → AJO Converter", layout="centered")
st.title("SFMC → AJO Converter")
st.write("Envie um **HTML com AMPScript** e a planilha **SFMCtoAJOComparision.xlsx** (A: SFMC, B: AJO).")

html_file = st.file_uploader("HTML de entrada", type=["html","htm","txt"])
excel_file = st.file_uploader("Planilha de-para (SFMCtoAJOComparision.xlsx)", type=["xlsx"])

if st.button("Processar"):
    if not html_file or not excel_file:
        st.error("Envie o HTML e a planilha Excel.")
    else:
        output_bytes, report = process(html_file.getvalue(), excel_file.getvalue())
        st.success(f"Substituições aplicadas: {report['replaced']} • AMPScript comentado: {report['commented']}")
        if report["details"]:
            with st.expander("Ver primeiras substituições"):
                for snip, n in report["details"]:
                    st.write(f"- [{n}x] {snip}")
        st.download_button(
            "Baixar HTML Processado",
            data=output_bytes,
            file_name="output_processed.html",
            mime="text/html"
        )

st.caption("Processamento 100% em memória. Nenhum arquivo é salvo no servidor.")
