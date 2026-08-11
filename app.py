import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard de Premiação de Motoristas", layout="wide")

st.title("🏆 Dashboard de Premiação de Motoristas")

# Links compartilhados das planilhas no Google Drive
LINK_ABASTECIMENTOS = "https://docs.google.com/spreadsheets/d/1YZBLfxOgJinm1TJHYI49AEaOWVPmOEiA/edit?usp=sharing&ouid=102045408189620250881&rtpof=true&sd=true"
LINK_METAS = "https://docs.google.com/spreadsheets/d/1vX2JqzFLcyDxytrBP5vbHCvMRJsRrm6M/edit?usp=sharing&ouid=102045408189620250881&rtpof=true&sd=true"

def gerar_url_download_direto(url):
    """Extrai o ID da planilha e gera o link de download direto em formato Excel"""
    file_id = ""
    if "/d/" in url:
        file_id = url.split("/d/")[1].split("/")[0]
    elif "id=" in url:
        file_id = url.split("id=")[1].split("&")[0]
    else:
        file_id = url
    return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"

@st.cache_data(ttl=600)
def carregar_dados():
    url_abast = gerar_url_download_direto(LINK_ABASTECIMENTOS)
    url_metas = gerar_url_download_direto(LINK_METAS)
    
    # Lê as planilhas diretamente da URL de exportação do Google Sheets
    df_abast = pd.read_excel(url_abast)
    df_m = pd.read_excel(url_metas)
    
    return df_abast, df_m

try:
    with st.spinner("Carregando planilhas do Google Drive..."):
        df_abastecimento, df_metas = carregar_dados()
    
    st.success("Dados carregados com sucesso!")

    aba1, aba2 = st.tabs(["⛽ Dados de Abastecimento", "📊 Metas e Premiação"])
    
    with aba1:
        st.subheader("Base de Abastecimentos")
        st.dataframe(df_abastecimento, use_container_width=True)

    with aba2:
        st.subheader("Relatório de Premiação")
        st.dataframe(df_metas, use_container_width=True)

except Exception as e:
    st.error("Erro ao carregar os arquivos do Google Drive. Verifique se os links estão com acesso 'Qualquer pessoa com o link'.")
    st.exception(e)
