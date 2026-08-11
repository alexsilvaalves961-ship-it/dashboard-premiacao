import streamlit as st
import pandas as pd
import gdown
import os

st.set_page_config(page_title="Dashboard de Premiação de Motoristas", layout="wide")

st.title("🏆 Dashboard de Premiação de Motoristas")

# Cole seus links públicos do Google Drive aqui
LINK_ABASTECIMENTOS = "https://docs.google.com/spreadsheets/d/1YZBLfxOgJinm1TJHYI49AEaOWVPmOEiA/edit?usp=sharing&ouid=102045408189620250881&rtpof=true&sd=true"
LINK_METAS = "https://docs.google.com/spreadsheets/d/1vX2JqzFLcyDxytrBP5vbHCvMRJsRrm6M/edit?usp=sharing&ouid=102045408189620250881&rtpof=true&sd=true"

@st.cache_data(ttl=600)
def carregar_dados_drive(url, nome_arquivo_local):
    """Baixa a planilha do Google Drive usando gdown e carrega no pandas"""
    # Baixa o arquivo do Google Drive contornando telas de aviso
    gdown.download(url=url, output=nome_arquivo_local, quiet=True, fuzzy=True)
    
    # Lê o arquivo Excel baixado
    df = pd.read_excel(nome_arquivo_local)
    return df

try:
    with st.spinner("Baixando e processando planilhas do Google Drive..."):
        df_abastecimento = carregar_dados_drive(LINK_ABASTECIMENTOS, "abastecimentos.xlsx")
        df_metas = carregar_dados_drive(LINK_METAS, "metas.xlsx")
    
    st.success("Dados carregados com sucesso!")

    aba1, aba2 = st.tabs(["⛽ Dados de Abastecimento", "📊 Metas e Premiação"])
    
    with aba1:
        st.subheader("Base de Abastecimentos")
        st.dataframe(df_abastecimento, use_container_width=True)

    with aba2:
        st.subheader("Relatório de Premiação")
        st.dataframe(df_metas, use_container_width=True)

except Exception as e:
    st.error("Erro ao carregar os arquivos do Google Drive. Verifique se os links estão configurados com acesso público 'Qualquer pessoa com o link'.")
    st.exception(e)
