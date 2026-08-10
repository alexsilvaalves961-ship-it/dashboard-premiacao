import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard de Premiação de Motoristas", layout="wide")

st.title("🏆 Dashboard de Premiação de Motoristas")

# Insira aqui os links compartilhados do Google Drive
LINK_ABASTECIMENTOS = "COLE_AQUI_O_LINK_DA_PLANILHA_DE_ABASTECIMENTOS"
LINK_METAS = "COLE_AQUI_O_LINK_DA_PLANILHA_DE_METAS"

def converter_link_drive(url):
    """Converte o link de visualização do Google Drive em link de download direto"""
    if "file/d/" in url:
        file_id = url.split("file/d/")[1].split("/")[0]
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    return url

@st.cache_data(ttl=600)  # Atualiza os dados a cada 10 minutos
def carregar_dados():
    url_abastecimento = converter_link_drive(LINK_ABASTECIMENTOS)
    url_metas = converter_link_drive(LINK_METAS)
    
    df_abastecimento = pd.read_excel(url_abastecimento)
    df_metas = pd.read_excel(url_metas)
    
    return df_abastecimento, df_metas

try:
    with st.spinner("Carregando planilhas diretamente do Google Drive..."):
        df_abastecimento, df_metas = carregar_dados()
    
    st.success("Dados carregados com sucesso!")

    # Exibição das abas com os dados
    aba1, aba2 = st.tabs(["⛽ Dados de Abastecimento / Operação", "📊 Metas e Premiação"])
    
    with aba1:
        st.subheader("Base de Abastecimentos")
        st.dataframe(df_abastecimento, use_container_width=True)

    with aba2:
        st.subheader("Relatório de Premiação")
        st.dataframe(df_metas, use_container_width=True)

except Exception as e:
    st.error("Erro ao carregar os arquivos do Google Drive. Verifique se os links estão com acesso público 'Qualquer pessoa com o link'.")
    st.exception(e)
