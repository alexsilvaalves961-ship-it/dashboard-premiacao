import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard de Premiação", layout="wide")

st.title("🏆 Dashboard de Premiação")

# Upload dos arquivos do usuário
st.sidebar.header("Arquivos de Entrada")
file_vendas = st.sidebar.file_uploader("Upload: Vendas", type=["xlsx", "csv"])
file_metas = st.sidebar.file_uploader("Upload: Metas", type=["xlsx", "csv"])

if file_vendas and file_metas:
    # Leitura dos arquivos
    df_vendas = pd.read_excel(file_vendas) if file_vendas.name.endswith('.xlsx') else pd.read_csv(file_vendas)
    df_metas = pd.read_excel(file_metas) if file_metas.name.endswith('.xlsx') else pd.read_csv(file_metas)

    # Processamento dos dados (cole aqui a sua lógica de tratamento)
    st.success("Arquivos carregados com sucesso!")
    
    # Exemplo de exibição de dados
    st.subheader("Resultados das Metas")
    st.dataframe(df_metas)

    st.subheader("Relatório de Vendas")
    st.dataframe(df_vendas)
else:
    st.info("Por favor, faça o upload dos dois arquivos na barra lateral para carregar o painel.")