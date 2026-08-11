import streamlit as st
import pandas as pd

# Configuração da página e layout do Streamlit
st.set_page_config(
    page_title="Dashboard de Premiação de Motoristas",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Links do Google Drive
LINK_ABASTECIMENTOS = "https://docs.google.com/spreadsheets/d/1YZBLfxOgJinm1TJHYI49AEaOWVPmOEiA/edit?usp=sharing&ouid=102045408189620250881&rtpof=true&sd=true"
LINK_METAS = "https://docs.google.com/spreadsheets/d/1vX2JqzFLcyDxytrBP5vbHCvMRJsRrm6M/edit?usp=sharing&ouid=102045408189620250881&rtpof=true&sd=true"

def gerar_url_download_direto(url):
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
    
    df_abast = pd.read_excel(url_abast)
    df_m = pd.read_excel(url_metas)
    
    # Tratamento de colunas se necessário
    if 'CONDUTOR' in df_m.columns:
        df_m['CONDUTOR'] = df_m['CONDUTOR'].astype(str).str.strip()
        
    return df_abast, df_m

try:
    with st.spinner("Carregando inteligência de dados do Google Drive..."):
        df_abastecimento, df_metas = carregar_dados()

    # --- BARRA LATERAL (FILTROS INTERATIVOS) ---
    st.sidebar.image("https://cdn-icons-png.flaticon.com/512/2554/2554936.png", width=100)
    st.sidebar.title("📌 Filtros do Painel")
    
    # Filtro por Condutor
    lista_condutores = ["Todos"] + sorted(list(df_metas["CONDUTOR"].dropna().unique())) if "CONDUTOR" in df_metas.columns else ["Todos"]
    condutor_selecionado = st.sidebar.selectbox("Filtrar por Motorista:", lista_condutores)

    # Filtro por Categoria
    if "CATEGORIAS" in df_metas.columns:
        lista_categorias = ["Todas"] + sorted(list(df_metas["CATEGORIAS"].dropna().unique()))
        categoria_selecionada = st.sidebar.selectbox("Filtrar por Categoria:", lista_categorias)
    else:
        categoria_selecionada = "Todas"

    # Aplicação dos Filtros
    df_metas_filtrado = df_metas.copy()
    
    if condutor_selecionado != "Todos":
        df_metas_filtrado = df_metas_filtrado[df_metas_filtrado["CONDUTOR"] == condutor_selecionado]
        
    if categoria_selecionada != "Todas" and "CATEGORIAS" in df_metas_filtrado.columns:
        df_metas_filtrado = df_metas_filtrado[df_metas_filtrado["CATEGORIAS"] == categoria_selecionada]

    # --- CABEÇALHO DO DASHBOARD ---
    st.title("🏆 Dashboard de Premiação de Motoristas")
    st.markdown("---")

    # --- CARTOES DE MÉTRICAS (KPIs) ---
    col1, col2, col3, col4 = st.columns(4)
    
    total_motoristas = len(df_metas_filtrado["CONDUTOR"].unique()) if "CONDUTOR" in df_metas_filtrado.columns else 0
    total_km = df_metas_filtrado["KM_TOTAL"].sum() if "KM_TOTAL" in df_metas_filtrado.columns else 0
    total_litros = df_metas_filtrado["LITROS_TOTAL"].sum() if "LITROS_TOTAL" in df_metas_filtrado.columns else 0
    total_premiacao = df_metas_filtrado["VALOR_TOTAL_PAGAR"].sum() if "VALOR_TOTAL_PAGAR" in df_metas_filtrado.columns else 0

    col1.metric("👥 Motoristas", f"{total_motoristas}")
    col2.metric("🛣️ KM Percorridos", f"{total_km:,.0f}".replace(",", "."))
    col3.metric("⛽ Consumo Total (L)", f"{total_litros:,.2f}".replace(",", "."))
    col4.metric("💰 Total em Premiações", f"R$ {total_premiacao:,.2f}".replace(",", "."))

    st.markdown("---")

    # --- NAVEGAÇÃO POR ABAS ---
    aba1, aba2, aba3 = st.tabs(["📊 Relatório de Premiação", "🥇 Ranking de Desempenho", "⛽ Base de Abastecimentos"])

    with aba1:
        st.subheader("📋 Detalhamento das Premiações")
        st.dataframe(
            df_metas_filtrado, 
            use_container_width=True,
            height=400
        )

    with aba2:
        st.subheader("🏆 Top Motoristas por Valor a Receber")
        if "VALOR_TOTAL_PAGAR" in df_metas_filtrado.columns and "CONDUTOR" in df_metas_filtrado.columns:
            top_motoristas = df_metas_filtrado.sort_values(by="VALOR_TOTAL_PAGAR", ascending=False)[["CONDUTOR", "CATEGORIAS", "KM_TOTAL", "VALOR_TOTAL_PAGAR"]]
            
            # Gráfico de Barras no Streamlit
            st.bar_chart(
                data=top_motoristas.set_index("CONDUTOR")["VALOR_TOTAL_PAGAR"],
                color="#28a745"
            )
            
            st.dataframe(top_motoristas, use_container_width=True)

    with aba3:
        st.subheader("⛽ Histórico Completo de Abastecimentos")
        st.dataframe(df_abastecimento, use_container_width=True, height=400)

except Exception as e:
    st.error("Erro ao carregar o dashboard interativo. Verifique se as colunas da planilha correspondem às métricas.")
    st.exception(e)
