import io
import re
import unicodedata
from datetime import datetime
from typing import Tuple

import numpy as np
import pandas as pd
import requests
import streamlit as st

# ================================================================
# CONFIGURAÇÕES DA APLICAÇÃO E LINKS DO GOOGLE DRIVE
# ================================================================
class AppConfig:
    LINK_PASTA4 = "https://docs.google.com/spreadsheets/d/1NXtnCAG5DtgDXegtiuefD6oI7RmLbi28/edit?usp=drive_link"
    LINK_PASTA2 = "https://docs.google.com/spreadsheets/d/1b8F074yKCPX-iZpuB7oW3s0D2oWIGYxX/edit?usp=drive_link"
    LINK_FROTA = "https://docs.google.com/spreadsheets/d/1_X6Li8z-Fkfv3U6_D-8ByG0AP5c8dZeT/edit?usp=drive_link"
    LINK_MOTORISTAS = "https://docs.google.com/spreadsheets/d/1vX2JqzFLcyDxytrBP5vbHCvMRJsRrm6M/edit?usp=drive_link"
    LINK_ABASTECIMENTOS = "https://docs.google.com/spreadsheets/d/1YZBLfxOgJinm1TJHYI49AEaOWVPmOEiA/edit?usp=drive_link"

    @staticmethod
    def extrair_url_export(url: str) -> str:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
        if match:
            doc_id = match.group(1)
            return f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
        return url


# ================================================================
# FUNÇÕES UTILITÁRIAS
# ================================================================
class DataUtils:
    @staticmethod
    def normalizar_texto(texto) -> str:
        if pd.isna(texto):
            return ""
        texto = str(texto).upper().strip()
        texto = unicodedata.normalize("NFD", texto)
        texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
        return re.sub(r"\s+", " ", texto)

    @staticmethod
    def padronizar_placa(placa) -> str:
        if pd.isna(placa):
            return ""
        p = re.sub(r"[^A-Z0-9]", "", str(placa).upper().strip())
        return p if len(p) == 7 else ""

    @staticmethod
    def placa_equivalente_mercosul(placa: str) -> str:
        mapa_letras = {"0": "A", "1": "B", "2": "C", "3": "D", "4": "E", "5": "F", "6": "G", "7": "H", "8": "I", "9": "J"}
        mapa_nums = {v: k for k, v in mapa_letras.items()}

        if len(placa) != 7:
            return ""

        if placa[4].isdigit():
            return placa[:4] + mapa_letras.get(placa[4], placa[4]) + placa[5:]
        else:
            return placa[:4] + mapa_nums.get(placa[4], placa[4]) + placa[5:]

    @staticmethod
    def converter_numero(valor) -> float:
        if pd.isna(valor):
            return np.nan
        val_str = str(valor).replace("R$", "").replace(" ", "").strip()
        if "," in val_str and "." in val_str:
            val_str = val_str.replace(".", "").replace(",", ".")
        elif "," in val_str:
            val_str = val_str.replace(",", ".")
        try:
            return float(val_str)
        except ValueError:
            return np.nan

    @staticmethod
    def encontrar_coluna(df: pd.DataFrame, opcoes_nomes: list):
        colunas_normalizadas = {col: DataUtils.normalizar_texto(col) for col in df.columns}
        for opcao in opcoes_nomes:
            opcao_norm = DataUtils.normalizar_texto(opcao)
            for col_orig, col_norm in colunas_normalizadas.items():
                if opcao_norm in col_norm:
                    return col_orig
        return None

    @staticmethod
    def gerar_excel_download(df_dict: dict) -> bytes:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for aba_nome, dataframe in df_dict.items():
                dataframe.to_excel(writer, sheet_name=aba_nome[:31], index=False)
        return output.getvalue()


# ================================================================
# LEITURA E PROCESSAMENTO DAS PLANILHAS
# ================================================================
class DataLoader:
    def __init__(self, config: AppConfig):
        self.config = config

    def _baixar_excel(self, url: str, sheet_name=0, header=0) -> pd.DataFrame:
        url_export = self.config.extrair_url_export(url)
        resp = requests.get(url_export)
        resp.raise_for_status()
        return pd.read_excel(io.BytesIO(resp.content), sheet_name=sheet_name, header=header)

    def _processar_tabela_precos(self, df: pd.DataFrame) -> pd.DataFrame:
        col_tipo = DataUtils.encontrar_coluna(df, ["TIPO", "TIPO VEICULO", "TIPO DE VEICULO", "CATEGORIA"])
        col_media = DataUtils.encontrar_coluna(df, ["MEDIA", "MÉDIA", "MEDIA OBJETIVO"])
        col_premio = DataUtils.encontrar_coluna(df, ["TOTAL", "PREMIO", "PRÊMIO", "VALOR", "VALOR PREMIO"])

        if None in (col_tipo, col_media, col_premio) and df.shape[1] >= 3:
            df_temp = df.dropna(how="all").copy()
            if len(df_temp.columns) >= 5:
                df_temp = df_temp.iloc[1:].copy()
                df_temp.columns = ["DEL", "TIPO", "FATOR", "MEDIA", "PREMIO"] + list(df_temp.columns[5:])
                col_tipo, col_media, col_premio = "TIPO", "MEDIA", "PREMIO"
                df = df_temp

        if col_tipo is None or col_tipo not in df.columns:
            return pd.DataFrame(columns=["TIPO", "MEDIA", "PREMIO"])

        resultado = pd.DataFrame({
            "TIPO": df[col_tipo].apply(DataUtils.normalizar_texto),
            "MEDIA": df[col_media].apply(DataUtils.converter_numero) if col_media else np.nan,
            "PREMIO": df[col_premio].apply(DataUtils.converter_numero) if col_premio else 0.0,
        }).dropna(subset=["MEDIA", "PREMIO"])

        resultado["TIPO"] = resultado["TIPO"].replace({"TOCO": "TRUCK"})
        return resultado[resultado["TIPO"] != ""]

    def carregar_precos(self) -> pd.DataFrame:
        df_p4 = self._baixar_excel(self.config.LINK_PASTA4, sheet_name=0)
        df_p2 = self._baixar_excel(self.config.LINK_PASTA2, sheet_name=0)

        res_p4 = self._processar_tabela_precos(df_p4)
        res_p2 = self._processar_tabela_precos(df_p2)

        combinado = pd.concat([res_p4, res_p2], ignore_index=True)
        combinado = combinado.drop_duplicates(subset=["TIPO"], keep="last")

        if combinado.empty:
            raise ValueError("Não foi possível identificar colunas válidas em Pasta2 ou Pasta4.")

        return combinado.reset_index(drop=True)

    def carregar_frota(self) -> Tuple[pd.DataFrame, dict]:
        df = self._baixar_excel(self.config.LINK_FROTA)
        col_placa = DataUtils.encontrar_coluna(df, ["CAVALO", "PLACA", "PLACA CAVALO", "PLACA DO CAVALO"])
        col_tipo = DataUtils.encontrar_coluna(df, ["TIPO", "TIPO VEICULO", "TIPO DE VEICULO", "CATEGORIA"])

        resultado = pd.DataFrame({
            "PLACA_PADRONIZADA": df[col_placa].apply(DataUtils.padronizar_placa),
            "TIPO": df[col_tipo].apply(DataUtils.normalizar_texto),
        })
        resultado = resultado[resultado["PLACA_PADRONIZADA"] != ""].copy()
        resultado["TIPO"] = resultado["TIPO"].replace({"TOCO": "TRUCK"})
        resultado = resultado.drop_duplicates("PLACA_PADRONIZADA", keep="last")

        mapa = dict(zip(resultado["PLACA_PADRONIZADA"], resultado["TIPO"]))
        return resultado, mapa

    def carregar_cadastro_motoristas(self) -> pd.DataFrame:
        bruto = self._baixar_excel(self.config.LINK_MOTORISTAS, header=None)
        
        cab_idx = None
        for i in range(min(len(bruto), 20)):
            vals = [DataUtils.normalizar_texto(x) for x in bruto.iloc[i].tolist()]
            linha_texto = " ".join(vals)
            tem_motorista = any(m in linha_texto for m in ["MOTORISTA", "CONDUTOR", "NOME"])
            tem_tipo = any(t in linha_texto for t in ["TIPO", "CATEGORIA", "VEICULO"])
            
            if tem_motorista and tem_tipo:
                cab_idx = i
                break

        if cab_idx is not None:
            df = self._baixar_excel(self.config.LINK_MOTORISTAS, header=cab_idx)
        else:
            df = self._baixar_excel(self.config.LINK_MOTORISTAS, header=0)

        col_mot = DataUtils.encontrar_coluna(df, ["MOTORISTA", "MOTORISTAS", "CONDUTOR", "NOME"])
        col_tipo = DataUtils.encontrar_coluna(df, ["TIPO", "CATEGORIA", "TIPO VEICULO", "TIPO DE VEICULO"])
        col_base = DataUtils.encontrar_coluna(df, ["BASE", "FILIAL", "UNIDADE"])

        if not col_mot or not col_tipo:
            cadastro = pd.DataFrame({
                "MOTORISTA_CADASTRO": df.iloc[:, 0].apply(DataUtils.normalizar_texto),
                "TIPO_CADASTRO": df.iloc[:, 1].apply(DataUtils.normalizar_texto).replace({"TOCO": "TRUCK"}),
                "BASE_CADASTRO": df.iloc[:, 2].apply(DataUtils.normalizar_texto) if df.shape[1] > 2 else "",
            })
        else:
            cadastro = pd.DataFrame({
                "MOTORISTA_CADASTRO": df[col_mot].apply(DataUtils.normalizar_texto),
                "TIPO_CADASTRO": df[col_tipo].apply(DataUtils.normalizar_texto).replace({"TOCO": "TRUCK"}),
                "BASE_CADASTRO": df[col_base].apply(DataUtils.normalizar_texto) if col_base else "",
            })

        cadastro["EH_FOLGUISTA"] = cadastro["TIPO_CADASTRO"].eq("FOLGUISTA")
        cadastro = cadastro[(cadastro["MOTORISTA_CADASTRO"] != "") & (cadastro["TIPO_CADASTRO"] != "")]
        return cadastro.drop_duplicates("MOTORISTA_CADASTRO", keep="last")

    def carregar_abastecimentos(self, mapa_frota: dict) -> pd.DataFrame:
        df = self._baixar_excel(self.config.LINK_ABASTECIMENTOS)

        col_placa = DataUtils.encontrar_coluna(df, ["PLACA", "CAVALO"])
        col_km = DataUtils.encontrar_coluna(df, ["KM ATUAL", "KM", "KM_1", "QUILOMETRAGEM"])
        col_litros = DataUtils.encontrar_coluna(df, ["QTDE", "LITROS", "QUANTIDADE", "QTD"])
        col_valor = DataUtils.encontrar_coluna(df, ["VALOR TOTAL", "VALOR", "TOTAL", "VALOR_TOTAL", "VR TOTAL", "VALOR COMBUSTIVEL"])
        col_motorista = DataUtils.encontrar_coluna(df, ["CONDUTOR", "MOTORISTA"])
        col_data = DataUtils.encontrar_coluna(df, ["DATA", "Data"])

        resultado = df.copy()
        resultado["_ORDEM_ORIGINAL"] = np.arange(len(resultado))
        resultado["PLACA_PADRONIZADA"] = resultado[col_placa].apply(DataUtils.padronizar_placa)
        resultado["KM_ATUAL_NUM"] = resultado[col_km].apply(DataUtils.converter_numero)
        resultado["QTDE_NUM"] = resultado[col_litros].apply(DataUtils.converter_numero)
        resultado["VALOR_NUM"] = resultado[col_valor].apply(DataUtils.converter_numero).fillna(0.0) if col_valor else 0.0
        resultado["CONDUTOR_NORMALIZADO"] = resultado[col_motorista].fillna("SEM MOTORISTA").apply(DataUtils.normalizar_texto)
        resultado["DATA_NUM"] = pd.to_datetime(resultado[col_data], errors="coerce", dayfirst=True) if col_data else pd.NaT

        resultado["TIPO"] = resultado["PLACA_PADRONIZADA"].map(mapa_frota)

        sem_tipo_mask = resultado["TIPO"].isna()
        placas_sem_tipo = resultado.loc[sem_tipo_mask, "PLACA_PADRONIZADA"].unique()
        mapa_mercosul = {}

        for p in placas_sem_tipo:
            eq = DataUtils.placa_equivalente_mercosul(p)
            if eq and eq in mapa_frota:
                mapa_mercosul[p] = mapa_frota[eq]

        for idx in resultado.index[sem_tipo_mask]:
            placa = resultado.at[idx, "PLACA_PADRONIZADA"]
            if placa in mapa_mercosul:
                resultado.at[idx, "TIPO"] = mapa_mercosul[placa]

        resultado["REGISTRO_VALIDO"] = (
            (resultado["PLACA_PADRONIZADA"] != "")
            & resultado["KM_ATUAL_NUM"].notna() & (resultado["KM_ATUAL_NUM"] > 0)
            & resultado["QTDE_NUM"].notna() & (resultado["QTDE_NUM"] > 0)
        )

        return resultado


# ================================================================
# MOTOR DE CÁLCULO DE KM, CONSUMO E PREMIAÇÃO
# ================================================================
class CalculadorPremio:
    @staticmethod
    def processar_consumo_por_abastecimento(df_abast: pd.DataFrame) -> pd.DataFrame:
        df = df_abast[df_abast["REGISTRO_VALIDO"]].copy()
        df = df.sort_values(by=["PLACA_PADRONIZADA", "KM_ATUAL_NUM", "DATA_NUM"])

        df["KM_ANTERIOR"] = df.groupby("PLACA_PADRONIZADA")["KM_ATUAL_NUM"].shift(1)
        df["KM_RODADO"] = df["KM_ATUAL_NUM"] - df["KM_ANTERIOR"]

        df["KM_RODADO_VALIDO"] = np.where((df["KM_RODADO"] > 0) & (df["KM_RODADO"] < 5000), df["KM_RODADO"], np.nan)
        
        # Divisão segura de consumo individual
        litros_seguros = df["QTDE_NUM"].replace(0, np.nan)
        df["CONSUMO_KML"] = df["KM_RODADO_VALIDO"] / litros_seguros

        return df.sort_values("_ORDEM_ORIGINAL")

    @staticmethod
    def calcular_premiacao_motoristas(
        df_abast_proc: pd.DataFrame,
        df_cadastro: pd.DataFrame,
        df_precos: pd.DataFrame
    ) -> pd.DataFrame:
        
        resumo_mot = df_abast_proc.groupby("CONDUTOR_NORMALIZADO").agg(
            KM_TOTAL=("KM_RODADO_VALIDO", "sum"),
            LITROS_TOTAL=("QTDE_NUM", "sum"),
            VALOR_TOTAL_GASS=("VALOR_NUM", "sum"),
            TOTAL_ABASTECIMENTOS=("REGISTRO_VALIDO", "count")
        ).reset_index()

        # Divisão segura de consumo por motorista
        resumo_mot["MEDIA_ALCANCADA"] = np.where(
            resumo_mot["LITROS_TOTAL"] > 0,
            resumo_mot["KM_TOTAL"] / resumo_mot["LITROS_TOTAL"],
            0.0
        )

        merged = pd.merge(
            resumo_mot,
            df_cadastro,
            left_on="CONDUTOR_NORMALIZADO",
            right_on="MOTORISTA_CADASTRO",
            how="left"
        )

        merged["TIPO_FINAL"] = merged["TIPO_CADASTRO"].fillna("NÃO CADASTRADO")
        merged["BASE_CADASTRO"] = merged["BASE_CADASTRO"].fillna("NÃO DEFINIDA")

        final_df = pd.merge(
            merged,
            df_precos,
            left_on="TIPO_FINAL",
            right_on="TIPO",
            how="left"
        )

        final_df["MEDIA_META"] = final_df["MEDIA"].fillna(0.0)
        final_df["VALOR_PREMIO_POTENCIAL"] = final_df["PREMIO"].fillna(0.0)

        final_df["ATINGIU_META"] = (
            (final_df["MEDIA_ALCANCADA"] >= final_df["MEDIA_META"]) & 
            (final_df["MEDIA_META"] > 0) &
            (final_df["KM_TOTAL"] >= 100)
        )

        final_df["PREMIO_RECEBER"] = np.where(
            final_df["ATINGIU_META"],
            final_df["VALOR_PREMIO_POTENCIAL"],
            0.0
        )

        # CÁLCULO SEGURO DE EFICIÊNCIA (Evita ZeroDivisionError)
        meta_segura = final_df["MEDIA_META"].replace(0, np.nan)
        eficiencia = (final_df["MEDIA_ALCANCADA"] / meta_segura) * 100
        final_df["EFICIENCIA_PCT"] = eficiencia.fillna(0.0)

        return final_df.sort_values(by="PREMIO_RECEBER", ascending=False)


# ================================================================
# DASHBOARD INTERATIVO STREAMLIT
# ================================================================
def main():
    st.set_page_config(
        page_title="Gestão de Premiação de Motoristas",
        page_icon="🚚",
        layout="wide"
    )

    st.title("🚚 Dashboard de Premiação e Eficiência de Frota")

    config = AppConfig()
    loader = DataLoader(config)

    @st.cache_data(ttl=600)
    def carregar_dados_completos():
        df_precos = loader.carregar_precos()
        _, mapa_frota = loader.carregar_frota()
        df_cad = loader.carregar_cadastro_motoristas()
        df_abast = loader.carregar_abastecimentos(mapa_frota)
        return df_precos, mapa_frota, df_cad, df_abast

    with st.spinner("Conectando às planilhas do Google Drive..."):
        try:
            df_precos, mapa_frota, df_cad, df_abast = carregar_dados_completos()
        except Exception as e:
            st.error(f"Erro ao carregar os dados: {e}")
            st.stop()

    df_abast_proc = CalculadorPremio.processar_consumo_por_abastecimento(df_abast)

    # FILTROS LATERAIS (SIDEBAR)
    st.sidebar.header("🔍 Filtros de Pesquisa")

    datas_validas = df_abast_proc["DATA_NUM"].dropna()
    min_date = datas_validas.min().date() if not datas_validas.empty else datetime.now().date()
    max_date = datas_validas.max().date() if not datas_validas.empty else datetime.now().date()

    data_inicio, data_fim = st.sidebar.date_input(
        "Período de Abastecimento",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

    mask_data = (df_abast_proc["DATA_NUM"].dt.date >= data_inicio) & (df_abast_proc["DATA_NUM"].dt.date <= data_fim)
    df_abast_filtrado = df_abast_proc[mask_data]

    bases_disponiveis = ["TODAS"] + sorted(list(df_cad["BASE_CADASTRO"].dropna().unique()))
    base_selecionada = st.sidebar.selectbox("Base Operacional", bases_disponiveis)

    tipos_disponiveis = ["TODOS"] + sorted(list(df_precos["TIPO"].dropna().unique()))
    tipo_selecionado = st.sidebar.selectbox("Tipo de Veículo", tipos_disponiveis)

    busca_motorista = st.sidebar.text_input("Buscar Motorista (Nome)")

    df_resultado = CalculadorPremio.calcular_premiacao_motoristas(df_abast_filtrado, df_cad, df_precos)

    if base_selecionada != "TODAS":
        df_resultado = df_resultado[df_resultado["BASE_CADASTRO"] == base_selecionada]

    if tipo_selecionado != "TODOS":
        df_resultado = df_resultado[df_resultado["TIPO_FINAL"] == tipo_selecionado]

    if busca_motorista:
        df_resultado = df_resultado[df_resultado["CONDUTOR_NORMALIZADO"].str.contains(DataUtils.normalizar_texto(busca_motorista))]

    # KPIS PRINCIPAIS (Divisão segura)
    total_motoristas = len(df_resultado)
    premiados = len(df_resultado[df_resultado["ATINGIU_META"]])
    total_pago = df_resultado["PREMIO_RECEBER"].sum()
    km_total_rodado = df_resultado["KM_TOTAL"].sum()
    litros_consumidos = df_resultado["LITROS_TOTAL"].sum()
    
    media_geral_frota = (km_total_rodado / litros_consumidos) if litros_consumidos > 0 else 0.0
    pct_premiados = (premiados / total_motoristas * 100) if total_motoristas > 0 else 0.0

    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Motoristas Analisados", f"{total_motoristas}")
    kpi2.metric("Motoristas Premiados", f"{premiados} ({pct_premiados:.1f}%)")
    kpi3.metric("Total de Prêmios (R$)", f"R$ {total_pago:,.2f}")
    kpi4.metric("KM Total Rodado", f"{km_total_rodado:,.0f} km")
    kpi5.metric("Média Geral Frota", f"{media_geral_frota:.2f} km/l")

    st.markdown("---")

    # ABAS
    tab1, tab2, tab3, tab4 = st.tabs([
        "🏆 Ranking e Premiação",
        "📊 Análise por Tipo de Veículo",
        "⚠️ Auditoria & Abastecimentos",
        "📥 Exportar Relatórios"
    ])

    with tab1:
        st.subheader("Resultado de Eficiência e Premiação por Motorista")

        col_tabela = [
            "CONDUTOR_NORMALIZADO", "BASE_CADASTRO", "TIPO_FINAL",
            "KM_TOTAL", "LITROS_TOTAL", "MEDIA_ALCANCADA", "MEDIA_META",
            "EFICIENCIA_PCT", "ATINGIU_META", "PREMIO_RECEBER"
        ]

        df_display = df_resultado[col_tabela].copy()
        df_display.columns = [
            "Motorista", "Base", "Tipo Veículo",
            "KM Rodado", "Litros", "Média (km/l)", "Meta (km/l)",
            "Eficiência (%)", "Atingiu Meta?", "Prêmio (R$)"
        ]

        st.dataframe(
            df_display.style.format({
                "KM Rodado": "{:,.1f}",
                "Litros": "{:,.1f}",
                "Média (km/l)": "{:.2f}",
                "Meta (km/l)": "{:.2f}",
                "Eficiência (%)": "{:.1f}%",
                "Prêmio (R$)": "R$ {:,.2f}"
            }),
            use_container_width=True,
            height=450
        )

    with tab2:
        st.subheader("Consumo Médio por Categoria de Veículo")
        resumo_veiculo = df_resultado.groupby("TIPO_FINAL").agg(
            QTD_MOTORISTAS=("CONDUTOR_NORMALIZADO", "count"),
            KM_TOTAL=("KM_TOTAL", "sum"),
            LITROS_TOTAL=("LITROS_TOTAL", "sum"),
            TOTAL_PREMIO=("PREMIO_RECEBER", "sum")
        ).reset_index()

        litros_cat_seguros = resumo_veiculo["LITROS_TOTAL"].replace(0, np.nan)
        resumo_veiculo["MEDIA_CATEGORIA"] = (resumo_veiculo["KM_TOTAL"] / litros_cat_seguros).fillna(0.0)

        st.dataframe(
            resumo_veiculo.style.format({
                "KM_TOTAL": "{:,.0f}",
                "LITROS_TOTAL": "{:,.0f}",
                "MEDIA_CATEGORIA": "{:.2f} km/l",
                "TOTAL_PREMIO": "R$ {:,.2f}"
            }),
            use_container_width=True
        )

    with tab3:
        st.subheader("Abastecimentos Processados")

        st.dataframe(
            df_abast_filtrado[[
                "DATA_NUM", "PLACA_PADRONIZADA", "CONDUTOR_NORMALIZADO",
                "KM_ATUAL_NUM", "KM_RODADO_VALIDO", "QTDE_NUM", "CONSUMO_KML", "TIPO"
            ]].head(100),
            use_container_width=True
        )

    with tab4:
        st.subheader("Download de Relatórios Formatados")
        st.write("Clique no botão abaixo para gerar a planilha em formato Excel (.xlsx).")

        bytes_excel = DataUtils.gerar_excel_download({
            "Ranking_Premiaçao": df_resultado,
            "Resumo_Veiculos": resumo_veiculo,
            "Base_Abastecimentos": df_abast_filtrado
        })

        st.download_button(
            label="📥 Baixar Relatório Excel Completo",
            data=bytes_excel,
            file_name=f"Relatorio_Premiacao_Motoristas_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
