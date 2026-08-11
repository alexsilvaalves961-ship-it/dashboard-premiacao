import os
import re
import io
import requests
import unicodedata
from typing import Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# ================================================================
# CONFIGURAÇÃO DA PÁGINA STREAMLIT
# ================================================================
st.set_page_config(
    page_title="Painel do Prêmio de Motoristas",
    page_icon="🚚",
    layout="wide"
)

# Estilização CSS para aproximar o visual do Streamlit ao layout original
st.markdown("""
<style>
    /* Estilo do Card do Cabeçalho */
    .header-card {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 15px 25px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border: 1px solid #E2E8F0;
        margin-bottom: 20px;
    }
    
    /* Logo Ciapetro */
    .ciapetro-flag {
        width: 140px; 
        height: 38px; 
        position: relative; 
        border-radius: 4px; 
        overflow: hidden; 
        display: flex; 
        flex-direction: column; 
        border: 1px solid #CBD5E1;
    }
    .flag-blue-top { height: 33.3%; background-color: #0099DA; }
    .flag-yellow { height: 33.3%; background-color: #FFD700; }
    .flag-blue-bottom { height: 33.3%; background-color: #1E2B7A; }
    .flag-text {
        position: absolute; 
        top: 50%; 
        left: 50%; 
        transform: translate(-50%, -50%); 
        font-family: 'Arial Black', sans-serif; 
        font-size: 14px; 
        font-weight: 900; 
        color: #FFFFFF; 
        text-shadow: 1px 1px 2px #000;
    }

    /* Cartões de Métricas no Topo */
    .kpi-box {
        background-color: #F8FAFC;
        border-radius: 8px;
        padding: 10px 15px;
        border: 1px solid #E2E8F0;
    }
    .kpi-title { font-size: 11px; font-weight: bold; color: #64748B; text-transform: uppercase; }
    .kpi-value { font-size: 18px; font-weight: bold; color: #0F172A; margin-top: 2px; }

    /* Estilo das Abas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
        border-bottom: 2px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        font-weight: 600;
        font-size: 14px;
        padding-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)


# ================================================================
# UTILITÁRIOS E TRATAMENTO DE TEXTO / PLACAS
# ================================================================
class DataUtils:
    NUMERO_PARA_LETRA = {"0": "A", "1": "B", "2": "C", "3": "D", "4": "E", "5": "F", "6": "G", "7": "H", "8": "I", "9": "J"}
    LETRA_PARA_NUMERO = {v: k for k, v in NUMERO_PARA_LETRA.items()}

    @staticmethod
    def normalizar_texto(valor) -> str:
        if pd.isna(valor):
            return ""
        texto = str(valor).strip().upper()
        texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
        return re.sub(r"\s+", " ", texto)

    @staticmethod
    def padronizar_placa(valor) -> str:
        if pd.isna(valor):
            return ""
        return re.sub(r"[^A-Z0-9]", "", str(valor).upper().strip())

    @classmethod
    def coluna_canonica(cls, nome: str) -> str:
        return re.sub(r"[^A-Z0-9]", "", cls.normalizar_texto(nome))

    @classmethod
    def encontrar_coluna(cls, df: pd.DataFrame, alternativas: list) -> Optional[str]:
        mapa = {cls.coluna_canonica(col): col for col in df.columns}
        for alt in alternativas:
            chave = cls.coluna_canonica(alt)
            if chave in mapa:
                return mapa[chave]
        return None

    @staticmethod
    def converter_numero(valor) -> float:
        if pd.isna(valor):
            return np.nan
        if isinstance(valor, (int, float, np.integer, np.floating)):
            return float(valor)
        texto = str(valor).strip().replace("R$", "").replace(" ", "")
        if not texto:
            return np.nan
        if "." in texto and "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        try:
            return float(texto)
        except ValueError:
            return np.nan

    @classmethod
    def placa_equivalente_mercosul(cls, placa: str) -> Optional[str]:
        placa = cls.padronizar_placa(placa)
        if len(placa) != 7:
            return None
        char = placa[4]
        if char in cls.NUMERO_PARA_LETRA:
            novo = cls.NUMERO_PARA_LETRA[char]
        elif char in cls.LETRA_PARA_NUMERO:
            novo = cls.LETRA_PARA_NUMERO[char]
        else:
            return None
        return placa[:4] + novo + placa[5:]


# ================================================================
# LINQUE E CARREGAMENTO DAS PLANILHAS
# ================================================================
LINK_ABASTECIMENTOS = "https://docs.google.com/spreadsheets/d/1YZBLfxOgJinm1TJHYI49AEaOWVPmOEiA/edit?usp=sharing&ouid=102045408189620250881&rtpof=true&sd=true"
LINK_METAS = "https://docs.google.com/spreadsheets/d/1vX2JqzFLcyDxytrBP5vbHCvMRJsRrm6M/edit?usp=sharing&ouid=102045408189620250881&rtpof=true&sd=true"

def gerar_url_download_direto(url: str) -> str:
    if "/d/" in url:
        file_id = url.split("/d/")[1].split("/")[0]
    elif "id=" in url:
        file_id = url.split("id=")[1].split("&")[0]
    else:
        file_id = url
    return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"

class DataLoader:
    def __init__(self, precos_src, frota_src, motoristas_src, abastecimentos_src):
        self.precos_src = precos_src
        self.frota_src = frota_src
        self.motoristas_src = motoristas_src
        self.abastecimentos_src = abastecimentos_src

    def carregar_precos(self) -> pd.DataFrame:
        df = pd.read_excel(self.precos_src, sheet_name="Planilha1")
        col_tipo = DataUtils.encontrar_coluna(df, ["TIPO", "TIPO VEICULO", "TIPO DE VEICULO"])
        col_media = DataUtils.encontrar_coluna(df, ["MEDIA", "MÉDIA"])
        col_premio = DataUtils.encontrar_coluna(df, ["TOTAL", "PREMIO", "PRÊMIO", "VALOR"])

        if None in (col_tipo, col_media, col_premio) and df.shape[1] == 5:
            df = df.iloc[1:].copy()
            df.columns = ["DEL", "TIPO", "FATOR", "MEDIA", "PREMIO"]
            col_tipo, col_media, col_premio = "TIPO", "MEDIA", "PREMIO"

        resultado = pd.DataFrame({
            "TIPO": df[col_tipo].apply(DataUtils.normalizar_texto),
            "MEDIA": df[col_media].apply(DataUtils.converter_numero),
            "PREMIO": df[col_premio].apply(DataUtils.converter_numero),
        }).dropna(subset=["MEDIA", "PREMIO"])

        resultado["TIPO"] = resultado["TIPO"].replace({"TOCO": "TRUCK"})
        return resultado[resultado["TIPO"] != ""].reset_index(drop=True)

    def carregar_frota(self) -> Tuple[pd.DataFrame, dict]:
        df = pd.read_excel(self.frota_src)
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
        bruto = pd.read_excel(self.motoristas_src, header=None)
        cab_idx, linha_cab = None, None

        for i in range(min(len(bruto), 15)):
            vals = [str(x).strip().upper() for x in bruto.iloc[i].tolist()]
            if "MOTORISTAS" in vals and "TIPO" in vals:
                cab_idx, linha_cab = i, vals
                break

        if cab_idx is None:
            raise ValueError("Cabeçalho não encontrado no arquivo de Motoristas.")

        idx_mot = linha_cab.index("MOTORISTAS")
        idx_tipo = linha_cab.index("TIPO")
        idx_base = linha_cab.index("BASE") if "BASE" in linha_cab else None

        cadastro = bruto.iloc[cab_idx + 1:].copy()
        cadastro["MOTORISTA_CADASTRO"] = cadastro.iloc[:, idx_mot].apply(DataUtils.normalizar_texto)
        cadastro["TIPO_CADASTRO"] = cadastro.iloc[:, idx_tipo].apply(DataUtils.normalizar_texto).replace({"TOCO": "TRUCK"})
        cadastro["BASE_CADASTRO"] = cadastro.iloc[:, idx_base].apply(DataUtils.normalizar_texto) if idx_base is not None else ""
        cadastro["EH_FOLGUISTA"] = cadastro["TIPO_CADASTRO"].eq("FOLGUISTA")

        cadastro = cadastro[(cadastro["MOTORISTA_CADASTRO"] != "") & (cadastro["TIPO_CADASTRO"] != "")]
        return cadastro.drop_duplicates("MOTORISTA_CADASTRO", keep="last")

    def carregar_abastecimentos(self, mapa_frota: dict) -> pd.DataFrame:
        df = pd.read_excel(self.abastecimentos_src)

        col_placa = DataUtils.encontrar_coluna(df, ["PLACA", "CAVALO"])
        col_km = DataUtils.encontrar_coluna(df, ["KM ATUAL", "KM", "KM_1", "QUILOMETRAGEM"])
        col_litros = DataUtils.encontrar_coluna(df, ["QTDE", "LITROS", "QUANTIDADE", "QTD"])
        col_valor = DataUtils.encontrar_coluna(df, ["VALOR TOTAL", "VALOR", "TOTAL", "VALOR_TOTAL", "VR TOTAL", "VLR TOTAL", "VALOR COMBUSTIVEL", "VALOR (R$)"])
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
# MOTOR DE REGRAS E CÁLCULOS
# ================================================================
class RewardEngine:
    @staticmethod
    def calcular_eventos_consumo(abastecimentos: pd.DataFrame) -> pd.DataFrame:
        base = abastecimentos[abastecimentos["REGISTRO_VALIDO"]].copy()
        base["_DATA_ORDENACAO"] = base["DATA_NUM"].fillna(pd.Timestamp("1900-01-01"))
        base = base.sort_values(["PLACA_PADRONIZADA", "_DATA_ORDENACAO", "_ORDEM_ORIGINAL"], kind="stable").copy()

        base["KM_ANTERIOR"] = base.groupby("PLACA_PADRONIZADA")["KM_ATUAL_NUM"].shift(1)
        base["KM_RODADO_EVENTO"] = base["KM_ATUAL_NUM"] - base["KM_ANTERIOR"]

        base["REGISTRO_CONSUMO_VALIDO"] = (
            base["KM_ANTERIOR"].notna()
            & base["KM_RODADO_EVENTO"].notna() & (base["KM_RODADO_EVENTO"] > 0)
            & base["QTDE_NUM"].notna() & (base["QTDE_NUM"] > 0)
        )

        base["KM_CONSUMO"] = np.where(base["REGISTRO_CONSUMO_VALIDO"], base["KM_RODADO_EVENTO"], np.nan)
        base["LITROS_CONSUMO"] = np.where(base["REGISTRO_CONSUMO_VALIDO"], base["QTDE_NUM"], np.nan)
        base["TIPO_CALCULO"] = base["TIPO"].fillna("GERAL").replace({"TOCO": "TRUCK"})

        return base.reset_index(drop=True)

    def faixa_mais_proxima(self, media: float, tipo: str, precos: pd.DataFrame) -> dict:
        if pd.isna(media):
            return {"MEDIA_FAIXA": np.nan, "PREMIO": 0.0, "STATUS_PREMIO": "SEM MEDIA"}

        tabela = precos[precos["TIPO"] == tipo].copy() if tipo in precos["TIPO"].values else precos.copy()
        if tabela.empty:
            return {"MEDIA_FAIXA": np.nan, "PREMIO": 0.0, "STATUS_PREMIO": "SEM FAIXA"}

        tabela = tabela.sort_values("MEDIA").copy()
        media_min, media_max = float(tabela["MEDIA"].min()), float(tabela["MEDIA"].max())
        media_class = round(float(media), 2)

        if media_class < media_min:
            return {"MEDIA_FAIXA": np.nan, "PREMIO": 0.0, "STATUS_PREMIO": "DESCLASSIFICADO"}

        if media_class > media_max:
            linha_max = tabela.iloc[-1]
            return {"MEDIA_FAIXA": float(linha_max["MEDIA"]), "PREMIO": float(linha_max["PREMIO"]), "STATUS_PREMIO": "TETO"}

        medias = tabela["MEDIA"].astype(float).to_numpy()
        pos = max(0, min(int(np.searchsorted(medias, media_class, side="right") - 1), len(tabela) - 1))
        linha = tabela.iloc[pos]

        return {"MEDIA_FAIXA": float(linha["MEDIA"]), "PREMIO": float(linha["PREMIO"]), "STATUS_PREMIO": "OK"}

    def calcular_premios(self, eventos: pd.DataFrame, precos: pd.DataFrame, cadastro: pd.DataFrame) -> pd.DataFrame:
        base = eventos[eventos["REGISTRO_CONSUMO_VALIDO"]].copy()
        base["CATEGORIA_ABASTECIMENTO"] = base["TIPO"].fillna("GERAL").replace({"TOCO": "TRUCK"})
        base["MOTORISTA_CHAVE"] = base["CONDUTOR_NORMALIZADO"].fillna("").astype(str).str.strip().str.upper()

        cad = cadastro.copy()
        cad["MOTORISTA_CHAVE"] = cad["MOTORISTA_CADASTRO"].astype(str).str.strip().str.upper()

        base = base.merge(cad[["MOTORISTA_CHAVE", "TIPO_CADASTRO", "BASE_CADASTRO", "EH_FOLGUISTA"]], on="MOTORISTA_CHAVE", how="left")
        base["TIPO_CADASTRO"] = base["TIPO_CADASTRO"].fillna("")
        base["EH_FOLGUISTA"] = base["EH_FOLGUISTA"].fillna(False)

        registros = []
        for _, grupo in base.groupby("MOTORISTA_CHAVE", sort=False):
            eh_folguista = grupo["EH_FOLGUISTA"].iloc[0]
            tipo_cad = grupo["TIPO_CADASTRO"].iloc[0]

            if eh_folguista:
                soma = grupo.groupby("CATEGORIA_ABASTECIMENTO")["KM_CONSUMO"].sum()
                cat_elegivel = soma.idxmax() if not soma.empty else tipo_cad
            else:
                cat_elegivel = tipo_cad if tipo_cad else grupo["CATEGORIA_ABASTECIMENTO"].iloc[0]

            sub = grupo[grupo["CATEGORIA_ABASTECIMENTO"] == cat_elegivel]
            if not sub.empty:
                registros.append(sub)

        if not registros:
            return pd.DataFrame()

        df_eleg = pd.concat(registros, ignore_index=True)

        resumo = df_eleg.groupby(["MOTORISTA_CHAVE", "CATEGORIA_ABASTECIMENTO"], as_index=False).agg(
            MOTORISTA=("CONDUTOR_NORMALIZADO", "first"),
            BASE=("BASE_CADASTRO", "first"),
            CATEGORIA=("CATEGORIA_ABASTECIMENTO", "first"),
            KM_TOTAL=("KM_CONSUMO", "sum"),
            LITROS_TOTAL=("LITROS_CONSUMO", "sum"),
            QTD_ABASTECIMENTOS=("LITROS_CONSUMO", "count"),
            PLACAS=("PLACA_PADRONIZADA", lambda s: " | ".join(sorted(set(s)))),
        )

        resumo["MEDIA_CALCULADA"] = np.where(resumo["LITROS_TOTAL"] > 0, resumo["KM_TOTAL"] / resumo["LITROS_TOTAL"], np.nan)
        faixas = resumo.apply(lambda r: self.faixa_mais_proxima(r["MEDIA_CALCULADA"], r["CATEGORIA"], precos), axis=1)

        resumo_df = pd.concat([resumo, pd.DataFrame(list(faixas))], axis=1)
        resumo_df["PREMIO_BRUTO"] = resumo_df["PREMIO"]
        resumo_df["DIAS_AUSENCIA"] = 0
        resumo_df["DIAS_EFETIVOS"] = 30

        return resumo_df


# ================================================================
# REGRAS DO PILAR 1 E FUNÇÕES DE RELATÓRIO / RECIBO
# ================================================================
CRITERIOS_PILAR_1 = [
    "1 - Controle de velocidade, limite máx de 80 km/h (1 Ponto/evento até 129)",
    "2 - Controle de jornada, macros e intervalos incorretos (1 Ponto/evento até 129)",
    "3 - Não realização correta do check list diário (1 Ponto/evento até 129)",
    "4 - Deszelo com documentação e comprovantes de carga/descarga (1 Ponto/evento até 129)",
    "5 - Distração com freadas bruscas, risco de colisão ou manobra perigosa [DESCLASSIFICADO]",
    "6 - Uso do celular em direção [DESCLASSIFICADO]",
    "7 - Estacionar ou parada com veículo em L [DESCLASSIFICADO]",
    "8 - Identificação de ausência do motorista [DESCLASSIFICADO]",
    "9 - Paradas em locais proibidos [DESCLASSIFICADO]",
    "10 - Ausência do cinto de segurança em direção [DESCLASSIFICADO]",
    "11 - Vedar as câmeras [DESCLASSIFICADO]",
    "12 - Comportamento inadequado do motorista dentro e fora do veículo [DESCLASSIFICADO]",
    "13 - Picos acima de 80 km/h, com mais de 30 infrações por período [DESCLASSIFICADO]",
    "14 - Não cumprir determinações e escalas estipuladas [DESCLASSIFICADO]",
    "15 - Erros operacionais (Carregamento/Descarregamento incorreto ou Derramamento/Contaminação) [DESCLASSIFICADO]"
]

def aplicar_regras_gerais(df_resumo_original: pd.DataFrame, df_ausencias: pd.DataFrame, df_desclassificacoes: pd.DataFrame) -> pd.DataFrame:
    if df_resumo_original.empty:
        return df_resumo_original.copy()

    res = df_resumo_original.copy()

    motivos_iniciais = []
    for idx in res.index:
        st_p = res.at[idx, "STATUS_PREMIO"]
        if st_p == "DESCLASSIFICADO":
            motivos_iniciais.append("Média de consumo abaixo do limite mínimo da categoria")
        else:
            motivos_iniciais.append("Elegível / Em conformidade")
    res["MOTIVO_DESCLASSIFICACAO"] = motivos_iniciais

    if not df_ausencias.empty:
        soma_dias = df_ausencias.groupby("MOTORISTA")["DIAS"].sum().to_dict()
        res["DIAS_AUSENCIA"] = res["MOTORISTA"].map(soma_dias).fillna(0).astype(int)
        res["DIAS_EFETIVOS"] = np.maximum(0, 30 - res["DIAS_AUSENCIA"])
        res["PREMIO"] = res.apply(lambda r: max(0.0, r["PREMIO_BRUTO"] * (r["DIAS_EFETIVOS"] / 30.0)), axis=1)
    else:
        res["DIAS_AUSENCIA"] = 0
        res["DIAS_EFETIVOS"] = 30
        res["PREMIO"] = res["PREMIO_BRUTO"]

    if not df_desclassificacoes.empty:
        for idx in res.index:
            m_nome = res.at[idx, "MOTORISTA"]
            g = df_desclassificacoes[df_desclassificacoes["MOTORISTA"] == m_nome]
            if not g.empty:
                motivos_pilar1 = []
                diretos = g[g["TIPO_IMPACTO"] == "DESCLASSIFICADO"]
                if not diretos.empty:
                    crit_limpos = [c.split("[")[0].strip() for c in diretos["CRITERIO"].unique()]
                    motivos_pilar1.append(f"Infração Pilar 1: {', '.join(crit_limpos)}")

                tot_pontos = g["PONTOS"].sum()
                if tot_pontos > 129:
                    motivos_pilar1.append(f"Excesso de Pontos Pilar 1 ({tot_pontos} pts - máx 129)")

                if motivos_pilar1:
                    res.at[idx, "PREMIO"] = 0.0
                    res.at[idx, "STATUS_PREMIO"] = "DESCLASSIFICADO"

                    mot_existente = res.at[idx, "MOTIVO_DESCLASSIFICACAO"]
                    novo_mot = " | ".join(motivos_pilar1)
                    if mot_existente and "Média" in mot_existente:
                        res.at[idx, "MOTIVO_DESCLASSIFICACAO"] = f"{mot_existente} + {novo_mot}"
                    else:
                        res.at[idx, "MOTIVO_DESCLASSIFICACAO"] = novo_mot

    return res

def gerar_tabela_rh(df_resumo: pd.DataFrame) -> pd.DataFrame:
    if df_resumo.empty:
        return pd.DataFrame(columns=["NOME", "FILIAL", "VALOR PAGO"])
    
    rh_df = pd.DataFrame()
    rh_df["NOME"] = df_resumo["MOTORISTA"]
    rh_df["FILIAL"] = df_resumo["BASE"].fillna("CIANORTE")
    rh_df["VALOR PAGO"] = df_resumo["PREMIO"].map(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    return rh_df

def gerar_html_unico_recibo(row_data: pd.Series, motorista_sel: str, periodo_ini: str, periodo_fim: str, fator_c: str) -> str:
    base_val = row_data.get("BASE", "")
    if pd.isna(base_val) or str(base_val).strip() == "":
        base_val = "CIANORTE"

    tipo_val = row_data.get("CATEGORIA", "")
    km_val = f"{row_data.get('KM_TOTAL', 0):,.0f}".replace(",", ".")
    dias_efetivos_val = row_data.get("DIAS_EFETIVOS", 30)

    media_raw = row_data.get("MEDIA_CALCULADA", np.nan)
    media_val = f"{media_raw:.2f}".replace(".", ",") if pd.notna(media_raw) and isinstance(media_raw, (int, float)) else str(media_raw)

    premio_raw = row_data.get("PREMIO", 0.0)
    val_total_str = f"R$ {premio_raw:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    motivo_descl = row_data.get("MOTIVO_DESCLASSIFICACAO", "Elegível / Em conformidade")
    status_premio = row_data.get("STATUS_PREMIO", "OK")
    eh_desclassificado = (status_premio == "DESCLASSIFICADO") or (premio_raw == 0.0 and motivo_descl != "Elegível / Em conformidade")

    bg_motivo = "#FEE2E2" if eh_desclassificado else "#FFFFFF"
    fg_motivo = "#991B1B" if eh_desclassificado else "#000000"
    hdr_motivo_bg = "#EF4444" if eh_desclassificado else "#D0E0F0"
    hdr_motivo_fg = "#FFFFFF" if eh_desclassificado else "#000000"

    return f"""
    <div style="background-color: #FFFFFF; padding: 24px; border-radius: 12px; max-width: 650px; margin: 15px auto; font-family: Arial, sans-serif; color: #000000; border: 1px solid #CBD5E1; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);">
        <div style="text-align: center; margin-bottom: 12px;">
            <div style="display: inline-block; width: 190px;">
                <div style="width: 190px; height: 46px; position: relative; border-radius: 2px; overflow: hidden; display: flex; flex-direction: column; border: 1px solid #CBD5E1;">
                    <div style="height: 33.3%; background-color: #0099DA;"></div>
                    <div style="height: 33.3%; background-color: #FFD700;"></div>
                    <div style="height: 33.3%; background-color: #1E2B7A;"></div>
                    <span style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); font-family: 'Arial Black', sans-serif; font-size: 18px; font-weight: 900; color: #FFFFFF; text-shadow: 1px 1px 2px #000;">Ciapetro</span>
                </div>
            </div>
        </div>

        <h3 style="text-align: center; margin: 10px 0 16px 0; font-size: 18px; font-weight: bold; color: #000000;">Recibo de Premiação</h3>

        <table style="width: 100%; border-collapse: collapse; border: 2px solid #000000; font-size: 12px; font-weight: bold;">
            <tr style="border-bottom: 1px solid #000000;">
                <td style="width: 45%; background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">MOTORISTA</td>
                <td style="width: 55%; padding: 4px 8px; text-align: center; background-color: #FFFFFF;">{motorista_sel}</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">BASE</td>
                <td style="padding: 4px 8px; text-align: center;">{base_val}</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">TIPO</td>
                <td style="padding: 4px 8px; text-align: center;">{tipo_val}</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">DIAS EFETIVOS</td>
                <td style="padding: 4px 8px; text-align: center;">{dias_efetivos_val}</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">KM RODADO</td>
                <td style="padding: 4px 8px; text-align: center;">{km_val}</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">MEDIA</td>
                <td style="padding: 4px 8px; text-align: center;">{media_val}</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">FATOR CARGA</td>
                <td style="padding: 4px 8px; text-align: center;">{fator_c}</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000; background-color: {bg_motivo};">
                <td style="background-color: {hdr_motivo_bg}; color: {hdr_motivo_fg}; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">MOTIVO DESCLASSIFICAÇÃO</td>
                <td style="padding: 4px 8px; text-align: center; color: {fg_motivo}; font-weight: bold;">{motivo_descl}</td>
            </tr>
            <tr>
                <td style="background-color: #00A3E0; color: #FFFFFF; padding: 6px 8px; border-right: 1px solid #000000; text-align: center; font-size: 13px;">VALOR TOTAL</td>
                <td style="padding: 6px 8px; text-align: center; font-size: 13px; background-color: #FFFFFF; font-weight: bold;">{val_total_str}</td>
            </tr>
        </table>

        <p style="text-align: center; margin-top: 18px; margin-bottom: 20px; font-size: 11px; line-height: 1.4; color: #000000;">
            Eu, <strong>{motorista_sel}</strong>, conferi e concordo com as informações, pois estão de acordo com a Política de Premiação dos Motoristas.
        </p>

        <div style="margin-top: 25px; text-align: left; font-size: 11px; color: #000000;">
            <span>Ass.: __________________________________________________</span><br/>
            <div style="margin-left: 35px; margin-top: 4px; font-weight: bold;">{motorista_sel}</div>
        </div>

        <div style="margin-top: 15px; text-align: left; font-size: 11px; color: #000000;">
            LOCAL/DATA _____________, ______/______/ 2026
        </div>

        <div style="margin-top: 12px; text-align: center; font-size: 11px; font-weight: bold; color: #000000;">
            Período de Controle: {periodo_ini} a {periodo_fim}
        </div>
    </div>
    """


# ================================================================
# GERENCIAMENTO DE ESTADO E CARREGAMENTO
# ================================================================
if "df_ausencias" not in st.session_state:
    st.session_state.df_ausencias = pd.DataFrame(columns=["MOTORISTA", "TIPO_AUSENCIA", "DATA_INICIO", "DIAS", "OBSERVACAO"])

if "df_desclassificacoes" not in st.session_state:
    st.session_state.df_desclassificacoes = pd.DataFrame(columns=["MOTORISTA", "CRITERIO", "PONTOS", "TIPO_IMPACTO", "OBSERVACAO"])

@st.cache_data(ttl=600, show_spinner=False)
def carregar_dados_drive():
    def obter_fonte(path_local, drive_link=None):
        if os.path.exists(path_local):
            return path_local
        if drive_link:
            try:
                url_dl = gerar_url_download_direto(drive_link)
                res = requests.get(url_dl)
                if res.status_code == 200:
                    return io.BytesIO(res.content)
            except Exception:
                return None
        return None

    src_p = obter_fonte("Pasta2.xlsx")
    src_f = obter_fonte("frota.xlsx")
    src_m = obter_fonte("Pasta4.xlsx", LINK_METAS)
    src_a = obter_fonte("uah abastecimentos_3.xlsx", LINK_ABASTECIMENTOS)

    if not all([src_p, src_f, src_m, src_a]):
        return None, None

    loader = DataLoader(src_p, src_f, src_m, src_a)
    engine = RewardEngine()

    precos = loader.carregar_precos()
    _, mapa_frota = loader.carregar_frota()
    cadastro = loader.carregar_cadastro_motoristas()
    abastecimentos = loader.carregar_abastecimentos(mapa_frota)

    eventos = engine.calcular_eventos_consumo(abastecimentos)
    resumo_base = engine.calcular_premios(eventos, precos, cadastro)

    return resumo_base, eventos

with st.spinner("⏳ Carregando dados de consumo e cadastros..."):
    resumo_base, eventos = carregar_dados_drive()

if resumo_base is None:
    st.error("⚠️ Erro ao carregar planilhas locais ou do Google Drive. Verifique se os arquivos estão na pasta do app.")
    st.stop()

# Aplica ausências e infrações
resumo_atualizado = aplicar_regras_gerais(resumo_base, st.session_state.df_ausencias, st.session_state.df_desclassificacoes)


# ================================================================
# 1. CABEÇALHO DO PAINEL (LOGO CIAPETRO + TÍTULO DA FOTO)
# ================================================================
st.markdown("""
<div class="header-card">
    <div style="display: flex; align-items: center; gap: 20px;">
        <div class="ciapetro-flag">
            <div class="flag-blue-top"></div>
            <div class="flag-yellow"></div>
            <div class="flag-blue-bottom"></div>
            <span class="flag-text">Ciapetro</span>
        </div>
        <div>
            <h2 style="margin: 0; font-size: 22px; font-weight: 800; color: #0F172A;">Painel do Prêmio de Motoristas</h2>
            <p style="margin: 2px 0 0 0; font-size: 13px; color: #64748B;">Visão gerencial de consumo, desempenho e prêmio</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


# ================================================================
# 2. FILTROS PRINCIPAIS NO CORPO (EXATAMENTE COMO NA FOTO)
# ================================================================
mots_lista = ["TODOS"] + sorted(list(resumo_atualizado["MOTORISTA"].dropna().unique()))
cats_lista = ["TODAS"] + sorted(list(resumo_atualizado["CATEGORIA"].dropna().unique()))
filiais_lista = ["TODAS"] + sorted([str(b) for b in resumo_atualizado["BASE"].dropna().unique() if str(b).strip() != ""])

f_col1, f_col2, f_col3, f_col4 = st.columns(4)

with f_col1:
    f_mot = st.selectbox("👤 Filtrar Motorista", options=mots_lista, index=0)
with f_col2:
    f_plc = st.text_input("🔍 Filtrar Placa", placeholder="Ex: ABC1234 ou ABC1D23")
with f_col3:
    f_cat = st.selectbox("🏷️ Categoria de Veículo", options=cats_lista, index=0)
with f_col4:
    f_fil = st.selectbox("🏢 Filtrar Filial / Base", options=filiais_lista, index=0)

# Filtragem lógica dos dados
res_f = resumo_atualizado.copy()
evt_f = eventos.copy() if eventos is not None else pd.DataFrame()

if f_mot != "TODOS":
    m_norm = DataUtils.normalizar_texto(f_mot)
    res_f = res_f[res_f["MOTORISTA"].apply(DataUtils.normalizar_texto).str.contains(m_norm, na=False)]
    if not evt_f.empty:
        evt_f = evt_f[evt_f["CONDUTOR_NORMALIZADO"].apply(DataUtils.normalizar_texto).str.contains(m_norm, na=False)]

if f_plc.strip():
    p_norm = DataUtils.padronizar_placa(f_plc)
    res_f = res_f[res_f["PLACAS"].apply(DataUtils.padronizar_placa).str.contains(p_norm, na=False)]
    if not evt_f.empty:
        evt_f = evt_f[evt_f["PLACA_PADRONIZADA"] == p_norm]

if f_cat != "TODAS":
    c_norm = DataUtils.normalizar_texto(f_cat)
    res_f = res_f[res_f["CATEGORIA"] == c_norm]
    if not evt_f.empty:
        evt_f = evt_f[evt_f["TIPO_CALCULO"] == c_norm]

if f_fil != "TODAS":
    f_norm = DataUtils.normalizar_texto(f_fil)
    res_f = res_f[res_f["BASE"].apply(DataUtils.normalizar_texto) == f_norm]


# ================================================================
# 3. CARTÕES DE MÉTRICAS / KPIS (LAYOUT DA FOTO)
# ================================================================
tot_premio = res_f["PREMIO"].sum() if "PREMIO" in res_f.columns else 0.0
tot_km = res_f["KM_TOTAL"].sum() if "KM_TOTAL" in res_f.columns else 0.0
tot_litros = res_f["LITROS_TOTAL"].sum() if "LITROS_TOTAL" in res_f.columns else 0.0
tot_gasto = evt_f["VALOR_NUM"].sum() if not evt_f.empty and "VALOR_NUM" in evt_f.columns else 0.0
tot_media = (tot_km / tot_litros) if tot_litros > 0 else 0.0
tot_mots = len(res_f)

k1, k2, k3, k4, k5, k6 = st.columns(6)

with k1:
    st.markdown(f'<div class="kpi-box"><div class="kpi-title">💰 Total em Prêmios</div><div class="kpi-value">R$ {tot_premio:,.2f}</div></div>', unsafe_allow_html=True)
with k2:
    st.markdown(f'<div class="kpi-box"><div class="kpi-title">💳 Total Gasto Combustível</div><div class="kpi-value">R$ {tot_gasto:,.2f}</div></div>', unsafe_allow_html=True)
with k3:
    st.markdown(f'<div class="kpi-box"><div class="kpi-title">🛣️ TOTAL RODADO (KM)</div><div class="kpi-value">{tot_km:,.0f} km</div></div>', unsafe_allow_html=True)
with k4:
    st.markdown(f'<div class="kpi-box"><div class="kpi-title">⛽ TOTAL COMBUSTÍVEL</div><div class="kpi-value">{tot_litros:,.0f} L</div></div>', unsafe_allow_html=True)
with k5:
    st.markdown(f'<div class="kpi-box"><div class="kpi-title">🎯 MÉDIA GERAL (KM/L)</div><div class="kpi-value">{tot_media:.2f} km/L</div></div>', unsafe_allow_html=True)
with k6:
    st.markdown(f'<div class="kpi-box"><div class="kpi-title">👥 MOTORISTAS NA LISTA</div><div class="kpi-value">{tot_mots}</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ================================================================
# 4. AS 6 ABAS EXATAS DA SUAS IMAGENS DA FOTO
# ================================================================
aba1, aba2, aba3, aba4, aba5, aba6 = st.tabs([
    "📊 Resumo de Premiações por Motorista",
    "👔 Relatório RH - Lançamento de Pagamento",
    "📋 Detalhamento dos Abastecimentos",
    "📄 Recibo de Premiação",
    "🏥 Lançamento de Atestados e Férias",
    "🚫 Gestão de Desclassificações (Pilar 1)"
])

# ----------------------------------------------------------------
# ABA 1: RESUMO DE PREMIAÇÕES POR MOTORISTA
# ----------------------------------------------------------------
with aba1:
    st.subheader("📊 Resumo de Premiações por Motorista")
    
    df_exibicao = res_f.copy()
    if not df_exibicao.empty:
        df_exibicao["PREMIO_FORMATADO"] = df_exibicao["PREMIO"].map(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        cols_ordem = ["MOTORISTA", "BASE", "CATEGORIA", "KM_TOTAL", "LITROS_TOTAL", "MEDIA_CALCULADA", "STATUS_PREMIO", "PREMIO_FORMATADO", "MOTIVO_DESCLASSIFICACAO"]
        cols_presentes = [c for c in cols_ordem if c in df_exibicao.columns]
        st.dataframe(df_exibicao[cols_presentes], use_container_width=True, height=450)
    else:
        st.info("Nenhum registro encontrado para os filtros selecionados.")

# ----------------------------------------------------------------
# ABA 2: RELATÓRIO RH - LANÇAMENTO DE PAGAMENTO
# ----------------------------------------------------------------
with aba2:
    st.subheader("👔 Relatório RH - Lançamento de Pagamento")
    st.caption("Relação simplificada enviada ao RH para lançamento em folha de pagamento.")
    
    df_rh = gerar_tabela_rh(res_f)
    st.dataframe(df_rh, use_container_width=True, height=450)

# ----------------------------------------------------------------
# ABA 3: DETALHAMENTO DOS ABASTECIMENTOS
# ----------------------------------------------------------------
with aba3:
    st.subheader("📋 Detalhamento dos Abastecimentos Processados")
    if not evt_f.empty:
        st.dataframe(evt_f, use_container_width=True, height=450)
    else:
        st.info("Nenhum abastecimento encontrado.")

# ----------------------------------------------------------------
# ABA 4: RECIBO DE PREMIAÇÃO
# ----------------------------------------------------------------
with aba4:
    st.subheader("📄 Gerador de Recibo de Premiação")
    
    rc1, rc2, rc3, rc4, rc5 = st.columns(5)
    with rc1:
        rec_filial = st.selectbox("Filial", options=filiais_lista, key="rec_filial_tab")
    
    mots_rec_opt = ["SELECIONE...", "TODOS OS MOTORISTAS"]
    if rec_filial != "TODAS":
        f_norm = DataUtils.normalizar_texto(rec_filial)
        mots_rec_opt += sorted(list(resumo_atualizado[resumo_atualizado["BASE"].apply(DataUtils.normalizar_texto) == f_norm]["MOTORISTA"].dropna().unique()))
    else:
        mots_rec_opt += sorted(list(resumo_atualizado["MOTORISTA"].dropna().unique()))

    with rc2:
        rec_motorista = st.selectbox("Motorista", options=mots_rec_opt, key="rec_motorista_tab")
    with rc3:
        rec_ini = st.text_input("Data Início", value="26/06/2026")
    with rc4:
        rec_fim = st.text_input("Data Fim", value="25/07/2026")
    with rc5:
        rec_fator = st.text_input("Fator Carga", value="50%")

    if rec_motorista != "SELECIONE...":
        if rec_motorista == "TODOS OS MOTORISTAS":
            sub_mots = resumo_atualizado["MOTORISTA"].unique()
        else:
            sub_mots = [rec_motorista]

        html_recibos = ""
        for m_nome in sub_mots:
            row_m = resumo_atualizado[resumo_atualizado["MOTORISTA"] == m_nome]
            if not row_m.empty:
                html_recibos += gerar_html_unico_recibo(row_m.iloc[0], m_nome, rec_ini, rec_fim, rec_fator)
        
        st.markdown(html_recibos, unsafe_allow_html=True)
    else:
        st.info("👉 Selecione um motorista para visualizar o recibo de pagamento.")

# ----------------------------------------------------------------
# ABA 5: LANÇAMENTO DE ATESTADOS E FÉRIAS
# ----------------------------------------------------------------
with aba5:
    st.subheader("🏥 Lançamento de Atestados e Férias")
    st.write("Registre os dias de ausência para cálculo proporcional da premiação.")
    
    with st.form("form_atestados", clear_on_submit=True):
        a_col1, a_col2, a_col3 = st.columns(3)
        with a_col1:
            mot_aus = st.selectbox("Motorista", options=sorted(list(resumo_atualizado["MOTORISTA"].dropna().unique())))
            tipo_aus = st.radio("Tipo de Ausência", ["Atestado Médico", "Férias", "Outro Afastamento"])
        with a_col2:
            dt_aus = st.text_input("Data Início", value="01/07/2026")
            dias_aus = st.number_input("Dias Ausente", min_value=1, max_value=30, value=1)
        with a_col3:
            obs_aus = st.text_area("Observação / CID", placeholder="Detalhes do afastamento...")
        
        btn_add_aus = st.form_submit_button("➕ Registrar Ausência")

    if btn_add_aus:
        nova_aus = pd.DataFrame([{
            "MOTORISTA": mot_aus,
            "TIPO_AUSENCIA": tipo_aus,
            "DATA_INICIO": dt_aus,
            "DIAS": int(dias_aus),
            "OBSERVACAO": obs_aus
        }])
        st.session_state.df_ausencias = pd.concat([st.session_state.df_ausencias, nova_aus], ignore_index=True)
        st.success(f"Ausência de {dias_aus} dia(s) gravada para {mot_aus}!")
        st.rerun()

    st.markdown("#### Histórico de Ausências Registradas")
    st.dataframe(st.session_state.df_ausencias, use_container_width=True)
    
    if not st.session_state.df_ausencias.empty:
        if st.button("🗑️ Limpar Atestados/Férias"):
            st.session_state.df_ausencias = pd.DataFrame(columns=["MOTORISTA", "TIPO_AUSENCIA", "DATA_INICIO", "DIAS", "OBSERVACAO"])
            st.rerun()

# ----------------------------------------------------------------
# ABA 6: GESTÃO DE DESCLASSIFICAÇÕES (PILAR 1)
# ----------------------------------------------------------------
with aba6:
    st.subheader("🚫 1º Pilar - Controles Administrativos e Operacionais")
    st.caption("Registrador de infrações operacionais. Os critérios de 5 a 15 desclassificam o motorista diretamente. Já os critérios de 1 a 4 acumulam pontos (o acúmulo de mais de 129 pontos nestes critérios também gera a desclassificação automática).")

    with st.form("form_desclassificacoes", clear_on_submit=True):
        d_col1, d_col2, d_col3 = st.columns([2, 3, 1])
        with d_col1:
            mot_descl = st.selectbox("👤 Motorista", options=sorted(list(resumo_atualizado["MOTORISTA"].dropna().unique())))
        with d_col2:
            crit_descl = st.selectbox("📌 Critério / Infração", options=CRITERIOS_PILAR_1)
        with d_col3:
            pts_descl = st.number_input("📊 Quantidade de Pontos / Eventos", min_value=1, value=1)

        obs_descl = st.text_input("📝 Observação / Detalhes da Infração", placeholder="Ex: Ocorrência de freada brusca gravada pela telemetria, não uso do cinto...")
        
        btn_add_descl = st.form_submit_button("🚫 Registrar Infração/Desclassificação")

    if btn_add_descl:
        num_crit = int(crit_descl.split("-")[0].strip()) if "-" in crit_descl else 1
        eh_direto = num_crit >= 5
        tipo_imp = "DESCLASSIFICADO" if eh_direto else "PONTUAÇÃO"

        nova_descl = pd.DataFrame([{
            "MOTORISTA": mot_descl,
            "CRITERIO": crit_descl,
            "PONTOS": int(pts_descl),
            "TIPO_IMPACTO": tipo_imp,
            "OBSERVACAO": obs_descl
        }])
        st.session_state.df_desclassificacoes = pd.concat([st.session_state.df_desclassificacoes, nova_descl], ignore_index=True)
        st.success(f"Infração registrada para {mot_descl}!")
        st.rerun()

    st.markdown("#### Histórico de Desclassificações e Infrações Registradas")
    st.dataframe(st.session_state.df_desclassificacoes, use_container_width=True)

    if not st.session_state.df_desclassificacoes.empty:
        c_del1, c_del2 = st.columns([3, 1])
        with c_del1:
            idx_para_remover = st.selectbox("Selecione um registro para excluir", options=st.session_state.df_desclassificacoes.index)
        with c_del2:
            if st.button("🗑️ Excluir Registro"):
                st.session_state.df_desclassificacoes = st.session_state.df_desclassificacoes.drop(idx_para_remover).reset_index(drop=True)
                st.rerun()
