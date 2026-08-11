import importlib.util
import io
import os
import re
import subprocess
import sys
import unicodedata
from dataclasses import dataclass
from typing import Optional, Tuple

print("⏳ [1/5] Verificando e instalando bibliotecas...")
def instalar_dependencias():
    pacotes = ("pandas", "numpy", "openpyxl", "gradio", "requests")
    for pacote in pacotes:
        if importlib.util.find_spec(pacote) is None:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pacote])



import numpy as np
import pandas as pd
import requests
import gradio as gr

# ================================================================
# CONFIGURAÇÃO DE CAMINHOS E GOOGLE DRIVE (LINKS PÚBLICOS)
# ================================================================
@dataclass
class AppConfig:
    LINK_PRECOS: str = "https://docs.google.com/spreadsheets/d/1YZBLfxOgJinm1TJHYI49AEaOWVPmOEiA/edit?usp=sharing"
    LINK_FROTA: str = "https://docs.google.com/spreadsheets/d/1vX2JqzFLcyDxytrBP5vbHCvMRJsRrm6M/edit?usp=sharing"
    LINK_MOTORISTAS: str = "https://docs.google.com/spreadsheets/d/1YZBLfxOgJinm1TJHYI49AEaOWVPmOEiA/edit?usp=sharing"
    LINK_ABASTECIMENTOS: str = "https://docs.google.com/spreadsheets/d/1vX2JqzFLcyDxytrBP5vbHCvMRJsRrm6M/edit?usp=sharing"

    @staticmethod
    def extrair_url_export(url: str) -> str:
        """Converte o link de compartilhamento do Google Drive em um link de exportação XLSX."""
        if "/d/" in url:
            file_id = url.split("/d/")[1].split("/")[0]
            return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        return url


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
# LEITURA DE PLANILHAS VIA GOOGLE DRIVE (HTTP)
# ================================================================
class DataLoader:
    def __init__(self, config: AppConfig):
        self.config = config

    def _baixar_excel(self, url: str, sheet_name=0, header=0) -> pd.DataFrame:
        url_export = self.config.extrair_url_export(url)
        resp = requests.get(url_export)
        resp.raise_for_status()
        return pd.read_excel(io.BytesIO(resp.content), sheet_name=sheet_name, header=header)

def carregar_precos(self) -> pd.DataFrame:
        df = self._baixar_excel(self.config.LINK_PRECOS, sheet_name=0)
        
        col_tipo = DataUtils.encontrar_coluna(df, ["TIPO", "TIPO VEICULO", "TIPO DE VEICULO", "CATEGORIA"])
        col_media = DataUtils.encontrar_coluna(df, ["MEDIA", "MÉDIA", "MEDIA OBJETIVO"])
        col_premio = DataUtils.encontrar_coluna(df, ["TOTAL", "PREMIO", "PRÊMIO", "VALOR", "VALOR PREMIO"])

        # Trata estrutura sem cabeçalho padrão
        if None in (col_tipo, col_media, col_premio) and df.shape[1] >= 3:
            # Caso a planilha venha sem nome de colunas reconhecido, tenta mapear por posição
            df_temp = df.dropna(how="all").copy()
            if len(df_temp.columns) >= 5:
                df_temp = df_temp.iloc[1:].copy()
                df_temp.columns = ["DEL", "TIPO", "FATOR", "MEDIA", "PREMIO"] + list(df_temp.columns[5:])
                col_tipo, col_media, col_premio = "TIPO", "MEDIA", "PREMIO"
                df = df_temp

        # Validação de segurança para evitar crash silencioso
        if col_tipo is None or col_tipo not in df.columns:
            raise ValueError(
                f"Não foi possível localizar a coluna de 'TIPO' na planilha de Preços. "
                f"Colunas encontradas no arquivo baixado: {list(df.columns)}. "
                f"Verifique se o link do Google Drive está liberado como 'Qualquer pessoa com o link'."
            )

        resultado = pd.DataFrame({
            "TIPO": df[col_tipo].apply(DataUtils.normalizar_texto),
            "MEDIA": df[col_media].apply(DataUtils.converter_numero) if col_media else np.nan,
            "PREMIO": df[col_premio].apply(DataUtils.converter_numero) if col_premio else 0.0,
        }).dropna(subset=["MEDIA", "PREMIO"])

        resultado["TIPO"] = resultado["TIPO"].replace({"TOCO": "TRUCK"})
        return resultado[resultado["TIPO"] != ""].reset_index(drop=True)

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
        df = self._baixar_excel(self.config.LINK_ABASTECIMENTOS)

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
# CÁLCULO DE PREMIAÇÕES
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
# CONSTANTES DE CRITÉRIOS DE DESCLASSIFICAÇÃO (PILAR 1)
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

# ================================================================
# RECALCULO DE AUSÊNCIAS E DESCLASSIFICAÇÕES
# ================================================================
def aplicar_regras_gerais(df_resumo_original: pd.DataFrame, df_ausencias: pd.DataFrame, df_desclassificacoes: pd.DataFrame) -> pd.DataFrame:
    res = df_resumo_original.copy()

    motivos_iniciais = []
    for idx in res.index:
        st = res.at[idx, "STATUS_PREMIO"]
        if st == "DESCLASSIFICADO":
            motivos_iniciais.append("Média de consumo abaixo do limite mínimo da categoria")
        else:
            motivos_iniciais.append("Elegível / Em conformidade")
    res["MOTIVO_DESCLASSIFICACAO"] = motivos_iniciais

    # 1. Aplicar ausências (Férias / Atestados)
    if not df_ausencias.empty:
        soma_dias = df_ausencias.groupby("MOTORISTA")["DIAS"].sum().to_dict()
        res["DIAS_AUSENCIA"] = res["MOTORISTA"].map(soma_dias).fillna(0).astype(int)
        res["DIAS_EFETIVOS"] = np.maximum(0, 30 - res["DIAS_AUSENCIA"])
        res["PREMIO"] = res.apply(lambda r: max(0.0, r["PREMIO_BRUTO"] * (r["DIAS_EFETIVOS"] / 30.0)), axis=1)
    else:
        res["DIAS_AUSENCIA"] = 0
        res["DIAS_EFETIVOS"] = 30
        res["PREMIO"] = res["PREMIO_BRUTO"]

    # 2. Aplicar Desclassificações (Pilar 1)
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


# ================================================================
# GERADOR DE DATAFRAME EXCLUSIVO DO RH
# ================================================================
def gerar_tabela_rh(df_resumo: pd.DataFrame) -> pd.DataFrame:
    if df_resumo.empty:
        return pd.DataFrame(columns=["NOME", "FILIAL", "VALOR PAGO"])
    
    rh_df = pd.DataFrame()
    rh_df["NOME"] = df_resumo["MOTORISTA"]
    rh_df["FILIAL"] = df_resumo["BASE"].fillna("CIANORTE")
    rh_df["VALOR PAGO"] = df_resumo["PREMIO"].map(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    return rh_df


# ================================================================
# FILTRAGEM DO DASHBOARD
# ================================================================
def aplicar_filtros(motorista, placa, categoria, filial, df_resumo, df_eventos):
    res_f = df_resumo.copy()
    evt_f = df_eventos.copy()

    if motorista and motorista != "TODOS":
        m_norm = DataUtils.normalizar_texto(motorista)
        res_f = res_f[res_f["MOTORISTA"].apply(DataUtils.normalizar_texto).str.contains(m_norm, na=False)]
        evt_f = evt_f[evt_f["CONDUTOR_NORMALIZADO"].apply(DataUtils.normalizar_texto).str.contains(m_norm, na=False)]

    if placa and placa.strip():
        p_norm = DataUtils.padronizar_placa(placa)
        res_f = res_f[res_f["PLACAS"].apply(DataUtils.padronizar_placa).str.contains(p_norm, na=False)]
        evt_f = evt_f[evt_f["PLACA_PADRONIZADA"] == p_norm]

    if categoria and categoria != "TODAS":
        c_norm = DataUtils.normalizar_texto(categoria)
        res_f = res_f[res_f["CATEGORIA"] == c_norm]
        evt_f = evt_f[evt_f["TIPO_CALCULO"] == c_norm]

    if filial and filial != "TODAS":
        f_norm = DataUtils.normalizar_texto(filial)
        res_f = res_f[res_f["BASE"].apply(DataUtils.normalizar_texto) == f_norm]
        mots_da_filial = res_f["MOTORISTA"].apply(DataUtils.normalizar_texto).unique()
        evt_f = evt_f[evt_f["CONDUTOR_NORMALIZADO"].apply(DataUtils.normalizar_texto).isin(mots_da_filial)]

    tot_premio = res_f["PREMIO"].sum() if "PREMIO" in res_f.columns else 0.0
    tot_km = res_f["KM_TOTAL"].sum() if "KM_TOTAL" in res_f.columns else 0.0
    tot_litros = res_f["LITROS_TOTAL"].sum() if "LITROS_TOTAL" in res_f.columns else 0.0
    tot_gasto_combustivel = evt_f["VALOR_NUM"].sum() if "VALOR_NUM" in evt_f.columns else 0.0
    tot_media_geral = (tot_km / tot_litros) if tot_litros > 0 else 0.0
    tot_mots = len(res_f)

    res_view = res_f.copy()
    if "PREMIO" in res_view.columns:
        res_view["PREMIO"] = res_view["PREMIO"].map(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    if "MEDIA_CALCULADA" in res_view.columns:
        res_view["MEDIA_CALCULADA"] = res_view["MEDIA_CALCULADA"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "-")

    rh_view = gerar_tabela_rh(res_f)

    f_premio = f"R$ {tot_premio:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    f_gasto_comb = f"R$ {tot_gasto_combustivel:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    f_km = f"{tot_km:,.1f} km".replace(",", "X").replace(".", ",").replace("X", ".")
    f_litros = f"{tot_litros:,.1f} L".replace(",", "X").replace(".", ",").replace("X", ".")
    f_media = f"{tot_media_geral:.2f} km/L".replace(".", ",")
    f_mots = f"{tot_mots}"

    return f_premio, f_gasto_comb, f_km, f_litros, f_media, f_mots, res_view, rh_view, evt_f


# ================================================================
# GERADOR DE RECIBOS DE PREMIAÇÃO
# ================================================================
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
    <div class="recibo-card" style="background-color: #FFFFFF; padding: 28px; border-radius: 12px; max-width: 650px; margin: 0 auto; font-family: Arial, sans-serif; color: #000000; border: 1px solid #CBD5E1; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); page-break-after: always; break-after: page;">
        <div style="text-align: center; margin-bottom: 12px;">
            <div style="display: inline-block; width: 190px;">
                <div style="width: 190px; height: 46px; position: relative; border-radius: 2px; overflow: hidden; display: flex; flex-direction: column; border: 1px solid #CBD5E1;">
                    <div style="height: 33.3%; background-color: #0099DA;"></div>
                    <div style="height: 33.3%; background-color: #FFD700;"></div>
                    <div style="height: 33.3%; background-color: #1E2B7A;"></div>
                    <svg viewBox="0 0 24 24" style="position: absolute; top: 50%; left: 20px; transform: translateY(-50%); width: 24px; height: 24px;">
                        <path fill="#0099DA" stroke="#FFFFFF" stroke-width="1.5" d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z"/>
                        <path fill="#FFFFFF" d="M10 8.5a2.5 2.5 0 0 0 2.5 2.5 0.7 0.7 0 0 0-1.4 0z"/>
                    </svg>
                    <span style="position: absolute; top: 50%; left: 105px; transform: translate(-50%, -50%); font-family: 'Arial Black', sans-serif; font-size: 20px; font-weight: 900; color: #1E2B7A;">Ciapetro</span>
                </div>
            </div>
        </div>

        <h2 style="text-align: center; margin: 10px 0 16px 0; font-size: 18px; font-weight: bold; color: #000000;">Recibo de Premiação</h2>

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
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">OUTROS CONTROLES</td>
                <td style="padding: 4px 8px; text-align: center;">0</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">JORNADA</td>
                <td style="padding: 4px 8px; text-align: center;">0</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">EXCESSO DE VELOCIDADE</td>
                <td style="padding: 4px 8px; text-align: center;">0</td>
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
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">PONTOS NEG</td>
                <td style="padding: 4px 8px; text-align: center;">0</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">CONTROLES</td>
                <td style="padding: 4px 8px; text-align: center;">130</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">VALOR TOTAL CONTROLES</td>
                <td style="padding: 4px 8px; text-align: center;">R$ 0,00</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">% CONTROLES</td>
                <td style="padding: 4px 8px; text-align: center;">100%</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">R$ CONTROLES</td>
                <td style="padding: 4px 8px; text-align: center;">R$ 0,00</td>
            </tr>
            <tr style="border-bottom: 1px solid #000000;">
                <td style="background-color: #D0E0F0; padding: 4px 8px; border-right: 1px solid #000000; text-align: center;">VALOR MÉDIA</td>
                <td style="padding: 4px 8px; text-align: center;">{val_total_str}</td>
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

        <p style="text-align: center; margin-top: 18px; margin-bottom: 20px; font-size: 11px; font-weight: normal; line-height: 1.4; color: #000000;">
            Eu, <strong>{motorista_sel}</strong> ,Conferi e concordo com as informações, pois estão de acordo com a Política de Premiação dos Motoristas.
        </p>

        <div style="margin-top: 25px; text-align: left; font-size: 11px; color: #000000;">
            <span>Ass.: __________________________________________________</span><br/>
            <div style="margin-left: 35px; margin-top: 4px; font-weight: bold;">{motorista_sel}</div>
        </div>

        <div style="margin-top: 20px; text-align: left; font-size: 11px; color: #000000;">
            LOCAL/DATA _____________, ______/______/ 2026
        </div>

        <div style="margin-top: 12px; text-align: center; font-size: 11px; font-weight: bold; color: #000000;">
            Período de Controle: {periodo_ini} a {periodo_fim}
        </div>
    </div>
    """


def gerar_recibos_lote(filial_sel: str, motorista_sel: str, periodo_ini: str, periodo_fim: str, fator_c: str, df_resumo: pd.DataFrame) -> str:
    if not motorista_sel or motorista_sel in ("SELECIONE...", ""):
        return "<div style='text-align: center; padding: 40px; color: #64748B; font-size: 15px;'>👉 Por favor, selecione um motorista ou a opção de TODOS DA FILIAL para gerar os recibos.</div>"

    res_f = df_resumo.copy()

    if str(motorista_sel).startswith("TODOS"):
        if filial_sel and filial_sel != "TODAS":
            f_norm = DataUtils.normalizar_texto(filial_sel)
            res_f = res_f[res_f["BASE"].apply(DataUtils.normalizar_texto) == f_norm]

        lista_mots = sorted(list(res_f["MOTORISTA"].dropna().unique()))
        if not lista_mots:
            return f"<div style='text-align: center; padding: 40px; color: #EF4444;'>Nenhum motorista encontrado na filial '{filial_sel}'.</div>"
    else:
        lista_mots = [motorista_sel]

    recibos_html = []

    recibos_html.append(f"""
    <div style="background: #F8FAFC; border: 1px solid #E2E8F0; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center;">
        <span style="font-size: 14px; font-weight: bold; color: #1E293B;">
            📄 Total de Recibos Prontos: <span style="color: #2563EB;">{len(lista_mots)}</span>
        </span>
        <button onclick="window.print()" style="background-color: #2563EB; color: #FFFFFF; border: none; padding: 8px 18px; border-radius: 6px; font-weight: bold; cursor: pointer; font-size: 13px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            🖨️ Imprimir Todos os Recibos ({len(lista_mots)})
        </button>
    </div>
    """)

    for m_nome in lista_mots:
        row = df_resumo[df_resumo["MOTORISTA"] == m_nome]
        if not row.empty:
            card_html = gerar_html_unico_recibo(row.iloc[0], m_nome, periodo_ini, periodo_fim, fator_c)
            recibos_html.append(card_html)

    return "<div style='display: flex; flex-direction: column; gap: 30px;'>" + "".join(recibos_html) + "</div>"


def gerar_opcoes_exclusao_descl(df_descl: pd.DataFrame):
    if df_descl.empty:
        return gr.Dropdown(choices=["Nenhum registro para excluir"], value="Nenhum registro para excluir", interactive=False)
    opcoes = []
    for i, r in df_descl.iterrows():
        mot = r["MOTORISTA"]
        crit_curto = str(r["CRITERIO"]).split("-")[0].strip() if "-" in str(r["CRITERIO"]) else str(r["CRITERIO"])[:10]
        opcoes.append(f"[{i}] {mot} - Critério {crit_curto}")
    return gr.Dropdown(choices=opcoes, value=opcoes[0], interactive=True)


# ================================================================
# EXECUÇÃO PRINCIPAL E MONTAGEM DA INTERFACE
# ================================================================
print("⏳ [2/5] Carregando planilhas do Google Drive via HTTP...")
config = AppConfig()

loader = DataLoader(config)
engine = RewardEngine()

print("⏳ [3/5] Lendo e cruzando dados dos motoristas e frota...")
precos = loader.carregar_precos()
frota, mapa_frota = loader.carregar_frota()
cadastro = loader.carregar_cadastro_motoristas()
abastecimentos = loader.carregar_abastecimentos(mapa_frota)

print("⏳ [4/5] Processando consumo e calculando metas/premiações...")
eventos = engine.calcular_eventos_consumo(abastecimentos)
resumo_base = engine.calcular_premios(eventos, precos, cadastro)

# TABELAS EM MEMÓRIA (AUSÊNCIAS E DESCLASSIFICAÇÕES)
df_ausencias_global = pd.DataFrame(columns=["MOTORISTA", "TIPO_AUSENCIA", "DATA_INICIO", "DIAS", "OBSERVACAO"])
df_desclassificacoes_global = pd.DataFrame(columns=["MOTORISTA", "CRITERIO", "PONTOS", "TIPO_IMPACTO", "OBSERVACAO"])

print("⏳ [5/5] Construindo a interface executiva...")

mots_lista = ["TODOS"] + sorted(list(resumo_base["MOTORISTA"].dropna().unique()))
cats_lista = ["TODAS"] + sorted(list(resumo_base["CATEGORIA"].dropna().unique()))
filiais_lista = ["TODAS"] + sorted([str(b) for b in resumo_base["BASE"].dropna().unique() if str(b).strip() != ""])
mots_opcao = sorted(list(resumo_base["MOTORISTA"].dropna().unique()))
df_rh_inicial = gerar_tabela_rh(resumo_base)

# MONTAGEM VISUAL ROBUSTA COM GRADIO
with gr.Blocks(theme=gr.themes.Soft(), title="Dashboard do Prêmio de Motoristas") as app:

    state_resumo = gr.State(value=resumo_base)
    state_ausencias = gr.State(value=df_ausencias_global)
    state_desclassificacoes = gr.State(value=df_desclassificacoes_global)

    gr.HTML(
        """
        <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 16px; padding: 20px 28px; display: flex; align-items: center; gap: 24px; box-shadow: 0 2px 10px rgba(0, 0, 0, 0.03); margin-bottom: 20px;">
            <div style="display: flex; flex-direction: column; align-items: center; width: 95px; flex-shrink: 0;">
                <div style="width: 85px; height: 42px; position: relative; border-radius: 2px; overflow: hidden; display: flex; flex-direction: column; border: 1px solid #CBD5E1;">
                    <div style="height: 33.3%; background-color: #0099DA;"></div>
                    <div style="height: 33.3%; background-color: #FFD700;"></div>
                    <div style="height: 33.3%; background-color: #1E2B7A;"></div>
                    <svg viewBox="0 0 24 24" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 24px; height: 24px;">
                        <path fill="#0099DA" stroke="#FFFFFF" stroke-width="1.5" d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z"/>
                        <path fill="#FFFFFF" d="M10 8.5a2.5 2.5 0 0 0 2.5 2.5 0.7 0.7 0 0 0-1.4 0z"/>
                    </svg>
                </div>
                <span style="font-family: 'Arial Black', 'Helvetica Neue', sans-serif; font-weight: 900; color: #1E2B7A; font-size: 13px; margin-top: 3px; letter-spacing: -0.3px;">Ciapetro</span>
                <span style="font-family: sans-serif; font-size: 5px; color: #475569; text-align: center; line-height: 1; margin-top: 0px; text-transform: uppercase;">Distribuidora de Combustíveis</span>
            </div>

            <div style="display: flex; flex-direction: column; justify-content: center;">
                <h1 style="margin: 0; color: #0F172A; font-size: 26px; font-weight: 800; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; letter-spacing: -0.5px;">
                    Dashboard do Prêmio de Motoristas
                </h1>
                <p style="margin: 4px 0 0 0; color: #64748B; font-size: 14px; font-weight: 400; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
                    Visão gerencial de consumo, desempenho e prêmio
                </p>
            </div>
        </div>
        """
    )

    with gr.Row():
        f_mot = gr.Dropdown(choices=mots_lista, value="TODOS", label="👤 Filtrar Motorista", filterable=True)
        f_plc = gr.Textbox(label="🔍 Filtrar Placa", placeholder="Ex: ABC1234 ou ABC1D23")
        f_cat = gr.Dropdown(choices=cats_lista, value="TODAS", label="🏷️ Categoria de Veículo")
        f_fil = gr.Dropdown(choices=filiais_lista, value="TODAS", label="🏢 Filtrar Filial / Base", filterable=True)

    with gr.Row():
        btn_aplicar = gr.Button("⚡ Aplicar Filtros", variant="primary")
        btn_limpar = gr.Button("🔄 Limpar", variant="secondary")

    gr.Markdown("---")

    with gr.Row():
        kpi_p = gr.Textbox(label="💰 Total em Prêmios", value="R$ 0,00", interactive=False)
        kpi_gasto_comb = gr.Textbox(label="💳 Total Gasto Combustível", value="R$ 0,00", interactive=False)
        kpi_k = gr.Textbox(label="RODADO TOTAL (KM)", value="0 km", interactive=False)
        kpi_l = gr.Textbox(label="⛽ COMBUSTÍVEL TOTAL", value="0 L", interactive=False)
        kpi_avg = gr.Textbox(label="🎯 MÉDIA GERAL (KM/L)", value="0,00 km/L", interactive=False)
        kpi_m = gr.Textbox(label="👥 MOTORISTAS NA LISTA", value="0", interactive=False)

    gr.Markdown("---")

    with gr.Tabs():
        with gr.Tab("📊 Resumo de Premiações por Motorista"):
            grid_resumo = gr.Dataframe(value=resumo_base, interactive=False)

        with gr.Tab("👔 Relatório RH - Lançamento de Pagamento"):
            gr.Markdown("### 👔 **Relatório RH - Pagamento de Prêmios**")
            gr.Markdown("Relação simplificada de motoristas, filiais e valores devidos para lançamento na folha de pagamento pelo setor de RH.")
            grid_rh = gr.Dataframe(value=df_rh_inicial, interactive=False)

        with gr.Tab("📋 Detalhamento dos Abastecimentos"):
            grid_eventos = gr.Dataframe(value=eventos, interactive=False)

        with gr.Tab("📄 Recibo de Premiação"):
            gr.Markdown("### 🖨️ **Gerador de Recibo de Premiação para Assinatura**")

            opcoes_recibo_inic = ["SELECIONE...", "TODOS OS MOTORISTAS (TODAS AS FILIAIS)"] + sorted(list(resumo_base["MOTORISTA"].dropna().unique()))

            with gr.Row():
                rec_fil = gr.Dropdown(choices=filiais_lista, value="TODAS", label="🏢 Filtrar Filial / Base", filterable=True)
                rec_mot = gr.Dropdown(choices=opcoes_recibo_inic, value="TODOS OS MOTORISTAS (TODAS AS FILIAIS)", label="👤 Selecionar Motorista", filterable=True)
                rec_ini = gr.Textbox(label="📅 Período Início", value="26/06/2026")
                rec_fim = gr.Textbox(label="📅 Período Fim", value="25/07/2026")

            with gr.Row():
                rec_fator = gr.Textbox(label="⚖️ Fator Carga", value="50%")
                btn_recibo = gr.Button("📄 Gerar / Atualizar Recibo(s)", variant="primary")

            recibo_output = gr.HTML(
                value="<div style='text-align: center; padding: 40px; color: #64748B;'>👉 Selecione um motorista ou uma filial acima para gerar os recibos.</div>"
            )

        with gr.Tab("🏥 Lançamento de Atestados e Férias"):
            gr.Markdown("### 🏥 **Lançamento de Ausências (Desconto de Atestados / Férias)**")
            gr.Markdown("Os dias lançados aqui serão automaticamente descontados da quantidade de **dias efetivos** e do **valor final do prêmio** no dashboard e nos recibos.")

            with gr.Row():
                aus_mot = gr.Dropdown(choices=mots_opcao, label="👤 Motorista", filterable=True)
                aus_tipo = gr.Radio(choices=["Atestado Médico", "Férias", "Outro Afastamento"], value="Atestado Médico", label="📌 Tipo de Ausência")
                aus_data = gr.Textbox(label="📅 Data de Início", value="01/07/2026", placeholder="DD/MM/AAAA")
                aus_dias = gr.Number(label="🔢 Dias Ausente", value=1, precision=0)

            aus_obs = gr.Textbox(label="📝 Observação / Motivo", placeholder="Ex: CID 10, Licença médica, Férias regulamentares...")

            with gr.Row():
                btn_add_ausencia = gr.Button("➕ Lançar Ausência", variant="primary")
                btn_limpar_ausencias = gr.Button("🗑️ Limpar Todos os Lançamentos", variant="stop")

            gr.Markdown("#### 📋 **Histórico de Ausências Lançadas**")
            grid_ausencias = gr.Dataframe(value=df_ausencias_global, interactive=False)

        with gr.Tab("🚫 Gestão de Desclassificações (Pilar 1)"):
            gr.Markdown("### 🚫 **1º Pilar - Controles Administrativos e Operacionais**")
            gr.Markdown("Registrar infrações operacionais. Os critérios de **5 a 15** desclassificam o motorista diretamente. Já os critérios de **1 a 4** acumulam pontos (o acúmulo de **mais de 129 pontos** nestes critérios também gera a desclassificação automática com prêmio zerado em R$ 0,00).")

            with gr.Row():
                descl_mot = gr.Dropdown(choices=mots_opcao, label="👤 Motorista", filterable=True)
                descl_crit = gr.Dropdown(choices=CRITERIOS_PILAR_1, value=CRITERIOS_PILAR_1[0], label="📌 Critério / Infração")
                descl_pontos = gr.Number(label="🔢 Quantidade de Pontos / Eventos", value=1, precision=0)

            descl_obs = gr.Textbox(label="📝 Observação / Detalhes da Infração", placeholder="Ex: Ocorrência de freada brusca gravada pela telemetria, não uso do cinto...")

            with gr.Row():
                btn_add_desclassificacao = gr.Button("🚫 Registrar Infração / Desclassificação", variant="primary")

            gr.Markdown("---")
            gr.Markdown("#### 📋 **Histórico de Desclassificações e Infrações Registradas**")
            grid_desclassificacoes = gr.Dataframe(value=df_desclassificacoes_global, interactive=False)

            with gr.Row():
                descl_excluir_sel = gr.Dropdown(choices=["Nenhum registro para excluir"], value="Nenhum registro para excluir", label="🗑️ Selecionar Registro para Excluir", filterable=True, interactive=False)
                btn_excluir_item_descl = gr.Button("🗑️ Excluir Registro Selecionado", variant="stop")
                btn_limpar_desclassificacoes = gr.Button("⚠️ Limpar TODOS os Lançamentos", variant="secondary")


    # ================================================================
    # FUNÇÕES DE EVENTOS DO GRADIO
    # ================================================================

    # 1. ATUALIZAR DASHBOARD COMPLETO
    def atualizar_dashboard(mot, plc, cat, fil, res_atual):
        return aplicar_filtros(mot, plc, cat, fil, res_atual, eventos)

    inputs_filtros = [f_mot, f_plc, f_cat, f_fil, state_resumo]
    outputs_filtros = [kpi_p, kpi_gasto_comb, kpi_k, kpi_l, kpi_avg, kpi_m, grid_resumo, grid_rh, grid_eventos]

    f_mot.change(fn=atualizar_dashboard, inputs=inputs_filtros, outputs=outputs_filtros)
    f_plc.change(fn=atualizar_dashboard, inputs=inputs_filtros, outputs=outputs_filtros)
    f_cat.change(fn=atualizar_dashboard, inputs=inputs_filtros, outputs=outputs_filtros)
    f_fil.change(fn=atualizar_dashboard, inputs=inputs_filtros, outputs=outputs_filtros)
    btn_aplicar.click(fn=atualizar_dashboard, inputs=inputs_filtros, outputs=outputs_filtros)

    btn_limpar.click(
        fn=lambda: ("TODOS", "", "TODAS", "TODAS"),
        outputs=[f_mot, f_plc, f_cat, f_fil]
    ).then(
        fn=atualizar_dashboard,
        inputs=inputs_filtros,
        outputs=outputs_filtros
    )

    # 2. FILTRAR MOTORISTAS DA GUIA DE RECIBO POR FILIAL
    def atualizar_motoristas_recibo(filial, res_df):
        if not filial or filial == "TODAS":
            mots_f = sorted(list(res_df["MOTORISTA"].dropna().unique()))
            opt_todos = "TODOS OS MOTORISTAS (TODAS AS FILIAIS)"
        else:
            f_norm = DataUtils.normalizar_texto(filial)
            mots_f = sorted(list(res_df[res_df["BASE"].apply(DataUtils.normalizar_texto) == f_norm]["MOTORISTA"].dropna().unique()))
            opt_todos = f"TODOS OS MOTORISTAS DA FILIAL ({filial})"

        mots = ["SELECIONE...", opt_todos] + mots_f
        return gr.Dropdown(choices=mots, value=opt_todos)

    rec_fil.change(
        fn=atualizar_motoristas_recibo,
        inputs=[rec_fil, state_resumo],
        outputs=rec_mot
    )

    # 3. GERAR RECIBO(S)
    recibo_inputs = [rec_fil, rec_mot, rec_ini, rec_fim, rec_fator, state_resumo]

    btn_recibo.click(fn=gerar_recibos_lote, inputs=recibo_inputs, outputs=recibo_output)
    rec_mot.change(fn=gerar_recibos_lote, inputs=recibo_inputs, outputs=recibo_output)

    # 4. AÇÕES DE AUSÊNCIAS
    def adicionar_ausencia(mot, tipo, dt, dias, obs, df_aus_atual, df_descl_atual, m_filtro, p_filtro, c_filtro, fil_filtro, r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft):
        if not mot:
            kpi_vals = aplicar_filtros(m_filtro, p_filtro, c_filtro, fil_filtro, resumo_base, eventos)
            rec_html = gerar_recibos_lote(r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft, resumo_base)
            return df_aus_atual, resumo_base, df_aus_atual, *kpi_vals, rec_html

        novo = pd.DataFrame([{
            "MOTORISTA": mot,
            "TIPO_AUSENCIA": tipo,
            "DATA_INICIO": dt,
            "DIAS": int(dias if dias else 1),
            "OBSERVACAO": obs
        }])

        df_aus_novo = pd.concat([df_aus_atual, novo], ignore_index=True)
        res_calculado = aplicar_regras_gerais(resumo_base, df_aus_novo, df_descl_atual)

        kpi_vals = aplicar_filtros(m_filtro, p_filtro, c_filtro, fil_filtro, res_calculado, eventos)
        recibo_html = gerar_recibos_lote(r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft, res_calculado)

        return df_aus_novo, res_calculado, df_aus_novo, *kpi_vals, recibo_html

    def resetar_ausencias(df_descl_atual, m_filtro, p_filtro, c_filtro, fil_filtro, r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft):
        df_vazio = pd.DataFrame(columns=["MOTORISTA", "TIPO_AUSENCIA", "DATA_INICIO", "DIAS", "OBSERVACAO"])
        res_calculado = aplicar_regras_gerais(resumo_base, df_vazio, df_descl_atual)

        kpi_vals = aplicar_filtros(m_filtro, p_filtro, c_filtro, fil_filtro, res_calculado, eventos)
        recibo_html = gerar_recibos_lote(r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft, res_calculado)

        return df_vazio, res_calculado, df_vazio, *kpi_vals, recibo_html

    out_regras_aus = [state_ausencias, state_resumo, grid_ausencias, kpi_p, kpi_gasto_comb, kpi_k, kpi_l, kpi_avg, kpi_m, grid_resumo, grid_rh, grid_eventos, recibo_output]

    btn_add_ausencia.click(
        fn=adicionar_ausencia,
        inputs=[aus_mot, aus_tipo, aus_data, aus_dias, aus_obs, state_ausencias, state_desclassificacoes, f_mot, f_plc, f_cat, f_fil, rec_fil, rec_mot, rec_ini, rec_fim, rec_fator],
        outputs=out_regras_aus
    )

    btn_limpar_ausencias.click(
        fn=resetar_ausencias,
        inputs=[state_desclassificacoes, f_mot, f_plc, f_cat, f_fil, rec_fil, rec_mot, rec_ini, rec_fim, rec_fator],
        outputs=out_regras_aus
    )

    # 5. AÇÕES DE DESCLASSIFICAÇÃO (INCLUSÃO, EXCLUSÃO INDIVIDUAL E LIMPEZA)
    def adicionar_desclassificacao(mot, crit, pontos, obs, df_aus_atual, df_descl_atual, m_filtro, p_filtro, c_filtro, fil_filtro, r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft):
        if not mot:
            kpi_vals = aplicar_filtros(m_filtro, p_filtro, c_filtro, fil_filtro, resumo_base, eventos)
            rec_html = gerar_recibos_lote(r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft, resumo_base)
            opt_excl = gerar_opcoes_exclusao_descl(df_descl_atual)
            return df_descl_atual, resumo_base, df_descl_atual, opt_excl, *kpi_vals, rec_html

        num_crit = int(crit.split("-")[0].strip()) if "-" in crit else 1
        eh_desclassificacao_direta = num_crit >= 5

        tipo_impacto = "DESCLASSIFICADO" if eh_desclassificacao_direta else "PONTUAÇÃO"

        novo = pd.DataFrame([{
            "MOTORISTA": mot,
            "CRITERIO": crit,
            "PONTOS": int(pontos if pontos else 1),
            "TIPO_IMPACTO": tipo_impacto,
            "OBSERVACAO": obs
        }])

        df_descl_novo = pd.concat([df_descl_atual, novo], ignore_index=True)
        res_calculado = aplicar_regras_gerais(resumo_base, df_aus_atual, df_descl_novo)

        kpi_vals = aplicar_filtros(m_filtro, p_filtro, c_filtro, fil_filtro, res_calculado, eventos)
        recibo_html = gerar_recibos_lote(r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft, res_calculado)
        opt_excl = gerar_opcoes_exclusao_descl(df_descl_novo)

        return df_descl_novo, res_calculado, df_descl_novo, opt_excl, *kpi_vals, recibo_html

    def excluir_item_desclassificacao(item_sel, df_aus_atual, df_descl_atual, m_filtro, p_filtro, c_filtro, fil_filtro, r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft):
        if not item_sel or "Nenhum" in item_sel or df_descl_atual.empty:
            kpi_vals = aplicar_filtros(m_filtro, p_filtro, c_filtro, fil_filtro, resumo_base, eventos)
            rec_html = gerar_recibos_lote(r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft, resumo_base)
            opt_excl = gerar_opcoes_exclusao_descl(df_descl_atual)
            return df_descl_atual, resumo_base, df_descl_atual, opt_excl, *kpi_vals, rec_html

        match = re.search(r"^\[(\d+)\]", item_sel)
        if match:
            idx_excluir = int(match.group(1))
            if idx_excluir in df_descl_atual.index:
                df_descl_novo = df_descl_atual.drop(idx_excluir).reset_index(drop=True)
            else:
                df_descl_novo = df_descl_atual
        else:
            df_descl_novo = df_descl_atual

        res_calculado = aplicar_regras_gerais(resumo_base, df_aus_atual, df_descl_novo)
        kpi_vals = aplicar_filtros(m_filtro, p_filtro, c_filtro, fil_filtro, res_calculado, eventos)
        recibo_html = gerar_recibos_lote(r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft, res_calculado)
        opt_excl = gerar_opcoes_exclusao_descl(df_descl_novo)

        return df_descl_novo, res_calculado, df_descl_novo, opt_excl, *kpi_vals, rec_html

    def resetar_desclassificacoes(df_aus_atual, m_filtro, p_filtro, c_filtro, fil_filtro, r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft):
        df_vazio = pd.DataFrame(columns=["MOTORISTA", "CRITERIO", "PONTOS", "TIPO_IMPACTO", "OBSERVACAO"])
        res_calculado = aplicar_regras_gerais(resumo_base, df_aus_atual, df_vazio)

        kpi_vals = aplicar_filtros(m_filtro, p_filtro, c_filtro, fil_filtro, res_calculado, eventos)
        recibo_html = gerar_recibos_lote(r_fil_sel, r_mot_sel, r_ini, r_fim, r_ft, res_calculado)
        opt_excl = gerar_opcoes_exclusao_descl(df_vazio)

        return df_vazio, res_calculado, df_vazio, opt_excl, *kpi_vals, recibo_html

    out_descl = [state_desclassificacoes, state_resumo, grid_desclassificacoes, descl_excluir_sel, kpi_p, kpi_gasto_comb, kpi_k, kpi_l, kpi_avg, kpi_m, grid_resumo, grid_rh, grid_eventos, recibo_output]

    btn_add_desclassificacao.click(
        fn=adicionar_desclassificacao,
        inputs=[descl_mot, descl_crit, descl_pontos, descl_obs, state_ausencias, state_desclassificacoes, f_mot, f_plc, f_cat, f_fil, rec_fil, rec_mot, rec_ini, rec_fim, rec_fator],
        outputs=out_descl
    )

    btn_excluir_item_descl.click(
        fn=excluir_item_desclassificacao,
        inputs=[descl_excluir_sel, state_ausencias, state_desclassificacoes, f_mot, f_plc, f_cat, f_fil, rec_fil, rec_mot, rec_ini, rec_fim, rec_fator],
        outputs=out_descl
    )

    btn_limpar_desclassificacoes.click(
        fn=resetar_desclassificacoes,
        inputs=[state_ausencias, f_mot, f_plc, f_cat, f_fil, rec_fil, rec_mot, rec_ini, rec_fim, rec_fator],
        outputs=out_descl
    )

print("✅ Tudo pronto! Abrindo o painel...")
app.launch(inline=True, share=True)
