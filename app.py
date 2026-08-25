import os
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime, date
from typing import Optional, Tuple

import streamlit as st
import numpy as np
import openpyxl
import pandas as pd


# ================================================================
# CONFIGURAÇÃO DE CAMINHOS E PERSISTÊNCIA (LOCAL / RAILWAY)
# ================================================================
DATA_DIR = os.getenv("DATA_DIR", ".")
ARQUIVO_AUSENCIAS = os.path.join(DATA_DIR, "ausencias.csv")
ARQUIVO_DESCLASSIFICACOES = os.path.join(DATA_DIR, "desclassificacoes.csv")
ARQUIVO_CATEGORIAS_CUSTOM = os.path.join(DATA_DIR, "categorias_customizadas.csv")
ARQUIVO_FROTA_CUSTOM = os.path.join(DATA_DIR, "frota_customizada.csv")
ARQUIVO_MOTORISTAS_CUSTOM = os.path.join(DATA_DIR, "motoristas_customizados.csv")
ARQUIVO_INATIVOS = os.path.join(DATA_DIR, "inativos.csv")


def garantir_diretorio():
  """Garante a existência do diretório de dados antes de operações de escrita."""
  if DATA_DIR and DATA_DIR != "." and not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR, exist_ok=True)


def carregar_ausencias() -> pd.DataFrame:
  """Carrega as ausências salvas em disco ou retorna DataFrame vazio."""
  if os.path.exists(ARQUIVO_AUSENCIAS):
    try:
      df = pd.read_csv(ARQUIVO_AUSENCIAS, dtype=str, encoding="utf-8-sig")
      if "DATA_FIM" not in df.columns:
        df["DATA_FIM"] = ""
      if "DIAS" in df.columns:
        df["DIAS"] = pd.to_numeric(df["DIAS"], errors="coerce").fillna(1)
      else:
        df["DIAS"] = 1
      colunas = ["MOTORISTA", "TIPO_AUSENCIA", "DATA_INICIO", "DATA_FIM", "DIAS", "OBSERVACAO"]
      for coluna in colunas:
        if coluna not in df.columns:
          df[coluna] = ""
      return df[colunas]
    except Exception as e:
      print(f"Erro ao carregar ausências: {e}")
  return pd.DataFrame(
      columns=["MOTORISTA", "TIPO_AUSENCIA", "DATA_INICIO", "DIAS", "OBSERVACAO"]
  )


def salvar_ausencias(df: pd.DataFrame):
  """Salva as ausências no arquivo CSV com codificação UTF-8."""
  try:
    garantir_diretorio()
    df.to_csv(ARQUIVO_AUSENCIAS, index=False, encoding="utf-8-sig")
  except Exception as e:
    print(f"Erro ao salvar ausências: {e}")


def carregar_desclassificacoes() -> pd.DataFrame:
  """Carrega as desclassificações salvas em disco ou retorna DataFrame vazio."""
  if os.path.exists(ARQUIVO_DESCLASSIFICACOES):
    try:
      df = pd.read_csv(
          ARQUIVO_DESCLASSIFICACOES, dtype=str, encoding="utf-8-sig"
      )
      if "PONTOS" in df.columns:
        df["PONTOS"] = pd.to_numeric(df["PONTOS"], errors="coerce").fillna(1)
      return df
    except Exception as e:
      print(f"Erro ao carregar desclassificações: {e}")
  return pd.DataFrame(
      columns=["MOTORISTA", "CRITERIO", "PONTOS", "TIPO_IMPACTO", "OBSERVACAO"]
  )


def salvar_desclassificacoes(df: pd.DataFrame):
  """Salva as desclassificações no arquivo CSV com codificação UTF-8."""
  try:
    garantir_diretorio()
    df.to_csv(ARQUIVO_DESCLASSIFICACOES, index=False, encoding="utf-8-sig")
  except Exception as e:
    print(f"Erro ao salvar desclassificações: {e}")


def normalizar_chave_categoria_customizada(motorista: str, placa: str = "") -> str:
  """Gera chave persistente para categoria manual por motorista + placa."""
  mot = DataUtils.normalizar_texto(motorista)
  plc = DataUtils.padronizar_placa(placa)
  return f"{mot}|||{plc}" if plc else mot


def carregar_categorias_customizadas() -> dict:
  """Carrega os mapeamentos manuais de categoria.

  Compatibilidade:
  - MOTORISTA_CHAVE = MOTORISTA|||PLACA: regra por motorista e placa.
  - MOTORISTA_CHAVE = MOTORISTA: regra antiga por motorista.
  """
  if os.path.exists(ARQUIVO_CATEGORIAS_CUSTOM):
    try:
      df = pd.read_csv(
          ARQUIVO_CATEGORIAS_CUSTOM, dtype=str, encoding="utf-8-sig"
      )
      if "MOTORISTA_CHAVE" in df.columns and "CATEGORIA_ESCOLHIDA" in df.columns:
        df = df.fillna("")
        return {
            str(k).strip().upper(): DataUtils.normalizar_texto(v)
            for k, v in zip(df["MOTORISTA_CHAVE"], df["CATEGORIA_ESCOLHIDA"])
            if str(k).strip() and str(v).strip()
        }
    except Exception as e:
      print(f"Erro ao carregar categorias customizadas: {e}")
  return {}


def salvar_categorias_customizadas(mapa: dict):
  """Salva as seleções manuais de categorias por motorista e placa."""
  try:
    garantir_diretorio()
    df = pd.DataFrame(
        [
            [str(k).strip().upper(), DataUtils.normalizar_texto(v)]
            for k, v in (mapa or {}).items()
            if str(k).strip() and str(v).strip()
        ],
        columns=["MOTORISTA_CHAVE", "CATEGORIA_ESCOLHIDA"],
    )
    df.to_csv(ARQUIVO_CATEGORIAS_CUSTOM, index=False, encoding="utf-8-sig")
  except Exception as e:
    print(f"Erro ao salvar categorias customizadas: {e}")


def carregar_frota_customizada() -> pd.DataFrame:
  if os.path.exists(ARQUIVO_FROTA_CUSTOM):
    try:
      df = pd.read_csv(ARQUIVO_FROTA_CUSTOM, dtype=str, encoding="utf-8-sig")
      if "CAVALO" in df.columns and "TIPO" in df.columns:
        return df
    except Exception as e:
      print(f"Erro ao carregar frota customizada: {e}")
  return pd.DataFrame(columns=["CAVALO", "TIPO"])


def salvar_frota_customizada(df: pd.DataFrame):
  try:
    garantir_diretorio()
    df.to_csv(ARQUIVO_FROTA_CUSTOM, index=False, encoding="utf-8-sig")
  except Exception as e:
    print(f"Erro ao salvar frota customizada: {e}")


def carregar_motoristas_customizados() -> pd.DataFrame:
  if os.path.exists(ARQUIVO_MOTORISTAS_CUSTOM):
    try:
      df = pd.read_csv(
          ARQUIVO_MOTORISTAS_CUSTOM, dtype=str, encoding="utf-8-sig"
      )
      if (
          "MOTORISTAS" in df.columns
          and "TIPO" in df.columns
          and "BASE" in df.columns
      ):
        return df
    except Exception as e:
      print(f"Erro ao carregar motoristas customizados: {e}")
  return pd.DataFrame(columns=["MOTORISTAS", "TIPO", "BASE"])


def salvar_motoristas_customizados(df: pd.DataFrame):
  try:
    garantir_diretorio()
    df.to_csv(ARQUIVO_MOTORISTAS_CUSTOM, index=False, encoding="utf-8-sig")
  except Exception as e:
    print(f"Erro ao salvar motoristas customizados: {e}")


def carregar_inativos() -> dict:
  """Carrega as placas e motoristas marcados como INATIVOS do arquivo CSV."""
  if os.path.exists(ARQUIVO_INATIVOS):
    try:
      df = pd.read_csv(ARQUIVO_INATIVOS, dtype=str, encoding="utf-8-sig")
      if "DATA_INATIVACAO" not in df.columns:
        df["DATA_INATIVACAO"] = ""

      df_mot = df[df["TIPO"] == "MOTORISTA"].dropna(subset=["VALOR"])
      df_placa = df[df["TIPO"] == "PLACA"].dropna(subset=["VALOR"])

      mots = dict(zip(df_mot["VALOR"], df_mot["DATA_INATIVACAO"].fillna("")))
      placas = dict(
          zip(df_placa["VALOR"], df_placa["DATA_INATIVACAO"].fillna(""))
      )
      return {"MOTORISTA": mots, "PLACA": placas}
    except Exception as e:
      print(f"Erro ao carregar inativos: {e}")
  return {"MOTORISTA": {}, "PLACA": {}}


def alternar_inativo(tipo: str, valor: str, inativar: bool = True):
  """Grava ou remove o motorista/placa do arquivo de inativos."""
  garantir_diretorio()
  if tipo == "MOTORISTA":
    valor = DataUtils.normalizar_texto(valor)
  else:
    valor = DataUtils.padronizar_placa(valor)

  df = pd.DataFrame(columns=["TIPO", "VALOR", "DATA_INATIVACAO"])
  if os.path.exists(ARQUIVO_INATIVOS):
    try:
      df = pd.read_csv(ARQUIVO_INATIVOS, dtype=str, encoding="utf-8-sig")
      if "DATA_INATIVACAO" not in df.columns:
        df["DATA_INATIVACAO"] = ""
    except Exception:
      pass

  mask = (df["TIPO"] == tipo) & (df["VALOR"] == valor)
  if inativar:
    if not mask.any():
      data_atual = datetime.now().strftime("%d/%m/%Y")
      novo_reg = pd.DataFrame(
          [{"TIPO": tipo, "VALOR": valor, "DATA_INATIVACAO": data_atual}]
      )
      df = pd.concat([df, novo_reg], ignore_index=True)
  else:
    df = df[~mask]

  try:
    df.to_csv(ARQUIVO_INATIVOS, index=False, encoding="utf-8-sig")
  except Exception as e:
    print(f"Erro ao salvar inativos: {e}")


@dataclass
class AppConfig:
  CAMINHO_PRECOS: str = "Pasta2.xlsx"
  CAMINHO_FROTA: str = "frota.xlsx"
  CAMINHO_MOTORISTAS: str = "Pasta4.xlsx"
  CAMINHO_ABASTECIMENTOS: str = "uah_abastecimentos_3.xlsx"

  def _resolver_caminho_real(self, nome_arquivo: str) -> str:
    if os.path.isfile(nome_arquivo):
      return nome_arquivo

    if os.path.exists("."):
      for arquivo_real in os.listdir("."):
        if arquivo_real.lower() == nome_arquivo.lower():
          return arquivo_real
    return nome_arquivo

  def criar_arquivos_teste_se_ausentes(self):
    if not os.path.isfile(self.CAMINHO_PRECOS):
      df_p = pd.DataFrame({
          "TIPO": ["TRUCK", "BITRUCK", "BITRUCK UBERABA", "CARRETA"],
          "MEDIA": [2.5, 2.2, 2.1, 1.8],
          "PREMIO": [500.0, 600.0, 650.0, 700.0],
      })
      df_p.to_excel(self.CAMINHO_PRECOS, index=False)

    if not os.path.isfile(self.CAMINHO_FROTA):
      df_f = pd.DataFrame({
          "CAVALO": [
              "ABC1234",
              "DEF5678",
              "GHI9012",
              "WES0001",
              "WES0002",
              "WES0003",
              "WES0004",
          ],
          "TIPO": [
              "TRUCK",
              "BITRUCK",
              "CARRETA",
              "TRUCK",
              "BITRUCK",
              "CARRETA",
              "BITRUCK UBERABA",
          ],
      })
      df_f.to_excel(self.CAMINHO_FROTA, index=False)

    if not os.path.isfile(self.CAMINHO_MOTORISTAS):
      df_m = pd.DataFrame([
          ["MOTORISTAS", "TIPO", "BASE"],
          ["JOAO SILVA", "TRUCK", "CIANORTE"],
          ["MARIA SOUZA", "BITRUCK", "UBERABA"],
          ["CARLOS ALVES", "CARRETA", "MARINGA"],
          ["WESLEI", "MULTIPLASCATEGORIAS", "CIANORTE"],
      ])
      df_m.to_excel(self.CAMINHO_MOTORISTAS, header=False, index=False)

    if not os.path.isfile(self.CAMINHO_ABASTECIMENTOS):
      df_a = pd.DataFrame({
          "DATA": [
              "01/08/2026",
              "05/08/2026",
              "10/08/2026",
              "15/08/2026",
              "01/08/2026",
              "05/08/2026",
              "10/08/2026",
              "15/08/2026",
          ],
          "PLACA": [
              "ABC1234",
              "ABC1234",
              "DEF5678",
              "DEF5678",
              "WES0001",
              "WES0002",
              "WES0003",
              "WES0004",
          ],
          "MOTORISTA": [
              "JOAO SILVA",
              "JOAO SILVA",
              "MARIA SOUZA",
              "MARIA SOUZA",
              "WESLEI",
              "WESLEI",
              "WESLEI",
              "WESLEI",
          ],
          "KM": [1000, 1500, 2000, 2600, 100, 500, 1000, 1600],
          "QTDE": [200, 180, 280, 270, 50, 100, 150, 200],
          "VALOR": [
              1200.0,
              1080.0,
              1680.0,
              1620.0,
              300.0,
              600.0,
              900.0,
              1200.0,
          ],
      })
      df_a.to_excel(self.CAMINHO_ABASTECIMENTOS, index=False)

  def verificar_arquivos(self):
    self.criar_arquivos_teste_se_ausentes()
    self.CAMINHO_PRECOS = self._resolver_caminho_real(self.CAMINHO_PRECOS)
    self.CAMINHO_FROTA = self._resolver_caminho_real(self.CAMINHO_FROTA)
    self.CAMINHO_MOTORISTAS = self._resolver_caminho_real(
        self.CAMINHO_MOTORISTAS
    )
    self.CAMINHO_ABASTECIMENTOS = self._resolver_caminho_real(
        self.CAMINHO_ABASTECIMENTOS
    )


# ================================================================
# UTILITÁRIOS E TRATAMENTO DE TEXTO / PLACAS / DATAS
# ================================================================
class DataUtils:

  NUMERO_PARA_LETRA = {
      "0": "A",
      "1": "B",
      "2": "C",
      "3": "D",
      "4": "E",
      "5": "F",
      "6": "G",
      "7": "H",
      "8": "I",
      "9": "J",
  }
  LETRA_PARA_NUMERO = {v: k for k, v in NUMERO_PARA_LETRA.items()}

  @staticmethod
  def normalizar_texto(valor) -> str:
    if pd.isna(valor):
      return ""
    texto = str(valor).strip().upper()
    texto = (
        unicodedata.normalize("NFKD", texto)
        .encode("ASCII", "ignore")
        .decode("ASCII")
    )
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
  def converter_data(valor):
    if pd.isna(valor):
      return pd.NaT

    if isinstance(valor, (pd.Timestamp, datetime, np.datetime64)):
      try:
        return pd.Timestamp(valor).normalize()
      except Exception:
        return pd.NaT

    if isinstance(valor, (int, float, np.integer, np.floating)):
      try:
        if float(valor) > 20000:
          return pd.to_datetime(
              float(valor), unit="D", origin="1899-12-30", errors="coerce"
          ).normalize()
      except Exception:
        pass

    texto = str(valor).strip()
    if not texto:
      return pd.NaT

    try:
      if re.fullmatch(r"\d{2}/\d{2}/\d{4}(?:\s+.*)?", texto):
        return pd.to_datetime(
            texto[:10], format="%d/%m/%Y", errors="coerce"
        ).normalize()
      if re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:.*)?", texto):
        return pd.to_datetime(
            texto[:10], format="%Y-%m-%d", errors="coerce"
        ).normalize()
      if re.fullmatch(r"\d{8}", texto[:8]):
        return pd.Timestamp(
            year=int(texto[:4]), month=int(texto[4:6]), day=int(texto[6:8])
        )
      return pd.NaT
    except Exception:
      return pd.NaT

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


def parse_data_filtro(val) -> Optional[pd.Timestamp]:
  if val is None:
    return None
  try:
    if pd.isna(val):
      return None
  except Exception:
    pass

  if isinstance(val, (pd.Timestamp, datetime, np.datetime64)):
    try:
      return pd.Timestamp(val).normalize()
    except Exception:
      return None

  if isinstance(val, (int, float, np.integer, np.floating)):
    try:
      n = float(val)
      if n > 20000:
        return pd.to_datetime(
            n, unit="D", origin="1899-12-30", errors="coerce"
        ).normalize()
    except Exception:
      return None

  texto = str(val).strip()
  if not texto:
    return None

  try:
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", texto)
    if m:
      return pd.Timestamp(
          year=int(m.group(3)), month=int(m.group(2)), day=int(m.group(1))
      )

    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", texto)
    if m:
      return pd.Timestamp(
          year=int(m.group(1)), month=int(m.group(2)), day=int(m.group(3))
      )

    if re.fullmatch(r"\d{8}", texto[:8]):
      return pd.Timestamp(
          year=int(texto[:4]), month=int(texto[4:6]), day=int(texto[6:8])
      )
    return pd.NaT
  except Exception:
    return None


def criar_data_filtro(valor) -> pd.Timestamp:
  parsed = parse_data_filtro(valor)
  return parsed if parsed is not None else pd.NaT


def calcular_dias_ausencia(data_inicio, data_fim):
  """Calcula dias corridos inclusivos entre a data inicial e final."""
  dt_ini = parse_data_filtro(data_inicio)
  dt_fim = parse_data_filtro(data_fim)
  if dt_ini is None or dt_fim is None:
    return 0
  if pd.isna(dt_ini) or pd.isna(dt_fim) or pd.Timestamp(dt_fim) < pd.Timestamp(dt_ini):
    return 0
  return int((pd.Timestamp(dt_fim) - pd.Timestamp(dt_ini)).days + 1)


# ================================================================
# LEITURA DE PLANILHAS
# ================================================================
class DataLoader:

  def __init__(self, config: AppConfig):
    self.config = config

  def carregar_precos(self) -> pd.DataFrame:
    df = pd.read_excel(self.config.CAMINHO_PRECOS, sheet_name=0)

    col_tipo = DataUtils.encontrar_coluna(
        df, ["TIPO", "TIPO VEICULO", "TIPO DE VEICULO", "CATEGORIA"]
    )
    col_media = DataUtils.encontrar_coluna(
        df, ["MEDIA", "MÉDIA", "MEDIA KM/L", "MÉDIA KM/L"]
    )
    col_premio = DataUtils.encontrar_coluna(
        df,
        [
            "TOTAL",
            "PREMIO",
            "PRÊMIO",
            "VALOR",
            "PREMIAÇÃO",
            "PREMIACAO",
            "VALOR PREMIO",
            "BONUS",
            "R$",
        ],
    )

    if None in (col_tipo, col_media, col_premio):
      for i in range(min(len(df), 10)):
        vals = [str(x).strip().upper() for x in df.iloc[i].tolist()]
        if any("TIPO" in v or "CATEGORIA" in v for v in vals) and any(
            "MEDIA" in v for v in vals
        ):
          df.columns = [DataUtils.normalizar_texto(c) for c in df.iloc[i]]
          df = df.iloc[i + 1 :].reset_index(drop=True)
          col_tipo = DataUtils.encontrar_coluna(
              df, ["TIPO", "TIPO VEICULO", "TIPO DE VEICULO", "CATEGORIA"]
          )
          col_media = DataUtils.encontrar_coluna(df, ["MEDIA", "MÉDIA"])
          col_premio = DataUtils.encontrar_coluna(
              df,
              [
                  "TOTAL",
                  "PREMIO",
                  "PRÊMIO",
                  "VALOR",
                  "PREMIAÇÃO",
                  "BONUS",
                  "R$",
              ],
          )
          break

    if None in (col_tipo, col_media, col_premio) and df.shape[1] >= 3:
      df.columns = [
          "TIPO",
          "MEDIA",
          "PREMIO",
      ] + list(df.columns[3:])
      col_tipo, col_media, col_premio = "TIPO", "MEDIA", "PREMIO"

    resultado = pd.DataFrame({
        "TIPO": df[col_tipo].apply(DataUtils.normalizar_texto),
        "MEDIA": df[col_media].apply(DataUtils.converter_numero),
        "PREMIO": df[col_premio].apply(DataUtils.converter_numero),
    }).dropna(subset=["MEDIA", "PREMIO"])

    resultado["TIPO"] = resultado["TIPO"].replace({"TOCO": "TRUCK"})
    return resultado[resultado["TIPO"] != ""].reset_index(drop=True)

  def carregar_frota(self) -> Tuple[pd.DataFrame, dict]:
    df_excel = pd.read_excel(self.config.CAMINHO_FROTA, sheet_name=0)
    col_placa = DataUtils.encontrar_coluna(
        df_excel, ["CAVALO", "PLACA", "PLACA CAVALO", "PLACA DO CAVALO"]
    )
    col_tipo = DataUtils.encontrar_coluna(
        df_excel, ["TIPO", "TIPO VEICULO", "TIPO DE VEICULO", "CATEGORIA"]
    )

    df_base = pd.DataFrame({
        "CAVALO": df_excel[col_placa].apply(DataUtils.padronizar_placa),
        "TIPO": df_excel[col_tipo].apply(DataUtils.normalizar_texto),
    })

    df_custom = carregar_frota_customizada()
    if not df_custom.empty:
      df_custom["CAVALO"] = df_custom["CAVALO"].apply(
          DataUtils.padronizar_placa
      )
      df_custom["TIPO"] = df_custom["TIPO"].apply(DataUtils.normalizar_texto)
      df_base = pd.concat([df_base, df_custom], ignore_index=True)

    resultado = pd.DataFrame({
        "PLACA_PADRONIZADA": df_base["CAVALO"],
        "TIPO": df_base["TIPO"],
    })
    resultado = resultado[resultado["PLACA_PADRONIZADA"] != ""].copy()
    resultado["TIPO"] = resultado["TIPO"].replace({"TOCO": "TRUCK"})
    resultado = resultado.drop_duplicates("PLACA_PADRONIZADA", keep="last")

    inativos_dict = carregar_inativos().get("PLACA", {})
    resultado["STATUS"] = resultado["PLACA_PADRONIZADA"].apply(
        lambda x: "INATIVO" if x in inativos_dict else "ATIVO"
    )
    resultado["DATA_INATIVACAO"] = resultado["PLACA_PADRONIZADA"].apply(
        lambda x: inativos_dict.get(x, "")
    )

    mapa = dict(zip(resultado["PLACA_PADRONIZADA"], resultado["TIPO"]))
    return resultado, mapa

  def carregar_cadastro_motoristas(self) -> pd.DataFrame:
    bruto = pd.read_excel(
        self.config.CAMINHO_MOTORISTAS, sheet_name=0, header=None
    )
    cab_idx, linha_cab = None, None

    for i in range(min(len(bruto), 15)):
      vals = [str(x).strip().upper() for x in bruto.iloc[i].tolist()]
      if "MOTORISTAS" in vals and "TIPO" in vals:
        cab_idx, linha_cab = i, vals
        break

    if cab_idx is None:
      cadastro = pd.DataFrame({
          "MOTORISTA_CADASTRO": bruto.iloc[:, 0].apply(
              DataUtils.normalizar_texto
          ),
          "TIPO_CADASTRO": bruto.iloc[:, 1]
          .apply(DataUtils.normalizar_texto)
          .replace({"TOCO": "TRUCK"}),
          "BASE_CADASTRO": (
              bruto.iloc[:, 2].apply(DataUtils.normalizar_texto)
              if bruto.shape[1] > 2
              else ""
          ),
      })
    else:
      idx_mot = linha_cab.index("MOTORISTAS")
      idx_tipo = linha_cab.index("TIPO")
      idx_base = linha_cab.index("BASE") if "BASE" in linha_cab else None

      cadastro = bruto.iloc[cab_idx + 1 :].copy()
      cadastro["MOTORISTA_CADASTRO"] = cadastro.iloc[:, idx_mot].apply(
          DataUtils.normalizar_texto
      )
      cadastro["TIPO_CADASTRO"] = (
          cadastro.iloc[:, idx_tipo]
          .apply(DataUtils.normalizar_texto)
          .replace({"TOCO": "TRUCK"})
      )
      cadastro["BASE_CADASTRO"] = (
          cadastro.iloc[:, idx_base].apply(DataUtils.normalizar_texto)
          if idx_base is not None
          else ""
      )
      cadastro = cadastro[
          (cadastro["MOTORISTA_CADASTRO"] != "")
          & (cadastro["TIPO_CADASTRO"] != "")
      ][["MOTORISTA_CADASTRO", "TIPO_CADASTRO", "BASE_CADASTRO"]]

    df_custom = carregar_motoristas_customizados()
    if not df_custom.empty:
      df_c_fmt = pd.DataFrame({
          "MOTORISTA_CADASTRO": df_custom["MOTORISTAS"].apply(
              DataUtils.normalizar_texto
          ),
          "TIPO_CADASTRO": df_custom["TIPO"]
          .apply(DataUtils.normalizar_texto)
          .replace({"TOCO": "TRUCK"}),
          "BASE_CADASTRO": df_custom["BASE"].apply(DataUtils.normalizar_texto),
      })
      cadastro = pd.concat([cadastro, df_c_fmt], ignore_index=True)

    cadastro["EH_FOLGUISTA"] = cadastro["TIPO_CADASTRO"].eq("FOLGUISTA")
    cadastro = cadastro.drop_duplicates("MOTORISTA_CADASTRO", keep="last")

    inativos_dict = carregar_inativos().get("MOTORISTA", {})
    cadastro["STATUS"] = cadastro["MOTORISTA_CADASTRO"].apply(
        lambda x: "INATIVO" if x in inativos_dict else "ATIVO"
    )
    cadastro["DATA_INATIVACAO"] = cadastro["MOTORISTA_CADASTRO"].apply(
        lambda x: inativos_dict.get(x, "")
    )

    return cadastro

  def carregar_abastecimentos(self, mapa_frota: dict) -> pd.DataFrame:
    df = pd.read_excel(
        self.config.CAMINHO_ABASTECIMENTOS,
        sheet_name=0,
        dtype=str,
        keep_default_na=False,
    )

    col_placa = DataUtils.encontrar_coluna(df, ["PLACA", "CAVALO"])
    col_km = DataUtils.encontrar_coluna(
        df, ["KM ATUAL", "KM", "KM_1", "QUILOMETRAGEM"]
    )
    col_litros = DataUtils.encontrar_coluna(
        df, ["QTDE", "LITROS", "QUANTIDADE", "QTD"]
    )
    col_valor = DataUtils.encontrar_coluna(
        df,
        [
            "VALOR TOTAL",
            "VALOR",
            "TOTAL",
            "VALOR_TOTAL",
            "VR TOTAL",
            "VLR TOTAL",
            "VALOR COMBUSTIVEL",
            "VALOR (R$)",
        ],
    )
    col_motorista = DataUtils.encontrar_coluna(
        df, ["CONDUTOR", "MOTORISTA", "MOTORISTAS"]
    )
    col_data = DataUtils.encontrar_coluna(
        df,
        [
            "DATA",
            "Data",
            "DATA ABASTECIMENTO",
            "DATA_ABASTECIMENTO",
            "DATA DO ABASTECIMENTO",
            "DATA/HORA",
            "DATA HORA",
            "DATA EMISSAO",
            "DT_ABASTECIMENTO",
            "DT ABAST",
        ],
    )

    resultado = df.copy()
    resultado["_ORDEM_ORIGINAL"] = np.arange(len(resultado))
    resultado["PLACA_PADRONIZADA"] = resultado[col_placa].apply(
        DataUtils.padronizar_placa
    )
    resultado["KM_ATUAL_NUM"] = resultado[col_km].apply(
        DataUtils.converter_numero
    )
    resultado["QTDE_NUM"] = resultado[col_litros].apply(
        DataUtils.converter_numero
    )
    resultado["VALOR_NUM"] = (
        resultado[col_valor].apply(DataUtils.converter_numero).fillna(0.0)
        if col_valor
        else 0.0
    )
    resultado["CONDUTOR_NORMALIZADO"] = (
        resultado[col_motorista]
        .fillna("SEM MOTORISTA")
        .apply(DataUtils.normalizar_texto)
    )

    if col_data:
      resultado["DATA_ORIGINAL"] = resultado[col_data]
      resultado["DATA_FILTRO"] = resultado[col_data].apply(criar_data_filtro)
      resultado["DATA_NUM"] = resultado["DATA_FILTRO"]
      resultado["DATA"] = resultado["DATA_FILTRO"]
    else:
      resultado["DATA_ORIGINAL"] = pd.NaT
      resultado["DATA_NUM"] = pd.NaT
      resultado["DATA_FILTRO"] = pd.NaT

    resultado["TIPO"] = resultado["PLACA_PADRONIZADA"].map(mapa_frota)

    sem_tipo_mask = resultado["TIPO"].isna()
    placas_sem_tipo = resultado.loc[
        sem_tipo_mask, "PLACA_PADRONIZADA"
    ].unique()
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
        & resultado["KM_ATUAL_NUM"].notna()
        & (resultado["KM_ATUAL_NUM"] > 0)
        & resultado["QTDE_NUM"].notna()
        & (resultado["QTDE_NUM"] > 0)
    )

    return resultado


# ================================================================
# CÁLCULO DE PREMIAÇÕES
# ================================================================
class RewardEngine:

  @staticmethod
  def calcular_eventos_consumo(abastecimentos: pd.DataFrame) -> pd.DataFrame:
    base = abastecimentos[abastecimentos["REGISTRO_VALIDO"]].copy()
    base["_DATA_ORDENACAO"] = base["DATA_NUM"].fillna(
        pd.Timestamp("1900-01-01")
    )
    base = base.sort_values(
        ["PLACA_PADRONIZADA", "_DATA_ORDENACAO", "_ORDEM_ORIGINAL"],
        kind="stable",
    ).copy()

    base["KM_ANTERIOR"] = base.groupby("PLACA_PADRONIZADA")[
        "KM_ATUAL_NUM"
    ].shift(1)
    base["KM_RODADO_EVENTO"] = base["KM_ATUAL_NUM"] - base["KM_ANTERIOR"]

    base["REGISTRO_CONSUMO_VALIDO"] = (
        base["KM_ANTERIOR"].notna()
        & base["KM_RODADO_EVENTO"].notna()
        & (base["KM_RODADO_EVENTO"] > 0)
        & base["QTDE_NUM"].notna()
        & (base["QTDE_NUM"] > 0)
    )

    base["KM_CONSUMO"] = np.where(
        base["REGISTRO_CONSUMO_VALIDO"], base["KM_RODADO_EVENTO"], np.nan
    )
    base["LITROS_CONSUMO"] = np.where(
        base["REGISTRO_CONSUMO_VALIDO"], base["QTDE_NUM"], np.nan
    )
    base["TIPO_CALCULO"] = (
        base["TIPO"].fillna("GERAL").replace({"TOCO": "TRUCK"})
    )

    return base.reset_index(drop=True)

  def faixa_mais_proxima(
      self,
      media: float,
      tipo: str,
      precos: pd.DataFrame,
      base_motorista: str = "",
  ) -> dict:
    if pd.isna(media) or media <= 0:
      return {
          "MEDIA_FAIXA": np.nan,
          "PREMIO": 0.0,
          "STATUS_PREMIO": "SEM MEDIA",
      }

    tipo_norm = DataUtils.normalizar_texto(tipo)
    base_norm = DataUtils.normalizar_texto(base_motorista)

    tabela = pd.DataFrame()

    if ("BITRUCK" in tipo_norm or tipo_norm == "BITRUCK") and (
        "UBERABA" in base_norm or "UBERABA" in tipo_norm
    ):
      tabela = precos[precos["TIPO"].str.contains("UBERABA", na=False)].copy()

    if tabela.empty:
      tabela = precos[precos["TIPO"] == tipo_norm].copy()

    if tabela.empty:
      tabela = precos[precos["TIPO"].str.contains(tipo_norm, na=False)].copy()

    if tabela.empty and " " in tipo_norm:
      primeira_palavra = tipo_norm.split()[0]
      tabela = precos[
          precos["TIPO"].str.contains(primeira_palavra, na=False)
      ].copy()

    if tabela.empty:
      tabela = precos.copy()

    if tabela.empty:
      return {
          "MEDIA_FAIXA": np.nan,
          "PREMIO": 0.0,
          "STATUS_PREMIO": "SEM FAIXA",
      }

    tabela = tabela.sort_values("MEDIA").copy()
    media_min = float(tabela["MEDIA"].min())
    media_max = float(tabela["MEDIA"].max())
    media_class = round(float(media), 2)

    if media_class < media_min:
      return {
          "MEDIA_FAIXA": media_min,
          "PREMIO": 0.0,
          "STATUS_PREMIO": "DESCLASSIFICADO",
      }

    if media_class >= media_max:
      linha_max = tabela.iloc[-1]
      return {
          "MEDIA_FAIXA": float(linha_max["MEDIA"]),
          "PREMIO": float(linha_max["PREMIO"]),
          "STATUS_PREMIO": "TETO",
      }

    faixa_atingida = tabela[tabela["MEDIA"] <= media_class]
    if not faixa_atingida.empty:
      linha = faixa_atingida.iloc[-1]
      return {
          "MEDIA_FAIXA": float(linha["MEDIA"]),
          "PREMIO": float(linha["PREMIO"]),
          "STATUS_PREMIO": "OK",
      }

    return {
        "MEDIA_FAIXA": np.nan,
        "PREMIO": 0.0,
        "STATUS_PREMIO": "DESCLASSIFICADO",
    }

  def calcular_premios(
      self,
      eventos: pd.DataFrame,
      precos: pd.DataFrame,
      cadastro: pd.DataFrame,
      categorias_customizadas: dict = None,
  ) -> pd.DataFrame:
    colunas_saida = [
        "MOTORISTA_CHAVE",
        "CATEGORIA_ABASTECIMENTO",
        "MOTORISTA",
        "BASE",
        "CATEGORIA",
        "STATUS_MOTORISTA",
        "KM_TOTAL",
        "LITROS_TOTAL",
        "KM_CAT",
        "LITROS_CAT",
        "QTD_ABASTECIMENTOS",
        "PLACAS",
        "MEDIA_CALCULADA",
        "MEDIA_FAIXA",
        "PREMIO",
        "STATUS_PREMIO",
        "PREMIO_BRUTO",
        "DIAS_AUSENCIA",
        "DIAS_EFETIVOS",
        "MOTIVO_DESCLASSIFICACAO",
    ]

    if categorias_customizadas is None:
      categorias_customizadas = {}

    if eventos.empty:
      return pd.DataFrame(columns=colunas_saida)

    base = eventos[eventos["REGISTRO_CONSUMO_VALIDO"]].copy()
    if base.empty:
      return pd.DataFrame(columns=colunas_saida)

    base["CATEGORIA_ABASTECIMENTO"] = (
        base["TIPO"].fillna("GERAL").replace({"TOCO": "TRUCK"})
    )
    base["MOTORISTA_CHAVE"] = (
        base["CONDUTOR_NORMALIZADO"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    cad = cadastro.copy()
    cad["MOTORISTA_CHAVE"] = (
        cad["MOTORISTA_CADASTRO"].astype(str).str.strip().str.upper()
    )

    base = base.merge(
        cad[[
            "MOTORISTA_CHAVE",
            "TIPO_CADASTRO",
            "BASE_CADASTRO",
            "EH_FOLGUISTA",
            "STATUS",
        ]],
        on="MOTORISTA_CHAVE",
        how="left",
    )
    base["TIPO_CADASTRO"] = base["TIPO_CADASTRO"].fillna("")
    base["EH_FOLGUISTA"] = base["EH_FOLGUISTA"].fillna(False)
    base["STATUS"] = base["STATUS"].fillna("ATIVO")

    registros = []
    for _, grupo in base.groupby("MOTORISTA_CHAVE", sort=False):
      mot_chave = grupo["MOTORISTA_CHAVE"].iloc[0]
      eh_folguista = grupo["EH_FOLGUISTA"].iloc[0]
      tipo_cad = grupo["TIPO_CADASTRO"].iloc[0]
      custom_legacy = categorias_customizadas.get(mot_chave, "")

      # Regra nova: a categoria manual pode ser definida por MOTORISTA + PLACA.
      # Regra antiga continua aceita quando só houver MOTORISTA_CHAVE.
      grupo = grupo.copy()
      grupo["CHAVE_CAT_MANUAL"] = (
          grupo.apply(
              lambda r: normalizar_chave_categoria_customizada(
                  r["MOTORISTA_CHAVE"], r["PLACA_PADRONIZADA"]
              ),
              axis=1,
          )
      )
      grupo["CATEGORIA_MANUAL"] = grupo["CHAVE_CAT_MANUAL"].map(
          categorias_customizadas
      ).fillna("")

      if custom_legacy:
        grupo["CATEGORIA_MANUAL"] = grupo["CATEGORIA_MANUAL"].replace("", custom_legacy)

      if eh_folguista and not custom_legacy and not grupo["CATEGORIA_MANUAL"].astype(bool).any():
        soma = grupo.groupby("CATEGORIA_ABASTECIMENTO")["KM_CONSUMO"].sum()
        cat_elegivel = (
            soma.idxmax()
            if not soma.empty
            else (tipo_cad if tipo_cad else "TRUCK")
        )
        grupo["CATEGORIA_ELEGIVEL"] = cat_elegivel
        grupo["USA_CATEGORIA_MANUAL"] = False
      else:
        grupo["CATEGORIA_ELEGIVEL"] = np.where(
            grupo["CATEGORIA_MANUAL"].astype(str).str.strip() != "",
            grupo["CATEGORIA_MANUAL"],
            tipo_cad if tipo_cad else grupo["CATEGORIA_ABASTECIMENTO"].iloc[0],
        )
        grupo["USA_CATEGORIA_MANUAL"] = (
            grupo["CATEGORIA_MANUAL"].astype(str).str.strip() != ""
        )

      registros.append(grupo)

    if not registros:
      return pd.DataFrame(columns=colunas_saida)

    df_eleg = pd.concat(registros, ignore_index=True)

    df_eleg["EH_CATEGORIA_ELEGIVEL"] = np.where(
        df_eleg["USA_CATEGORIA_MANUAL"],
        True,
        df_eleg["CATEGORIA_ABASTECIMENTO"] == df_eleg["CATEGORIA_ELEGIVEL"],
    )
    df_eleg["KM_CAT_ELEGIVEL"] = np.where(
        df_eleg["EH_CATEGORIA_ELEGIVEL"], df_eleg["KM_CONSUMO"], 0
    )
    df_eleg["LITROS_CAT_ELEGIVEL"] = np.where(
        df_eleg["EH_CATEGORIA_ELEGIVEL"], df_eleg["LITROS_CONSUMO"], 0
    )

    resumo = df_eleg.groupby("MOTORISTA_CHAVE", as_index=False).agg(
        MOTORISTA=("CONDUTOR_NORMALIZADO", "first"),
        BASE=("BASE_CADASTRO", "first"),
        CATEGORIA=("CATEGORIA_ELEGIVEL", "first"),
        STATUS_MOTORISTA=("STATUS", "first"),
        KM_TOTAL=("KM_CONSUMO", "sum"),
        LITROS_TOTAL=("LITROS_CONSUMO", "sum"),
        KM_CAT=("KM_CAT_ELEGIVEL", "sum"),
        LITROS_CAT=("LITROS_CAT_ELEGIVEL", "sum"),
        QTD_ABASTECIMENTOS=("LITROS_CONSUMO", "count"),
        PLACAS=("PLACA_PADRONIZADA", lambda s: " | ".join(sorted(set(s)))),
    )

    resumo["MEDIA_CALCULADA"] = np.where(
        resumo["LITROS_CAT"] > 0,
        resumo["KM_CAT"] / resumo["LITROS_CAT"],
        np.nan,
    )

    faixas = resumo.apply(
        lambda r: self.faixa_mais_proxima(
            r["MEDIA_CALCULADA"],
            r["CATEGORIA"],
            precos,
            base_motorista=r["BASE"],
        ),
        axis=1,
    )

    resumo_df = pd.concat(
        [resumo.reset_index(drop=True), pd.DataFrame(list(faixas))], axis=1
    )
    resumo_df["PREMIO_BRUTO"] = resumo_df["PREMIO"]
    resumo_df["DIAS_AUSENCIA"] = 0
    resumo_df["DIAS_EFETIVOS"] = 30
    resumo_df["MOTIVO_DESCLASSIFICACAO"] = np.where(
        resumo_df["STATUS_PREMIO"] == "DESCLASSIFICADO",
        "Média de consumo abaixo do limite mínimo da categoria",
        "Elegível / Em conformidade",
    )

    return resumo_df


# ================================================================
# GERADOR DE DADOS DE MÚLTIPLAS PLACAS
# ================================================================
def gerar_dados_multiplas_placas(eventos: pd.DataFrame):
  if eventos.empty:
    return pd.DataFrame(), []

  base = eventos[eventos["REGISTRO_CONSUMO_VALIDO"]].copy()
  if base.empty:
    return pd.DataFrame(), []

  base["MOTORISTA_CHAVE"] = (
      base["CONDUTOR_NORMALIZADO"]
      .fillna("")
      .astype(str)
      .str.strip()
      .str.upper()
  )
  base["CATEGORIA_ABASTECIMENTO"] = (
      base["TIPO"].fillna("GERAL").replace({"TOCO": "TRUCK"})
  )

  placas_por_motorista = base.groupby("MOTORISTA_CHAVE")[
      "PLACA_PADRONIZADA"
  ].nunique()
  mots_multi = placas_por_motorista[placas_por_motorista > 1].index.tolist()

  base_multi = base[base["MOTORISTA_CHAVE"].isin(mots_multi)].copy()
  if base_multi.empty:
    return pd.DataFrame(), []

  resumo_placa = base_multi.groupby(
      [
          "MOTORISTA_CHAVE",
          "CONDUTOR_NORMALIZADO",
          "PLACA_PADRONIZADA",
          "CATEGORIA_ABASTECIMENTO",
      ],
      as_index=False,
  ).agg(
      KM_TOTAL=("KM_CONSUMO", "sum"),
      LITROS_TOTAL=("LITROS_CONSUMO", "sum"),
      QTD_ABASTECIMENTOS=("LITROS_CONSUMO", "count"),
  )

  resumo_placa["MEDIA_CALCULADA"] = np.where(
      resumo_placa["LITROS_TOTAL"] > 0,
      resumo_placa["KM_TOTAL"] / resumo_placa["LITROS_TOTAL"],
      np.nan,
  )

  df_exibicao = resumo_placa.rename(columns={
      "CONDUTOR_NORMALIZADO": "MOTORISTA",
      "PLACA_PADRONIZADA": "PLACA",
      "CATEGORIA_ABASTECIMENTO": "CATEGORIA_VEICULO",
      "KM_TOTAL": "KM TOTAL",
      "LITROS_TOTAL": "LITROS TOTAL",
      "MEDIA_CALCULADA": "MÉDIA (KM/L)",
      "QTD_ABASTECIMENTOS": "QTD ABASTECIMENTOS",
  })[[
      "MOTORISTA",
      "PLACA",
      "CATEGORIA_VEICULO",
      "KM TOTAL",
      "LITROS TOTAL",
      "MÉDIA (KM/L)",
      "QTD ABASTECIMENTOS",
  ]]

  df_exibicao["MÉDIA (KM/L)"] = df_exibicao["MÉDIA (KM/L)"].apply(
      lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else "-"
  )

  lista_nomes_mots = sorted(
      base_multi["CONDUTOR_NORMALIZADO"].unique().tolist()
  )
  return df_exibicao, lista_nomes_mots


# ================================================================
# CONSTANTES DE CRITÉRIOS DE DESCLASSIFICAÇÃO (PILAR 1)
# ================================================================
CRITERIOS_PILAR_1 = [
    (
        "1 - Controle de velocidade, limite máx de 80 km/h (1 Ponto/evento até"
        " 129)"
    ),
    (
        "2 - Controle de jornada, macros e intervalos incorretos (1"
        " Ponto/evento até 129)"
    ),
    "3 - Não realização correta do check list diário (1 Ponto/evento até 129)",
    (
        "4 - Deszelo com documentação e comprovantes de carga/descarga (1"
        " Ponto/evento até 129)"
    ),
    (
        "5 - Distração com freadas bruscas, risco de colisão ou manobra perigosa"
        " [DESCLASSIFICADO]"
    ),
    "6 - Uso do celular em direção [DESCLASSIFICADO]",
    "7 - Estacionar ou parada com veículo em L [DESCLASSIFICADO]",
    "8 - Identificação de ausência do motorista [DESCLASSIFICADO]",
    "9 - Paradas em locais proibidos [DESCLASSIFICADO]",
    "10 - Ausência do cinto de segurança em direção [DESCLASSIFICADO]",
    "11 - Vedar as câmeras [DESCLASSIFICADO]",
    (
        "12 - Comportamento inadequado do motorista dentro e fora do veículo"
        " [DESCLASSIFICADO]"
    ),
    (
        "13 - Picos acima de 80 km/h, com mais de 30 infrações por período"
        " [DESCLASSIFICADO]"
    ),
    "14 - Não cumprir determinações e escalas estipuladas [DESCLASSIFICADO]",
    (
        "15 - Erros operacionais (Carregamento/Descarregamento incorreto ou"
        " Derramamento/Contaminação) [DESCLASSIFICADO]"
    ),
]


# ================================================================
# RECALCULO DE AUSÊNCIAS E DESCLASSIFICAÇÕES
# ================================================================
def aplicar_regras_gerais(
    df_resumo_original: pd.DataFrame,
    df_ausencias: pd.DataFrame,
    df_desclassificacoes: pd.DataFrame,
) -> pd.DataFrame:
  if df_resumo_original.empty:
    return df_resumo_original.copy()

  res = df_resumo_original.copy()

  motivos_iniciais = []
  for idx in res.index:
    st = res.at[idx, "STATUS_PREMIO"]
    if st == "DESCLASSIFICADO":
      motivos_iniciais.append(
          "Média de consumo abaixo do limite mínimo da categoria"
      )
    else:
      motivos_iniciais.append("Elegível / Em conformidade")
  res["MOTIVO_DESCLASSIFICACAO"] = motivos_iniciais

  if not df_ausencias.empty and "MOTORISTA" in df_ausencias.columns:
    df_aus_tmp = df_ausencias.copy()
    df_aus_tmp["DIAS"] = pd.to_numeric(
        df_aus_tmp["DIAS"], errors="coerce"
    ).fillna(0)
    soma_dias = df_aus_tmp.groupby("MOTORISTA")["DIAS"].sum().to_dict()

    res["DIAS_AUSENCIA"] = (
        res["MOTORISTA"].map(soma_dias).fillna(0).astype(int)
    )
    res["DIAS_EFETIVOS"] = np.maximum(0, 30 - res["DIAS_AUSENCIA"])
    res["PREMIO"] = res.apply(
        lambda r: max(0.0, r["PREMIO_BRUTO"] * (r["DIAS_EFETIVOS"] / 30.0)),
        axis=1,
    )
  else:
    res["DIAS_AUSENCIA"] = 0
    res["DIAS_EFETIVOS"] = 30
    res["PREMIO"] = res["PREMIO_BRUTO"]

  if (
      not df_desclassificacoes.empty
      and "MOTORISTA" in df_desclassificacoes.columns
  ):
    for idx in res.index:
      m_nome = res.at[idx, "MOTORISTA"]
      g = df_desclassificacoes[df_desclassificacoes["MOTORISTA"] == m_nome]
      if not g.empty:
        motivos_pilar1 = []
        diretos = g[g["TIPO_IMPACTO"] == "DESCLASSIFICADO"]
        if not diretos.empty:
          crit_limpos = [
              c.split("[")[0].strip() for c in diretos["CRITERIO"].unique()
          ]
          motivos_pilar1.append(f"Infração Pilar 1: {', '.join(crit_limpos)}")

        tot_pontos = pd.to_numeric(g["PONTOS"], errors="coerce").fillna(0).sum()
        if tot_pontos > 129:
          motivos_pilar1.append(
              f"Excesso de Pontos Pilar 1 ({int(tot_pontos)} pts - máx 129)"
          )

        if motivos_pilar1:
          res.at[idx, "PREMIO"] = 0.0
          res.at[idx, "STATUS_PREMIO"] = "DESCLASSIFICADO"

          mot_existente = res.at[idx, "MOTIVO_DESCLASSIFICACAO"]
          novo_mot = " | ".join(motivos_pilar1)
          if mot_existente and "Média" in mot_existente:
            res.at[idx, "MOTIVO_DESCLASSIFICACAO"] = (
                f"{mot_existente} + {novo_mot}"
            )
          else:
            res.at[idx, "MOTIVO_DESCLASSIFICACAO"] = novo_mot

  # Motorista INATIVO não participa do pagamento do prêmio.
  # O histórico continua no sistema, mas o valor a pagar é zerado.
  if "STATUS_MOTORISTA" in res.columns:
    inativos = (
        res["STATUS_MOTORISTA"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("INATIVO")
    )
    if inativos.any():
      res.loc[inativos, "PREMIO"] = 0.0
      res.loc[inativos, "STATUS_PREMIO"] = "INATIVO"
      res.loc[inativos, "MOTIVO_DESCLASSIFICACAO"] = (
          "Motorista inativo - prêmio desconsiderado"
      )

  return res


# ================================================================
# GERADOR DE DATAFRAME EXCLUSIVO DO RH
# ================================================================
def gerar_tabela_rh(df_resumo: pd.DataFrame) -> pd.DataFrame:
  if df_resumo.empty or "MOTORISTA" not in df_resumo.columns:
    return pd.DataFrame(columns=["NOME", "FILIAL", "VALOR PAGO"])

  if "STATUS_MOTORISTA" in df_resumo.columns:
    df_rh = df_resumo[df_resumo["STATUS_MOTORISTA"] == "ATIVO"].copy()
  else:
    df_rh = df_resumo.copy()

  rh_df = pd.DataFrame()
  rh_df["NOME"] = df_rh["MOTORISTA"]
  rh_df["FILIAL"] = df_rh["BASE"].fillna("CIANORTE")

  def formatar_valor_pago(x):
    if pd.isna(x):
      return "R$ 0,00"
    try:
      return (
          f"R$ {float(x):,.2f}".replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
    except Exception:
      return "R$ 0,00"

  rh_df["VALOR PAGO"] = df_rh["PREMIO"].map(formatar_valor_pago)
  return rh_df


# ================================================================
# FILTRAGEM E RECÁLCULO DO DASHBOARD POR DATA E DEMAIS CRITÉRIOS
# ================================================================
def aplicar_filtros(
    dt_ini,
    dt_fim,
    motorista,
    placa,
    categoria,
    filial,
    df_ausencias,
    df_desclassificacoes,
    mapa_cat_custom=None,
):
  if mapa_cat_custom is None:
    mapa_cat_custom = carregar_categorias_customizadas()

  evt_f = eventos.copy()
  det_f = abastecimentos.copy()

  d_i = parse_data_filtro(dt_ini)
  d_f = parse_data_filtro(dt_fim)

  def aplicar_periodo(df):
    if df.empty:
      return df

    out = df.copy()
    if "DATA_FILTRO" in out.columns:
      datas = pd.to_datetime(out["DATA_FILTRO"], errors="coerce").dt.normalize()
    elif "DATA_NUM" in out.columns:
      datas = pd.to_datetime(out["DATA_NUM"], errors="coerce").dt.normalize()
    elif "DATA_ORIGINAL" in out.columns:
      datas = out["DATA_ORIGINAL"].apply(parse_data_filtro)
    else:
      return out

    mascara = datas.notna()

    if d_i is not None and pd.notna(d_i):
      ini = pd.Timestamp(d_i).normalize()
      mascara &= datas >= ini

    if d_f is not None and pd.notna(d_f):
      fim_exclusivo = pd.Timestamp(d_f).normalize() + pd.Timedelta(days=1)
      mascara &= datas < fim_exclusivo

    out["DATA_FILTRO"] = datas
    return out.loc[mascara].copy()

  evt_f = aplicar_periodo(evt_f)
  det_f = aplicar_periodo(det_f)

  if motorista and motorista != "TODOS":
    m_norm = DataUtils.normalizar_texto(motorista)
    evt_f = evt_f[
        evt_f["CONDUTOR_NORMALIZADO"]
        .apply(DataUtils.normalizar_texto)
        .str.contains(m_norm, na=False)
    ]
    det_f = det_f[
        det_f["CONDUTOR_NORMALIZADO"]
        .apply(DataUtils.normalizar_texto)
        .str.contains(m_norm, na=False)
    ]

  if placa and placa.strip():
    p_norm = DataUtils.padronizar_placa(placa)
    evt_f = evt_f[evt_f["PLACA_PADRONIZADA"] == p_norm]
    det_f = det_f[det_f["PLACA_PADRONIZADA"] == p_norm]

  if categoria and categoria != "TODAS":
    c_norm = DataUtils.normalizar_texto(categoria)
    evt_f = evt_f[evt_f["TIPO_CALCULO"] == c_norm]
    det_f = det_f[
        det_f["TIPO"]
        .fillna("")
        .replace({"TOCO": "TRUCK"})
        .apply(DataUtils.normalizar_texto)
        == c_norm
    ]

  if filial and filial != "TODAS":
    f_norm = DataUtils.normalizar_texto(filial)
    mapa_filial = {}
    if not cadastro.empty:
      cad_tmp = cadastro.copy()
      for _, row in cad_tmp.iterrows():
        chave = DataUtils.normalizar_texto(row.get("MOTORISTA_CADASTRO", ""))
        base = DataUtils.normalizar_texto(row.get("BASE_CADASTRO", ""))
        if chave:
          mapa_filial[chave] = base

    evt_f = evt_f[
        evt_f["CONDUTOR_NORMALIZADO"].map(
            lambda x: mapa_filial.get(DataUtils.normalizar_texto(x), "")
        )
        == f_norm
    ]
    det_f = det_f[
        det_f["CONDUTOR_NORMALIZADO"].map(
            lambda x: mapa_filial.get(DataUtils.normalizar_texto(x), "")
        )
        == f_norm
    ]

  resumo_periodo = engine.calcular_premios(
      evt_f, precos, cadastro, categorias_customizadas=mapa_cat_custom
  )
  res_f = aplicar_regras_gerais(
      resumo_periodo, df_ausencias, df_desclassificacoes
  )

  df_multi, mots_multi = gerar_dados_multiplas_placas(evt_f)

  tot_premio = (
      float(res_f["PREMIO"].sum())
      if (not res_f.empty and "PREMIO" in res_f.columns)
      else 0.0
  )
  if pd.isna(tot_premio):
    tot_premio = 0.0

  tot_km = (
      float(res_f["KM_TOTAL"].sum())
      if (not res_f.empty and "KM_TOTAL" in res_f.columns)
      else 0.0
  )
  if pd.isna(tot_km):
    tot_km = 0.0

  tot_litros = (
      float(res_f["LITROS_TOTAL"].sum())
      if (not res_f.empty and "LITROS_TOTAL" in res_f.columns)
      else 0.0
  )
  if pd.isna(tot_litros):
    tot_litros = 0.0

  tot_gasto_combustivel = (
      float(det_f["VALOR_NUM"].sum())
      if (not det_f.empty and "VALOR_NUM" in det_f.columns)
      else 0.0
  )
  if pd.isna(tot_gasto_combustivel):
    tot_gasto_combustivel = 0.0

  tot_media_geral = (tot_km / tot_litros) if tot_litros > 0 else 0.0
  if pd.isna(tot_media_geral):
    tot_media_geral = 0.0

  tot_mots = len(res_f)

  res_view = res_f.copy()
  if not res_view.empty:

    def fmt_premio(x):
      if pd.isna(x):
        return "R$ 0,00"
      try:
        return (
            f"R$ {float(x):,.2f}".replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
      except Exception:
        return "R$ 0,00"

    def fmt_media(x):
      if pd.isna(x):
        return "-"
      try:
        return f"{float(x):.2f}".replace(".", ",")
      except Exception:
        return str(x)

    if "PREMIO" in res_view.columns:
      res_view["PREMIO"] = res_view["PREMIO"].map(fmt_premio)
    if "PREMIO_BRUTO" in res_view.columns:
      res_view["PREMIO_BRUTO"] = res_view["PREMIO_BRUTO"].map(fmt_premio)
    if "MEDIA_CALCULADA" in res_view.columns:
      res_view["MEDIA_CALCULADA"] = res_view["MEDIA_CALCULADA"].map(fmt_media)
    if "MEDIA_FAIXA" in res_view.columns:
      res_view["MEDIA_FAIXA"] = res_view["MEDIA_FAIXA"].map(fmt_media)

  col_data_detalhe = (
      "DATA_FILTRO" if "DATA_FILTRO" in det_f.columns else "DATA_NUM"
  )
  det_view = det_f.sort_values(
      [col_data_detalhe, "_ORDEM_ORIGINAL"], kind="stable"
  ).copy()
  det_view = det_view.drop(
      columns=[c for c in ["_DATA_ORDENACAO"] if c in det_view.columns],
      errors="ignore",
  )

  rh_view = gerar_tabela_rh(res_f)

  f_premio = (
      f"R$ {tot_premio:,.2f}".replace(",", "X")
      .replace(".", ",")
      .replace("X", ".")
  )
  f_gasto_comb = (
      f"R$ {tot_gasto_combustivel:,.2f}".replace(",", "X")
      .replace(".", ",")
      .replace("X", ".")
  )
  f_km = (
      f"{tot_km:,.1f} km".replace(",", "X")
      .replace(".", ",")
      .replace("X", ".")
  )
  f_litros = (
      f"{tot_litros:,.1f} L".replace(",", "X")
      .replace(".", ",")
      .replace("X", ".")
  )
  f_media = f"{tot_media_geral:.2f} km/L".replace(".", ",")
  f_mots = f"{tot_mots}"

  dropdown_mots_multi = mots_multi if mots_multi else ["NENHUM MOTORISTA"]

  return (
      f_premio,
      f_gasto_comb,
      f_km,
      f_litros,
      f_media,
      f_mots,
      res_view,
      rh_view,
      det_view,
      df_multi,
      dropdown_mots_multi,
      res_f,
  )


# ================================================================
# GERADOR DE RECIBOS DE PREMIAÇÃO
# ================================================================
def gerar_html_unico_recibo(
    row_data: pd.Series,
    motorista_sel: str,
    periodo_ini: str,
    periodo_fim: str,
    fator_c: str,
) -> str:
  base_val = row_data.get("BASE", "")
  if pd.isna(base_val) or str(base_val).strip() == "":
    base_val = "CIANORTE"

  tipo_val = row_data.get("CATEGORIA", "")

  km_raw = row_data.get("KM_CAT", row_data.get("KM_TOTAL", 0))
  try:
    km_val = (
        "0" if pd.isna(km_raw) else f"{float(km_raw):,.0f}".replace(",", ".")
    )
  except Exception:
    km_val = "0"

  dias_ef = row_data.get("DIAS_EFETIVOS", 30)
  dias_efetivos_val = 30 if pd.isna(dias_ef) else int(dias_ef)

  media_raw = row_data.get("MEDIA_CALCULADA", np.nan)
  if pd.isna(media_raw):
    media_val = "-"
  else:
    try:
      media_val = f"{float(media_raw):.2f}".replace(".", ",")
    except Exception:
      media_val = str(media_raw)

  premio_raw = row_data.get("PREMIO", 0.0)
  premio_float = 0.0 if pd.isna(premio_raw) else float(premio_raw)
  val_total_str = (
      f"R$ {premio_float:,.2f}".replace(",", "X")
      .replace(".", ",")
      .replace("X", ".")
  )

  motivo_descl = row_data.get(
      "MOTIVO_DESCLASSIFICACAO", "Elegível / Em conformidade"
  )
  if pd.isna(motivo_descl) or not str(motivo_descl).strip():
    motivo_descl = "Elegível / Em conformidade"

  status_premio = row_data.get("STATUS_PREMIO", "OK")
  eh_desclassificado = (status_premio == "DESCLASSIFICADO") or (
      premio_float == 0.0 and motivo_descl != "Elegível / Em conformidade"
  )

  bg_motivo = "#FEE2E2" if eh_desclassificado else "#FFFFFF"
  fg_motivo = "#991B1B" if eh_desclassificado else "#000000"
  hdr_motivo_bg = "#EF4444" if eh_desclassificado else "#D0E0F0"
  hdr_motivo_fg = "#FFFFFF" if eh_desclassificado else "#000000"

  return f"""
    <div class="recibo-card" style="background-color: #FFFFFF; padding: 28px; border-radius: 12px; max-width: 650px; margin: 0 auto 20px auto; font-family: Arial, sans-serif; color: #000000; border: 1px solid #CBD5E1; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); page-break-after: always; break-after: page;">
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


def gerar_recibos_lote(
    filial_sel: str,
    motorista_sel: str,
    periodo_ini: str,
    periodo_fim: str,
    fator_c: str,
    df_resumo: pd.DataFrame,
) -> str:
  if not motorista_sel or motorista_sel in ("SELECIONE...", ""):
    return (
        "<div style='text-align: center; padding: 40px; color: #64748B; "
        "font-size: 15px;'>👉 Por favor, selecione um motorista ou a opção de "
        "TODOS DA FILIAL para gerar os recibos.</div>"
    )

  res_f = df_resumo.copy()

  if "STATUS_MOTORISTA" in res_f.columns:
    res_f = res_f[res_f["STATUS_MOTORISTA"] == "ATIVO"]

  if str(motorista_sel).startswith("TODOS"):
    if filial_sel and filial_sel != "TODAS":
      f_norm = DataUtils.normalizar_texto(filial_sel)
      res_f = res_f[res_f["BASE"].apply(DataUtils.normalizar_texto) == f_norm]

    lista_mots = sorted(list(res_f["MOTORISTA"].dropna().unique()))
    if not lista_mots:
      return (
          f"<div style='text-align: center; padding: 40px; color:"
          f" #EF4444;'>Nenhum motorista encontrado na filial '{filial_sel}'.</div>"
      )
  else:
    lista_mots = [motorista_sel]

  recibos_html = [f"""
    <style>
    @media print {{
        body * {{ visibility: hidden; }}
        .recibo-container, .recibo-container * {{ visibility: visible; }}
        .recibo-container {{ position: absolute; left: 0; top: 0; width: 100%; }}
        .no-print {{ display: none !important; }}
    }}
    </style>
    <div class="no-print" style="background: #F8FAFC; border: 1px solid #E2E8F0; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center;">
        <span style="font-size: 14px; font-weight: bold; color: #1E293B;">
            📄 Total de Recibos Prontos: <span style="color: #2563EB;">{len(lista_mots)}</span>
        </span>
        <button onclick="window.print()" style="background-color: #2563EB; color: #FFFFFF; border: none; padding: 8px 18px; border-radius: 6px; font-weight: bold; cursor: pointer; font-size: 13px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            🖨️ Imprimir Todos os Recibos ({len(lista_mots)})
        </button>
    </div>
    """]

  cards_html = []
  for m_nome in lista_mots:
    row = res_f[res_f["MOTORISTA"] == m_nome]
    if not row.empty:
      card_html = gerar_html_unico_recibo(
          row.iloc[0], m_nome, periodo_ini, periodo_fim, fator_c
      )
      cards_html.append(card_html)

  if not cards_html:
    return (
        "<div style='text-align: center; padding: 40px; color: #64748B;'>👉"
        " Nenhum recibo gerado (verifique se o motorista selecionado está"
        " Inativo).</div>"
    )

  return (
      "".join(recibos_html)
      + "<div class='recibo-container' style='display: flex; flex-direction:"
      " column; gap: 30px;'>"
      + "".join(cards_html)
      + "</div>"
  )




st.set_page_config(page_title="Dashboard do Prêmio de Motoristas", page_icon="🚚", layout="wide", initial_sidebar_state="expanded")

st.markdown('''
<style>
:root { --brand:#1E2B7A; --blue:#0099DA; --gold:#FFD700; --bg:#F6F8FC; --card:#FFFFFF; --muted:#64748B; }
.stApp { background: var(--bg); }
.block-container { max-width: 1500px; padding-top: 1.2rem; padding-bottom: 2rem; }
.hero { background: linear-gradient(135deg, #ffffff 0%, #f1f5ff 100%); border:1px solid #e2e8f0; border-radius:20px; padding:22px 28px; margin-bottom:18px; box-shadow:0 4px 18px rgba(15,23,42,.05); }
.hero h1 { margin:0; color:#0f172a; font-size:2rem; }
.hero p { margin:.3rem 0 0; color:var(--muted); }
.logo-bar { display:flex; align-items:center; gap:14px; }
.logo { width:58px; height:58px; border-radius:14px; background:linear-gradient(180deg,#0099DA 0 33%,#FFD700 33% 66%,#1E2B7A 66%); display:flex; align-items:center; justify-content:center; font-size:28px; }
.kpi { background:var(--card); border:1px solid #e2e8f0; border-radius:16px; padding:16px 18px; box-shadow:0 4px 14px rgba(15,23,42,.04); }
.kpi .label { color:var(--muted); font-size:.82rem; margin-bottom:4px; }
.kpi .value { color:#0f172a; font-size:1.55rem; font-weight:800; }
.section { background:var(--card); border:1px solid #e2e8f0; border-radius:18px; padding:18px; margin:10px 0; box-shadow:0 3px 14px rgba(15,23,42,.035); }
</style>
''', unsafe_allow_html=True)

@st.cache_resource(show_spinner=False)
def carregar_base():
    config = AppConfig(); config.verificar_arquivos()
    loader = DataLoader(config); engine = RewardEngine()
    precos = loader.carregar_precos()
    frota, mapa_frota = loader.carregar_frota()
    cadastro = loader.carregar_cadastro_motoristas()
    abastecimentos = loader.carregar_abastecimentos(mapa_frota)
    eventos = engine.calcular_eventos_consumo(abastecimentos)
    return config, loader, engine, precos, frota, mapa_frota, cadastro, abastecimentos, eventos

config, loader, engine, precos, frota, mapa_frota, cadastro, abastecimentos, eventos = carregar_base()
if "ausencias" not in st.session_state: st.session_state.ausencias = carregar_ausencias()
if "desclassificacoes" not in st.session_state: st.session_state.desclassificacoes = carregar_desclassificacoes()
if "mapa_cat_custom" not in st.session_state: st.session_state.mapa_cat_custom = carregar_categorias_customizadas()

datas_validas = abastecimentos["DATA_FILTRO"].dropna() if "DATA_FILTRO" in abastecimentos.columns else eventos["DATA_NUM"].dropna()
min_dt = datas_validas.min().date() if not datas_validas.empty else date(2026,1,1)
max_dt = datas_validas.max().date() if not datas_validas.empty else date(2026,12,31)

def aplicar_filtros_st(dt_ini, dt_fim, motorista, placa, categoria, filial):
    di = dt_ini.strftime('%d/%m/%Y') if isinstance(dt_ini, date) else str(dt_ini)
    df = dt_fim.strftime('%d/%m/%Y') if isinstance(dt_fim, date) else str(dt_fim)
    return aplicar_filtros(di, df, motorista, placa, categoria, filial, st.session_state.ausencias, st.session_state.desclassificacoes, st.session_state.mapa_cat_custom)

def ausencia_label(i, row):
    return f"[{i}] {row.get('MOTORISTA','')} — {row.get('TIPO_AUSENCIA','')} — {row.get('DATA_INICIO','')} até {row.get('DATA_FIM','')} ({row.get('DIAS',0)} dias)"

def descl_label(i, row):
    return f"[{i}] {row.get('MOTORISTA','')} — {str(row.get('CRITERIO','')).split('[',1)[0].strip()}"

# Valores iniciais para os filtros
initial = aplicar_filtros_st(min_dt, max_dt, "TODOS", "", "TODAS", "TODAS")
res_initial = initial[-1]
mots_lista = ["TODOS"] + sorted(res_initial["MOTORISTA"].dropna().unique().tolist())
cats_lista = ["TODAS"] + sorted(res_initial["CATEGORIA"].dropna().unique().tolist())
filiais_lista = ["TODAS"] + sorted([str(x) for x in cadastro["BASE_CADASTRO"].dropna().unique() if str(x).strip()])

st.markdown('''<div class="hero"><div class="logo-bar"><div class="logo">🚚</div><div><h1>Dashboard do Prêmio de Motoristas</h1><p>Visão gerencial de consumo, desempenho e premiação</p></div></div></div>''', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 🔎 Filtros")
    dt_ini = st.date_input("Data início", value=min_dt, format="DD/MM/YYYY")
    dt_fim = st.date_input("Data fim", value=max_dt, format="DD/MM/YYYY")
    mot = st.selectbox("Motorista", mots_lista)
    placa = st.text_input("Placa")
    cat = st.selectbox("Categoria", cats_lista)
    filial = st.selectbox("Filial / Base", filiais_lista)
    if st.button("🔄 Limpar filtros", use_container_width=True): st.rerun()

current = aplicar_filtros_st(dt_ini, dt_fim, mot, placa, cat, filial)
(f_premio, f_gasto, f_km, f_litros, f_media, f_mots, res_view, rh_view, det_view, df_multi, mots_multi, res_f) = current

kpis=st.columns(6)
for c,(lab,val) in zip(kpis,[("💰 Total em Prêmios",f_premio),("⛽ Gasto Combustível",f_gasto),("📍 KM Rodados",f_km),("🧪 Litros",f_litros),("🎯 Média KM/L",f_media),("👥 Motoristas",f_mots)]):
    c.markdown(f'<div class="kpi"><div class="label">{lab}</div><div class="value">{val}</div></div>',unsafe_allow_html=True)

tabs=st.tabs(["📊 Resumo","⚙️ Cadastros","🚚 Múltiplas Placas","🏷️ Categorias por Placa","👔 Relatório RH","⛽ Abastecimentos","📄 Recibos","🏥 Ausências","🚫 Desclassificações"])
with tabs[0]:
    st.subheader("Resumo de Premiações")
    st.dataframe(res_view,use_container_width=True,hide_index=True)
with tabs[1]:
    st.subheader("Gestão de Cadastros")

    a, b = st.columns(2)
    with a:
        st.markdown("#### 👤 Novo motorista")
        n = st.text_input("Nome", key="cad_n")
        t = st.selectbox(
            "Tipo",
            sorted(precos["TIPO"].unique()) + ["FOLGUISTA"],
            key="cad_t",
        )
        base = st.text_input("Filial/Base", key="cad_b")
        if st.button("➕ Cadastrar motorista", key="cad_mot_btn") and n and base:
            df = carregar_motoristas_customizados()
            novo = pd.DataFrame([{
                "MOTORISTAS": DataUtils.normalizar_texto(n),
                "TIPO": DataUtils.normalizar_texto(t),
                "BASE": DataUtils.normalizar_texto(base),
            }])
            salvar_motoristas_customizados(pd.concat([df, novo], ignore_index=True))
            st.cache_resource.clear()
            st.rerun()

    with b:
        st.markdown("#### 🚛 Nova placa")
        p = st.text_input("Placa", key="cad_p")
        tp = st.selectbox(
            "Tipo do veículo",
            sorted(precos["TIPO"].unique()),
            key="cad_tp",
        )
        if st.button("➕ Cadastrar placa", key="cad_placa_btn") and p:
            df = carregar_frota_customizada()
            novo = pd.DataFrame([{
                "CAVALO": DataUtils.padronizar_placa(p),
                "TIPO": DataUtils.normalizar_texto(tp),
            }])
            salvar_frota_customizada(pd.concat([df, novo], ignore_index=True))
            st.cache_resource.clear()
            st.rerun()

    st.markdown("#### 🔄 Inativar / Reativar Cadastros")
    st.caption(
        "Motoristas inativados permanecem no histórico, mas o prêmio é desconsiderado. "
        "Placas inativadas permanecem no histórico e ficam marcadas como inativas."
    )

    inativos_atual = carregar_inativos()
    col_mot, col_placa = st.columns(2)

    with col_mot:
        st.markdown("##### 👤 Motorista")
        mot_cad_lista = sorted(
            cadastro["MOTORISTA_CADASTRO"].dropna().astype(str).unique().tolist()
        )
        mot_escolhido = (
            st.selectbox(
                "Selecionar motorista para inativar/reativar",
                mot_cad_lista,
                key="inativar_mot_sel",
            )
            if mot_cad_lista else None
        )
        mot_norm = DataUtils.normalizar_texto(mot_escolhido) if mot_escolhido else ""
        mot_esta_inativo = mot_norm in inativos_atual["MOTORISTA"]
        if mot_escolhido:
            st.write(f"Status atual: {'🔴 INATIVO' if mot_esta_inativo else '🟢 ATIVO'}")
            b1, b2 = st.columns(2)
            with b1:
                if st.button("❌ Inativar Motorista", key="btn_inativar_mot",
                             disabled=mot_esta_inativo, use_container_width=True):
                    alternar_inativo("MOTORISTA", mot_escolhido, True)
                    st.cache_resource.clear()
                    st.rerun()
            with b2:
                if st.button("✅ Reativar Motorista", key="btn_reativar_mot",
                             disabled=not mot_esta_inativo, use_container_width=True):
                    alternar_inativo("MOTORISTA", mot_escolhido, False)
                    st.cache_resource.clear()
                    st.rerun()

    with col_placa:
        st.markdown("##### 🚛 Placa")
        placas_cad_lista = sorted(
            frota["PLACA_PADRONIZADA"].dropna().astype(str).unique().tolist()
        )
        placa_escolhida = (
            st.selectbox(
                "Selecionar placa para inativar/reativar",
                placas_cad_lista,
                key="inativar_placa_sel",
            )
            if placas_cad_lista else None
        )
        placa_norm = DataUtils.padronizar_placa(placa_escolhida) if placa_escolhida else ""
        placa_esta_inativa = placa_norm in inativos_atual["PLACA"]
        if placa_escolhida:
            st.write(f"Status atual: {'🔴 INATIVA' if placa_esta_inativa else '🟢 ATIVA'}")
            b3, b4 = st.columns(2)
            with b3:
                if st.button("❌ Inativar Placa", key="btn_inativar_placa",
                             disabled=placa_esta_inativa, use_container_width=True):
                    alternar_inativo("PLACA", placa_escolhida, True)
                    st.cache_resource.clear()
                    st.rerun()
            with b4:
                if st.button("✅ Reativar Placa", key="btn_reativar_placa",
                             disabled=not placa_esta_inativa, use_container_width=True):
                    alternar_inativo("PLACA", placa_escolhida, False)
                    st.cache_resource.clear()
                    st.rerun()

    cadastro_exib = cadastro.copy()
    if "STATUS" in cadastro_exib.columns:
        cadastro_exib["STATUS"] = cadastro_exib["STATUS"].replace({
            "ATIVO": "🟢 ATIVO",
            "INATIVO": "🔴 INATIVO",
        })
    st.markdown("##### 👥 Motoristas cadastrados")
    st.dataframe(cadastro_exib, use_container_width=True, hide_index=True)

    frota_exib = frota.copy()
    if "STATUS" in frota_exib.columns:
        frota_exib["STATUS"] = frota_exib["STATUS"].replace({
            "ATIVO": "🟢 ATIVA",
            "INATIVO": "🔴 INATIVA",
        })
    st.markdown("##### 🚚 Placas cadastradas")
    st.dataframe(frota_exib, use_container_width=True, hide_index=True)

with tabs[2]:
    st.subheader("Média separada por placa")
    st.dataframe(df_multi,use_container_width=True,hide_index=True)
with tabs[3]:
    st.subheader("Categoria considerada para pagamento")
    mot_opts=sorted(res_f["MOTORISTA"].dropna().unique().tolist()) if not res_f.empty else []
    sm=st.selectbox("Motorista",mot_opts) if mot_opts else ""
    plate_opts=sorted(eventos.loc[eventos["CONDUTOR_NORMALIZADO"]==sm,"PLACA_PADRONIZADA"].dropna().unique().tolist()) if sm else []
    sp=st.selectbox("Placa",plate_opts) if plate_opts else ""
    sc=st.selectbox("Categoria",sorted(precos["TIPO"].unique()))
    if st.button("💾 Salvar categoria",key="savecat") and sm and sp:
        st.session_state.mapa_cat_custom[normalizar_chave_categoria_customizada(sm,sp)]=DataUtils.normalizar_texto(sc); salvar_categorias_customizadas(st.session_state.mapa_cat_custom); st.rerun()
    st.dataframe(pd.DataFrame(list(st.session_state.mapa_cat_custom.items()),columns=["MOTORISTA|||PLACA","CATEGORIA"]),use_container_width=True,hide_index=True)
with tabs[4]:
    st.subheader("Relatório RH")
    st.dataframe(rh_view,use_container_width=True,hide_index=True)
with tabs[5]:
    st.subheader("Detalhamento dos Abastecimentos")
    st.dataframe(det_view,use_container_width=True,hide_index=True)
with tabs[6]:
    st.subheader("Recibo de Premiação")
    rec_fil=st.selectbox("Filial",filiais_lista,key="rf")
    rec_mots=["TODOS OS MOTORISTAS (TODAS AS FILIAIS)"]+sorted(res_f["MOTORISTA"].dropna().unique().tolist()) if not res_f.empty else ["TODOS OS MOTORISTAS (TODAS AS FILIAIS)"]
    rec_mot=st.selectbox("Motorista",rec_mots,key="rm")
    rec_fac=st.text_input("Fator Carga",value="50%")
    if st.button("📄 Gerar recibo",key="rr"):
        html=gerar_recibos_lote(rec_fil,rec_mot,dt_ini.strftime('%d/%m/%Y'),dt_fim.strftime('%d/%m/%Y'),rec_fac,res_f)
        st.components.v1.html(html,height=900,scrolling=True)
with tabs[7]:
    st.subheader("Lançamento de Ausências")
    st.info("Competência: dia 26 até dia 25. O período é contado de forma inclusiva.")
    a_mots=sorted(cadastro["MOTORISTA_CADASTRO"].dropna().unique().tolist())
    if a_mots:
        am=st.selectbox("Motorista",a_mots,key="am"); at=st.radio("Tipo",["Atestado Médico","Férias","Outro Afastamento"],horizontal=True,key="at")
        ai=st.date_input("Data de Início",dt_ini,key="ai"); af=st.date_input("Data Fim",dt_fim,key="af"); ad=calcular_dias_ausencia(ai.strftime('%d/%m/%Y'),af.strftime('%d/%m/%Y')); st.metric("Dias Ausente (calculado)",ad); ao=st.text_input("Observação",key="ao")
        if st.button("➕ Lançar Ausência",key="aadd"):
            if ad<=0: st.error("Data final inválida")
            else:
                novo=pd.DataFrame([{"MOTORISTA":am,"TIPO_AUSENCIA":at,"DATA_INICIO":ai.strftime('%d/%m/%Y'),"DATA_FIM":af.strftime('%d/%m/%Y'),"DIAS":ad,"OBSERVACAO":ao}]); st.session_state.ausencias=pd.concat([st.session_state.ausencias,novo],ignore_index=True); salvar_ausencias(st.session_state.ausencias); st.rerun()
    st.dataframe(st.session_state.ausencias,use_container_width=True,hide_index=True)
    if not st.session_state.ausencias.empty:
        opts=[ausencia_label(i,r) for i,r in st.session_state.ausencias.reset_index(drop=True).iterrows()]; sel=st.selectbox("🗑️ Registro para excluir",opts,key="ax");
        if st.button("🗑️ Excluir registro selecionado",key="axx"):
            idx=int(sel.split(']')[0].replace('[','')); st.session_state.ausencias=st.session_state.ausencias.drop(index=idx).reset_index(drop=True); salvar_ausencias(st.session_state.ausencias); st.rerun()
with tabs[8]:
    st.subheader("Gestão de Desclassificações (Pilar 1)")
    d_mots=sorted(cadastro["MOTORISTA_CADASTRO"].dropna().unique().tolist())
    if d_mots:
        dm=st.selectbox("Motorista",d_mots,key="dm"); dc=st.selectbox("Critério / Infração",CRITERIOS_PILAR_1,key="dc"); dp=st.number_input("Pontos / Eventos",min_value=1,value=1); do=st.text_input("Observação",key="do")
        if st.button("➕ Lançar",key="dadd"):
            num=int(str(dc).split('-')[0].strip()) if '-' in str(dc) else 1; ti="DESCLASSIFICADO" if num>=5 else "PONTOS"; novo=pd.DataFrame([{"MOTORISTA":dm,"CRITERIO":dc,"PONTOS":dp,"TIPO_IMPACTO":ti,"OBSERVACAO":do}]); st.session_state.desclassificacoes=pd.concat([st.session_state.desclassificacoes,novo],ignore_index=True); salvar_desclassificacoes(st.session_state.desclassificacoes); st.rerun()
    st.dataframe(st.session_state.desclassificacoes,use_container_width=True,hide_index=True)
    if not st.session_state.desclassificacoes.empty:
        opts=[descl_label(i,r) for i,r in st.session_state.desclassificacoes.reset_index(drop=True).iterrows()]; sel=st.selectbox("🗑️ Registro para excluir",opts,key="dx");
        if st.button("🗑️ Excluir registro selecionado",key="dxx"):
            idx=int(sel.split(']')[0].replace('[','')); st.session_state.desclassificacoes=st.session_state.desclassificacoes.drop(index=idx).reset_index(drop=True); salvar_desclassificacoes(st.session_state.desclassificacoes); st.rerun()

