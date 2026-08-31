import os
import base64
import hmac
import hashlib
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime, date
from typing import Optional, Tuple
import textwrap

import streamlit as st
import numpy as np
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt


# ================================================================
# CONFIGURAÇÃO DE CAMINHOS E PERSISTÊNCIA (LOCAL / RAILWAY)
# ================================================================
DATA_DIR = os.getenv("DATA_DIR", ".")

# Arquivos mensais de viagens: exemplo "viagens 08_2026.xlsx".
# O aplicativo procura automaticamente todos os arquivos nesse padrão no mesmo diretório
# dos demais arquivos hospedeiros.
def _listar_arquivos_viagens():
  diretorio = DATA_DIR if DATA_DIR and DATA_DIR != "." else "."
  try:
    nomes = os.listdir(diretorio)
  except OSError:
    return []
  encontrados = []
  for nome in nomes:
    if re.fullmatch(r"viagens\s+\d{2}_\d{4}\.(xlsx|xlsm|xls)", nome, flags=re.IGNORECASE):
      encontrados.append(os.path.join(diretorio, nome))
  return sorted(encontrados, key=lambda x: os.path.basename(x).lower())

def _token_arquivos_viagens():
  token = []
  for caminho in _listar_arquivos_viagens():
    try:
      token.append((caminho, os.path.getmtime(caminho), os.path.getsize(caminho)))
    except OSError:
      token.append((caminho, 0.0, 0))
  return tuple(token)


def _listar_arquivos_velocidade():
  diretorio = DATA_DIR if DATA_DIR and DATA_DIR != "." else "."
  try:
    nomes = os.listdir(diretorio)
  except OSError:
    return []
  encontrados = []
  padrao = r"Extrato\s+de\s+Velocidade\s+excedida\s+\d{2}-\d{2}\s+a\s+\d{2}-\d{2}(?:\s+\d{4})?\.(xlsx|xlsm|xls)$"
  for nome in nomes:
    if re.fullmatch(padrao, nome, flags=re.IGNORECASE):
      encontrados.append(os.path.join(diretorio, nome))
  return sorted(encontrados, key=lambda x: os.path.basename(x).lower())

def _token_arquivos_velocidade():
  token = []
  for caminho in _listar_arquivos_velocidade():
    try:
      token.append((caminho, os.path.getmtime(caminho), os.path.getsize(caminho)))
    except OSError:
      token.append((caminho, 0.0, 0))
  return tuple(token)

def _extrair_periodo_nome_velocidade(nome_arquivo):
  nome = os.path.basename(nome_arquivo)
  m = re.search(r"(\d{2})-(\d{2})\s+a\s+(\d{2})-(\d{2})", nome)
  if not m:
    return None
  return int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))

def carregar_excesso_velocidade_automatico(dt_ini, dt_fim):
  """Lê o extrato mensal correspondente à competência selecionada.

  A planilha possui a aba 'Geral' com Motorista, Filial e Picos de Velocidade.
  A competência é identificada pelos dias/meses no nome do arquivo, então o ano
  pode variar sem exigir alteração do código.
  """
  try:
    di = pd.Timestamp(dt_ini).date()
    df = pd.Timestamp(dt_fim).date()
  except Exception:
    return pd.DataFrame()

  alvo = (di.day, di.month, df.day, df.month)
  arquivos = []
  for caminho in _listar_arquivos_velocidade():
    periodo = _extrair_periodo_nome_velocidade(caminho)
    if periodo == alvo:
      arquivos.append(caminho)

  if not arquivos:
    return pd.DataFrame()

  caminho = arquivos[-1]
  try:
    bruto = pd.read_excel(caminho, sheet_name="Geral", engine="openpyxl", header=None, dtype=object)
  except Exception as exc:
    print(f"Erro ao ler extrato de velocidade {os.path.basename(caminho)}: {exc}")
    return pd.DataFrame()

  if bruto.empty:
    return pd.DataFrame()

  # Localiza a linha de cabeçalho que contém Motoristas/Filial/Picos de Velocidade.
  cab_idx = None
  cab = None
  for i in range(min(len(bruto), 20)):
    vals = [DataUtils.normalizar_texto(v) for v in bruto.iloc[i].tolist()]
    if "MOTORISTAS" in vals and "FILIAL" in vals and "PICOS DE VELOCIDADE" in vals:
      cab_idx = i
      cab = vals
      break
  if cab_idx is None:
    return pd.DataFrame()

  idx_mot = cab.index("MOTORISTAS")
  idx_fil = cab.index("FILIAL")
  idx_evt = cab.index("PICOS DE VELOCIDADE")

  dados = bruto.iloc[cab_idx + 1 :].copy()
  out = pd.DataFrame({
      "MOTORISTA": dados.iloc[:, idx_mot].apply(DataUtils.normalizar_texto),
      "FILIAL": dados.iloc[:, idx_fil].apply(DataUtils.normalizar_texto),
      "EVENTOS": pd.to_numeric(dados.iloc[:, idx_evt], errors="coerce").fillna(0),
  })
  out["EVENTOS"] = out["EVENTOS"].astype(int)
  out = out[(out["MOTORISTA"] != "") & (out["EVENTOS"] > 0)].copy()
  if out.empty:
    return pd.DataFrame()

  out["CATEGORIA"] = ""
  out["DATA_EVENTO"] = pd.Timestamp(df).strftime("%d/%m/%Y")
  out["OBSERVACAO"] = "Importado automaticamente do extrato mensal"
  out["FONTE_EVENTO"] = "EXTRATO_VELOCIDADE"
  out["ARQUIVO_FONTE"] = os.path.basename(caminho)
  return out[["MOTORISTA","CATEGORIA","DATA_EVENTO","EVENTOS","OBSERVACAO","FONTE_EVENTO","ARQUIVO_FONTE"]].reset_index(drop=True)

ARQUIVO_AUSENCIAS = os.path.join(DATA_DIR, "ausencias.csv")
ARQUIVO_DESCLASSIFICACOES = os.path.join(DATA_DIR, "desclassificacoes.csv")
ARQUIVO_CATEGORIAS_CUSTOM = os.path.join(DATA_DIR, "categorias_customizadas.csv")
ARQUIVO_CATEGORIAS_VIGENCIA = os.path.join(DATA_DIR, "categorias_customizadas_vigencia.csv")
ARQUIVO_FROTA_CUSTOM = os.path.join(DATA_DIR, "frota_customizada.csv")
ARQUIVO_MOTORISTAS_CUSTOM = os.path.join(DATA_DIR, "motoristas_customizados.csv")
ARQUIVO_INATIVOS = os.path.join(DATA_DIR, "inativos.csv")
ARQUIVO_DATAS_MOTORISTAS = os.path.join(DATA_DIR, "datas_motoristas.csv")
ARQUIVO_CODIGOS_FUNCIONAIS = os.path.join(DATA_DIR, "codigos_funcionais.csv")
ARQUIVO_EXCESSO_VELOCIDADE = os.path.join(DATA_DIR, "excesso_velocidade.csv")
ARQUIVO_CONTROLE_JORNADA = os.path.join(DATA_DIR, "controle_jornada.csv")

PADRAO_ABASTECIMENTOS_MENSAL = r"abastecimentos[_\s-]?(\d{2})[_-]?(\d{4})\.(xlsx|xlsm|xls)$"
ARQUIVO_ABASTECIMENTOS_LEGADO = os.path.join(DATA_DIR, "uah_abastecimentos_3.xlsx")

def _listar_arquivos_abastecimentos_mensais():
  diretorio = DATA_DIR if DATA_DIR and DATA_DIR != "." else "."
  try:
    nomes = os.listdir(diretorio)
  except OSError:
    return []
  encontrados = []
  for nome in nomes:
    m = re.fullmatch(PADRAO_ABASTECIMENTOS_MENSAL, nome, flags=re.IGNORECASE)
    if m:
      encontrados.append((int(m.group(1)), int(m.group(2)), os.path.join(diretorio, nome)))
  return sorted(encontrados, key=lambda x: (x[1], x[0], os.path.basename(x[2]).lower()))

def _token_arquivos_abastecimentos_mensais():
  token=[]
  for mm, yyyy, caminho in _listar_arquivos_abastecimentos_mensais():
    try: token.append((caminho, os.path.getmtime(caminho), os.path.getsize(caminho)))
    except OSError: token.append((caminho,0.0,0))
  try:
    token.append((ARQUIVO_ABASTECIMENTOS_LEGADO, os.path.getmtime(ARQUIVO_ABASTECIMENTOS_LEGADO), os.path.getsize(ARQUIVO_ABASTECIMENTOS_LEGADO)))
  except OSError:
    token.append((ARQUIVO_ABASTECIMENTOS_LEGADO,0.0,0))
  return tuple(token)

def _arquivo_historico_abastecimentos(mm, yyyy):
  return os.path.join(DATA_DIR, f"historico_abastecimentos_{mm:02d}{yyyy:04d}.csv")

def _competencia_do_mes_abastecimento(mm, yyyy):
  fim = pd.Timestamp(year=yyyy, month=mm, day=25)
  inicio = (fim - pd.DateOffset(months=1)).replace(day=26)
  return inicio.normalize(), fim.normalize()

def _ler_planilha_abastecimentos_generica(caminho):
  try:
    return pd.read_excel(caminho, sheet_name=0, dtype=str, keep_default_na=False)
  except Exception as exc:
    print(f"Erro ao ler abastecimentos {os.path.basename(caminho)}: {exc}")
    return pd.DataFrame()

def _normalizar_base_abastecimentos_raw(df, mapa_frota):
  if df is None or df.empty:
    return pd.DataFrame()
  col_placa = DataUtils.encontrar_coluna(df, ["PLACA", "CAVALO"])
  col_km = DataUtils.encontrar_coluna(df, ["KM ATUAL", "KM", "KM_1", "QUILOMETRAGEM"])
  col_litros = DataUtils.encontrar_coluna(df, ["QTDE", "LITROS", "QUANTIDADE", "QTD"])
  col_valor = DataUtils.encontrar_coluna(df, ["VALOR TOTAL","VALOR","TOTAL","VALOR_TOTAL","VR TOTAL","VLR TOTAL","VALOR COMBUSTIVEL","VALOR (R$)"])
  col_motorista = DataUtils.encontrar_coluna(df, ["CONDUTOR", "MOTORISTA", "MOTORISTAS"])
  col_data = DataUtils.encontrar_coluna(df, ["DATA","Data","DATA ABASTECIMENTO","DATA_ABASTECIMENTO","DATA DO ABASTECIMENTO","DATA/HORA","DATA HORA","DATA EMISSAO","DT_ABASTECIMENTO","DT ABAST"])
  col_hora = DataUtils.encontrar_coluna(df, ["HORA","Hora","HORARIO","HORA ABASTECIMENTO"])
  if any(c is None for c in [col_placa,col_km,col_litros,col_motorista,col_data]):
    return pd.DataFrame()
  out=df.copy()
  out["_ORDEM_ORIGINAL"] = np.arange(len(out))
  out["PLACA_PADRONIZADA"]=out[col_placa].apply(DataUtils.padronizar_placa)
  out["KM_ATUAL_NUM"]=out[col_km].apply(DataUtils.converter_numero)
  out["QTDE_NUM"]=out[col_litros].apply(DataUtils.converter_numero)
  out["VALOR_NUM"]=out[col_valor].apply(DataUtils.converter_numero).fillna(0.0) if col_valor else 0.0
  out["CONDUTOR_NORMALIZADO"]=out[col_motorista].fillna("SEM MOTORISTA").apply(DataUtils.normalizar_texto)
  out["DATA_ORIGINAL"]=out[col_data]
  out["DATA_FILTRO"]=out[col_data].apply(criar_data_filtro)
  out["DATA_NUM"]=out["DATA_FILTRO"]
  out["DATA"]=out["DATA_FILTRO"]
  out["DATA_HORA_ABASTECIMENTO"]=out[col_data].apply(_parse_datetime_flex)
  if col_hora:
    hs=out[col_hora].apply(_parse_datetime_flex)
    mask=hs.notna() & out["DATA_HORA_ABASTECIMENTO"].notna()
    out.loc[mask,"DATA_HORA_ABASTECIMENTO"]=out.loc[mask,"DATA_HORA_ABASTECIMENTO"].dt.normalize()+pd.to_timedelta(hs.loc[mask].dt.hour*3600+hs.loc[mask].dt.minute*60+hs.loc[mask].dt.second,unit="s")
  out["TIPO"]=out["PLACA_PADRONIZADA"].map(mapa_frota)
  sem=out["TIPO"].isna()
  for idx in out.index[sem]:
    eq=DataUtils.placa_equivalente_mercosul(out.at[idx,"PLACA_PADRONIZADA"])
    if eq and eq in mapa_frota: out.at[idx,"TIPO"]=mapa_frota[eq]
  out["REGISTRO_VALIDO"]=(out["PLACA_PADRONIZADA"]!="") & out["KM_ATUAL_NUM"].notna() & (out["KM_ATUAL_NUM"]>0) & out["QTDE_NUM"].notna() & (out["QTDE_NUM"]>0)
  return out

def _colunas_historico_abastecimentos():
  return ["ID_ABASTECIMENTO","_ORDEM_ORIGINAL","PLACA_PADRONIZADA","KM_ATUAL_NUM","QTDE_NUM","VALOR_NUM","CONDUTOR_NORMALIZADO","DATA_FILTRO","DATA_NUM","DATA","DATA_HORA_ABASTECIMENTO","TIPO","REGISTRO_VALIDO","MOTORISTA_ABASTECIMENTO_ORIGINAL"]

def _arquivar_abastecimentos_mensais(mapa_frota):
  garantir_diretorio()
  cols = _colunas_historico_abastecimentos()

  mensal_processados = set()
  for mm, yyyy, caminho in _listar_arquivos_abastecimentos_mensais():
    raw = _ler_planilha_abastecimentos_generica(caminho)
    norm = _normalizar_base_abastecimentos_raw(raw, mapa_frota)
    if norm.empty or "DATA_FILTRO" not in norm.columns:
      continue

    datas = pd.to_datetime(norm["DATA_FILTRO"], errors="coerce").dt.normalize()
    fim_comp = pd.Timestamp(year=yyyy, month=mm, day=25).normalize()
    ini_comp = (fim_comp - pd.DateOffset(months=1)).replace(day=26).normalize()
    norm = norm.loc[datas.notna() & (datas >= ini_comp) & (datas <= fim_comp)].copy()
    if norm.empty:
      continue

    norm["ID_ABASTECIMENTO"] = [
        f"{yyyy:04d}{mm:02d}-{int(i):06d}"
        for i in norm["_ORDEM_ORIGINAL"].fillna(0).astype(int)
    ]
    norm["MOTORISTA_ABASTECIMENTO_ORIGINAL"] = norm["CONDUTOR_NORMALIZADO"]

    for c in cols:
      if c not in norm.columns:
        norm[c] = ""
    part = norm[cols].copy()

    arq = _arquivo_historico_abastecimentos(mm, yyyy)
    part.to_csv(arq, index=False, encoding="utf-8-sig")
    mensal_processados.add((mm, yyyy))

  if os.path.exists(ARQUIVO_ABASTECIMENTOS_LEGADO):
    raw = _ler_planilha_abastecimentos_generica(ARQUIVO_ABASTECIMENTOS_LEGADO)
    norm = _normalizar_base_abastecimentos_raw(raw, mapa_frota)
    if not norm.empty and "DATA_FILTRO" in norm.columns:
      datas = pd.to_datetime(norm["DATA_FILTRO"], errors="coerce").dt.normalize()
      norm = norm.loc[datas.notna()].copy()
      if not norm.empty:
        fim_comp = datas.apply(
            lambda d: ((d + pd.DateOffset(months=1)).replace(day=25)
                       if pd.notna(d) and d.day >= 26
                       else (d.replace(day=25) if pd.notna(d) else pd.NaT))
        )
        comp_keys = fim_comp.dt.strftime("%m%Y")
        for key in comp_keys.dropna().unique():
          mm, yyyy = int(key[:2]), int(key[2:])
          if (mm, yyyy) in mensal_processados:
            continue
          arq = _arquivo_historico_abastecimentos(mm, yyyy)
          if os.path.exists(arq):
            continue
          mask = comp_keys == key
          part = norm.loc[mask].copy()
          if part.empty:
            continue
          part["ID_ABASTECIMENTO"] = [
              f"LEGACY-{yyyy:04d}{mm:02d}-{int(i):06d}"
              for i in part["_ORDEM_ORIGINAL"].fillna(0).astype(int)
          ]
          part["MOTORISTA_ABASTECIMENTO_ORIGINAL"] = part["CONDUTOR_NORMALIZADO"]
          for c in cols:
            if c not in part.columns:
              part[c] = ""
          part[cols].to_csv(arq, index=False, encoding="utf-8-sig")

def _carregar_historico_abastecimentos_total():
  registros=[]
  diretorio = DATA_DIR if DATA_DIR and DATA_DIR!="." else "."
  for nome in os.listdir(diretorio):
    m=re.fullmatch(r"historico_abastecimentos_(\d{2})(\d{4})\.csv",nome,re.IGNORECASE)
    if not m:
      continue
    mm, yyyy = int(m.group(1)), int(m.group(2))
    caminho=os.path.join(diretorio,nome)
    try:
      df=pd.read_csv(caminho,dtype=str,encoding="utf-8-sig")
      for c in _colunas_historico_abastecimentos():
        if c not in df.columns: df[c]=""
      df["_ORDEM_ORIGINAL"] = pd.to_numeric(df["_ORDEM_ORIGINAL"], errors="coerce").fillna(
          pd.Series(range(len(df)), index=df.index)
      ).astype(int)
      idv = df["ID_ABASTECIMENTO"].fillna("").astype(str).str.strip()
      falt = idv.eq("")
      if falt.any():
        df.loc[falt, "ID_ABASTECIMENTO"] = [
            f"{yyyy:04d}{mm:02d}-{i:06d}" for i in df.index[falt]
        ]
      for c in ["KM_ATUAL_NUM","QTDE_NUM","VALOR_NUM"]:
        df[c]=df[c].apply(DataUtils.converter_numero)
      df["REGISTRO_VALIDO"]=df["REGISTRO_VALIDO"].astype(str).str.lower().isin(["true","1","sim"])
      df["DATA_FILTRO"]=df["DATA_FILTRO"].apply(parse_data_filtro)
      df["DATA_NUM"]=df["DATA_FILTRO"]
      df["DATA"]=df["DATA_FILTRO"]
      df["DATA_HORA_ABASTECIMENTO"]=df["DATA_HORA_ABASTECIMENTO"].apply(_parse_datetime_flex)
      registros.append(df[_colunas_historico_abastecimentos()])
    except Exception as exc:
      print(f"Erro ao carregar {nome}: {exc}")
  if not registros:
    return pd.DataFrame()
  return pd.concat(registros,ignore_index=True)

ARQUIVO_USUARIOS_ACESSO = os.path.join(DATA_DIR, "usuarios_acesso.csv")

VALOR_PONTO_POR_CATEGORIA = {
    "TRUCK": 1.40,
    "BITRUCK": 1.63,
    "CARRETA": 1.87,
    "BITREM": 2.10,
    "RODOTREM": 2.45,
    "RODOENTREGA": 2.45,
}

def garantir_diretorio():
  if DATA_DIR and DATA_DIR != "." and not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR, exist_ok=True)

def carregar_ausencias() -> pd.DataFrame:
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
  try:
    garantir_diretorio()
    df.to_csv(ARQUIVO_AUSENCIAS, index=False, encoding="utf-8-sig")
  except Exception as e:
    print(f"Erro ao salvar ausências: {e}")

def carregar_desclassificacoes() -> pd.DataFrame:
  colunas = ["MOTORISTA", "CRITERIO", "PONTOS", "TIPO_IMPACTO", "DATA_EVENTO", "OBSERVACAO"]
  if os.path.exists(ARQUIVO_DESCLASSIFICACOES):
    try:
      df = pd.read_csv(
          ARQUIVO_DESCLASSIFICACOES, dtype=str, encoding="utf-8-sig"
      )
      if "PONTOS" in df.columns:
        df["PONTOS"] = pd.to_numeric(df["PONTOS"], errors="coerce").fillna(1)
      else:
        df["PONTOS"] = 1
      if "DATA_EVENTO" not in df.columns:
        df["DATA_EVENTO"] = ""
      if "OBSERVACAO" not in df.columns:
        df["OBSERVACAO"] = ""
      if "MOTORISTA" not in df.columns:
        df["MOTORISTA"] = ""
      if "CRITERIO" not in df.columns:
        df["CRITERIO"] = ""
      if "TIPO_IMPACTO" not in df.columns:
        df["TIPO_IMPACTO"] = "PONTOS"
      for c in colunas:
        if c not in df.columns:
          df[c] = ""
      return df[colunas]
    except Exception as e:
      print(f"Erro ao carregar desclassificações: {e}")
  return pd.DataFrame(columns=colunas)

def salvar_desclassificacoes(df: pd.DataFrame):
  try:
    garantir_diretorio()
    df.to_csv(ARQUIVO_DESCLASSIFICACOES, index=False, encoding="utf-8-sig")
  except Exception as e:
    print(f"Erro ao salvar desclassificações: {e}")

def normalizar_chave_categoria_customizada(motorista: str, placa: str = "") -> str:
  mot = DataUtils.normalizar_texto(motorista)
  plc = DataUtils.padronizar_placa(placa)
  return f"{mot}|||{plc}" if plc else mot

def carregar_categorias_customizadas() -> dict:
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

def carregar_categorias_vigencia() -> pd.DataFrame:
  cols = ["MOTORISTA_CHAVE", "CATEGORIA_ESCOLHIDA", "DATA_INICIO", "DATA_FIM"]
  if os.path.exists(ARQUIVO_CATEGORIAS_VIGENCIA):
    try:
      df = pd.read_csv(ARQUIVO_CATEGORIAS_VIGENCIA, dtype=str, encoding="utf-8-sig").fillna("")
      for c in cols:
        if c not in df.columns: df[c] = ""
      return df[cols]
    except Exception as e: print(f"Erro ao carregar vigências de categorias: {e}")
  legado = carregar_categorias_customizadas()
  if not legado: return pd.DataFrame(columns=cols)
  try: dt_ref = pd.Timestamp(datetime.fromtimestamp(os.path.getmtime(ARQUIVO_CATEGORIAS_CUSTOM))).normalize()
  except Exception: dt_ref = pd.Timestamp.now().normalize()
  if dt_ref.day >= 26:
    ini = dt_ref.replace(day=26); fim = (ini + pd.DateOffset(months=1)).replace(day=25)
  else:
    fim = dt_ref.replace(day=25); ini = (fim - pd.DateOffset(months=1)).replace(day=26)
  df = pd.DataFrame([{"MOTORISTA_CHAVE":str(k).strip().upper(),"CATEGORIA_ESCOLHIDA":DataUtils.normalizar_texto(v),"DATA_INICIO":ini.strftime("%d/%m/%Y"),"DATA_FIM":fim.strftime("%d/%m/%Y")} for k,v in legado.items() if str(k).strip() and str(v).strip()])
  try: df.to_csv(ARQUIVO_CATEGORIAS_VIGENCIA,index=False,encoding="utf-8-sig")
  except Exception as e: print(f"Erro ao criar arquivo de vigência de categorias: {e}")
  return df[cols] if not df.empty else pd.DataFrame(columns=cols)

def salvar_categorias_vigencia(df: pd.DataFrame):
  try:
    garantir_diretorio(); cols=["MOTORISTA_CHAVE","CATEGORIA_ESCOLHIDA","DATA_INICIO","DATA_FIM"]; out=df.copy()
    for c in cols:
      if c not in out.columns: out[c] = ""
    out[cols].fillna("").to_csv(ARQUIVO_CATEGORIAS_VIGENCIA,index=False,encoding="utf-8-sig")
  except Exception as e: print(f"Erro ao salvar vigências de categorias: {e}")

def categorias_ativas_na_competencia(df_vig: pd.DataFrame, data_ini, data_fim) -> dict:
  if df_vig is None or df_vig.empty: return {}
  ini_c=pd.Timestamp(data_ini).normalize(); fim_c=pd.Timestamp(data_fim).normalize(); ativos={}
  for _,row in df_vig.iterrows():
    chave=str(row.get("MOTORISTA_CHAVE","")).strip().upper(); cat=DataUtils.normalizar_texto(row.get("CATEGORIA_ESCOLHIDA",""))
    if not chave or not cat: continue
    di=parse_data_filtro(row.get("DATA_INICIO","")); df=parse_data_filtro(row.get("DATA_FIM",""))
    if di is None: di=ini_c
    if df is None: df=fim_c
    di=pd.Timestamp(di).normalize(); df=pd.Timestamp(df).normalize()
    if di <= fim_c and df >= ini_c: ativos[chave]=cat
  return ativos

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

LEGACY_CADASTRO_SNAPSHOT = [{'MOTORISTAS': 'ADEILSON DE OLIVEIRA ANGELINO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'AIRTON ANTONIO GONÇALVES', 'TIPO': 'BITRUCK', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'CLAUDINEI FRANCISCO FERREIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'CLAUDIO JOSE KREGENSKI', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'DANILO CASSIANO FERREIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'DIEISON APARECIDO DA CRUZ', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'DOUGLAS ENRIQUE DA SILVA LUIZ', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EDILSON LEITE DE CAMARGO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EDINEI MARCOS CORDEIRO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EDISON VIEIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EDSON RECOFKA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EMERSON APARECIDO PEREIRA DA SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'FABIANO CASTILHO CALEGARI', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'GEDIVALDO SOUZA LUZ ALVES', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'GILMAR LOPACINSKI', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'INACIO DOUTOR', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JOAO PAULO LISNIOWSKI', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JONAS GOGOLA DE ANDRADE', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JOSIVAN DA SILVA OLIVEIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JOSUE LOPES DE SENE', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JULIANA COQUES PAZ', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'LEOMAR MOREIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'LIDIOMAR DA SILVA DE SOUZA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'MARCELO DA SILVA E SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'MARCIO LEMOS MACHADO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'NELSON SOBOTHE', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'NILSON APARECIDO SAMPAIO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'NILSON RODRIGUES DE SOUZA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'NILTON DE JESUS RODRIGUES DE SOUZA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'ODAIR GONÇALVES MIRANDA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'PAULO DE MELO SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'PEDRO VANDERLEI BRASILINO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'RICARDO SERGIO DA SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'RODRIGO DE SOUZA MACHADO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'VALDECI CARVALHO DA SILVA JUNIOR', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'VALDECI FERREIRA DA SILVA JUNIOR', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'WANDERLEY LOPES SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'FELIPE TELES DA CRUZ', 'TIPO': 'BITRUCK', 'BASE': 'CAMPO GRANDE'}, {'MOTORISTAS': 'RENATO RIEFF MARIN', 'TIPO': 'CARRETA', 'BASE': 'CAMPO GRANDE'}, {'MOTORISTAS': 'ELICAR JUSTINO', 'TIPO': 'TRUCK', 'BASE': 'CHAPECO'}, {'MOTORISTAS': 'ANTONIO CARLOS BRAMBILA', 'TIPO': 'BITRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'JHONATAN ALVES DOS SANTOS', 'TIPO': 'BITRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'LINCOLN FRANCEL PIMENTA', 'TIPO': 'BITRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'RODRIGO LORENTINO', 'TIPO': 'BITRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'APARECIDO DIAMARAES', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'APARECIDO JOEL SANT ANA', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'APARECIDO RODRIGUES DA SILVA', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'CARLOS ELIER PIEROLI', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'ELISANGELA APARECIDA GOMES COELHO', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'FELIPE COMAR DIAS', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'MAURILIO FERREIRA DAS NEVES', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'RODOLFO MOZELLI SPAGOLLA', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'VALBER JUNIOR COSTA', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'WESLEI RIBEIRO JACOMINI', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'GILBERTO BEZERRA PINTO', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'JOSE CARLOS RODRIGUES', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'JOSE DOS SANTOS', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'NIVALMIR ANTUNES', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'REGINALDO MENDES OLIVEIRA', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'SERGIO APARECIDO GIRALDELLO', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'VILSON TOMACHAK', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'DENNER DOS SANTOS', 'TIPO': 'TRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'DIEGO FRANCISCO DE SOUZA', 'TIPO': 'TRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'ALEX DOUGLAS LOPES ALONSO', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'ANDERSON DE SOUZA SOARES GOMES', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'ANDERSON NUBIATO RODRIGUES DA SILVA', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'ANGELA MARIA GONÇALVES', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'ANTONIO ROBERTO BELTRAMINI', 'TIPO': 'CARRETA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'CELSO RICARDO RODRIGUES', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'CRISTIAN FABIANO LUIZ DA SILVA', 'TIPO': 'TOCO', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'EDE WILSON RODRIGUES', 'TIPO': 'CARRETA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'EDVALDO GONCALVES', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'FABIO CARLOS ARAUJO DO CARMO', 'TIPO': 'CARRETA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'FERNANDO EMIDIO DE SOUZA LIMA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'GEDIELCIO CARVALHO COSTA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'GILMAR DA SILVA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'GILMAR FERREIRA NEVES', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'GUSTAVO ROBERTO PEREIRA', 'TIPO': 'CARRETA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'JHONE GIMENES SANTOS', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'JOAO VITOR DOS SANTOS', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'JOSE NILSON MARTINS DE ARAUJO', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'LEANDRO DE OLIVEIRA FERREIRA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'LUIS HENRIQUE SANTIAGO FIALHO', 'TIPO': 'FOLGUISTA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'MICHEL ANTONIOLI', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'PAULO CESAR VICENTINI', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'TATIANE CAXIMIRO PEREIRA', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'VALDINEY FERREIRA PRIMO', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'WESLEY ZANETTI DE OLIVEIRA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'WILLIAM ANDRADE DE MOURA', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'FRANCISCO DAS CHAGAS CORREA CRISPIM', 'TIPO': 'TRUCK', 'BASE': 'ITAJAI'}, {'MOTORISTAS': 'ROGERIO FRANÇA DOS SANTOS', 'TIPO': 'TRUCK', 'BASE': 'ITAJAI'}, {'MOTORISTAS': 'SILVANO DA SILVA FREITAS', 'TIPO': 'TRUCK', 'BASE': 'ITAJAI'}, {'MOTORISTAS': 'ANTONIO APARECIDO PEREIRA', 'TIPO': 'TRUCK', 'BASE': 'PAULINIA'}, {'MOTORISTAS': 'JOSE AUGUSTO DOS SANTOS', 'TIPO': 'TRUCK', 'BASE': 'PAULINIA'}, {'MOTORISTAS': 'RENATO PEREIRA FRANÇA', 'TIPO': 'RODOTREM', 'BASE': 'PAULINIA'}, {'MOTORISTAS': 'AGUINALDO DOS SANTOS TEIXEIRA', 'TIPO': 'RODO ENTREGA', 'BASE': 'SAO JOSE DOS CAMPOS'}, {'MOTORISTAS': 'KERLEI MIRANDA MARTINS', 'TIPO': 'TRUCK', 'BASE': 'SAO JOSE DOS CAMPOS'}, {'MOTORISTAS': 'TADEU JOSE CAETANO DE SOUZA', 'TIPO': 'TRUCK', 'BASE': 'SAO JOSE DOS CAMPOS'}, {'MOTORISTAS': 'RONAN ROMULO ANTUNES', 'TIPO': 'RODO ENTREGA', 'BASE': 'SAO JOSE DOS CAMPOS'}, {'MOTORISTAS': 'SIDNEI DE OLIVEIRA MARIANO', 'TIPO': 'CARRETA', 'BASE': 'SARANDI'}, {'MOTORISTAS': 'NIVALDO REIS MACHADO', 'TIPO': 'BITRUCK', 'BASE': 'UBERABA'}, {'MOTORISTAS': 'SIDNEY RODRIGUES FERREIRA', 'TIPO': 'BITRUCK', 'BASE': 'UBERABA'}, {'MOTORISTAS': 'WELLINGTON DE MELO BATISTA', 'TIPO': 'BITRUCK', 'BASE': 'UBERABA'}, {'MOTORISTAS': 'HIGOR GABRIEL OLIVEIRA BITU', 'TIPO': 'BITRUCK', 'BASE': 'UBERLANDIA'}, {'MOTORISTAS': 'JOSE DONIZETE FERREIRA GOMES', 'TIPO': 'BITRUCK', 'BASE': 'UBERLANDIA'}, {'MOTORISTAS': 'ANTONIO JOSE DE SOUZA MARTINS', 'TIPO': 'BITREM', 'BASE': 'VARZEA GRANDE'}, {'MOTORISTAS': 'JORGE SANTOS DA SILVA', 'TIPO': 'RODOTREM', 'BASE': 'VARZEA GRANDE'}, {'MOTORISTAS': 'MARCOS ROBERTO DOS SANTOS', 'TIPO': 'RODOTREM', 'BASE': 'VARZEA GRANDE'}, {'MOTORISTAS': 'OTAVIO ROSA FRANCO', 'TIPO': 'RODOTREM', 'BASE': 'VARZEA GRANDE'}]

ALIASES_MOTORISTAS_MIGRACAO = {
    "WANDERLEY LOPES DA SILVA": "WANDERLEY LOPES SILVA",
    "MARCOS ROBERTO DOS SANTOS ROSA": "MARCOS ROBERTO DOS SANTOS",
}

def _chave_migracao_motorista(nome: str) -> str:
    nome_n = DataUtils.normalizar_texto(nome)
    return ALIASES_MOTORISTAS_MIGRACAO.get(nome_n, nome_n)

def migrar_cadastro_legado_uma_vez() -> int:
    try:
        atual = carregar_motoristas_customizados().copy()
        for c in ["MOTORISTAS", "TIPO", "BASE"]:
            if c not in atual.columns:
                atual[c] = ""
        atual = atual[["MOTORISTAS", "TIPO", "BASE"]].copy()
        atual["MOTORISTAS"] = atual["MOTORISTAS"].apply(DataUtils.normalizar_texto)
        atual["TIPO"] = atual["TIPO"].apply(DataUtils.normalizar_texto).replace({"TOCO":"TRUCK"})
        atual["BASE"] = atual["BASE"].apply(DataUtils.normalizar_texto)
        atual = atual[atual["MOTORISTAS"] != ""].copy()

        legado = pd.DataFrame(LEGACY_CADASTRO_SNAPSHOT, columns=["MOTORISTAS", "TIPO", "BASE"])
        legado["MOTORISTAS"] = legado["MOTORISTAS"].apply(DataUtils.normalizar_texto)
        legado["TIPO"] = legado["TIPO"].apply(DataUtils.normalizar_texto).replace({"TOCO":"TRUCK"})
        legado["BASE"] = legado["BASE"].apply(DataUtils.normalizar_texto)

        mesclado = {}
        for _, r in legado.iterrows():
            chave = _chave_migracao_motorista(r["MOTORISTAS"])
            if chave:
                mesclado[chave] = {
                    "MOTORISTAS": r["MOTORISTAS"],
                    "TIPO": r["TIPO"],
                    "BASE": r["BASE"],
                }

        for _, r in atual.iterrows():
            chave = _chave_migracao_motorista(r["MOTORISTAS"])
            if chave:
                mesclado[chave] = {
                    "MOTORISTAS": r["MOTORISTAS"],
                    "TIPO": r["TIPO"],
                    "BASE": r["BASE"],
                }

        final = pd.DataFrame(list(mesclado.values()), columns=["MOTORISTAS","TIPO","BASE"])
        final = final.sort_values("MOTORISTAS", kind="stable").reset_index(drop=True)

        antes = len(atual)
        if len(final) > antes or antes < len(LEGACY_CADASTRO_SNAPSHOT):
            salvar_motoristas_customizados(final)
            return len(final)
        return antes
    except Exception as e:
        print(f"Erro na migração inicial do cadastro: {e}")
        return len(carregar_motoristas_customizados())

def _carregar_eventos_pilar(caminho: str) -> pd.DataFrame:
    colunas = ["MOTORISTA", "CATEGORIA", "DATA_EVENTO", "EVENTOS", "OBSERVACAO"]
    if os.path.exists(caminho):
        try:
            df = pd.read_csv(caminho, dtype=str, encoding="utf-8-sig")
            for c in colunas:
                if c not in df.columns:
                    df[c] = ""
            df = df[colunas].copy()
            df["EVENTOS"] = pd.to_numeric(df["EVENTOS"], errors="coerce").fillna(0).astype(int)
            return df
        except Exception as e:
            print(f"Erro ao carregar eventos {caminho}: {e}")
    return pd.DataFrame(columns=colunas)

def _salvar_eventos_pilar(df: pd.DataFrame, caminho: str):
    try:
        garantir_diretorio()
        df.to_csv(caminho, index=False, encoding="utf-8-sig")
    except Exception as e:
        print(f"Erro ao salvar eventos {caminho}: {e}")

def carregar_excesso_velocidade() -> pd.DataFrame:
    return _carregar_eventos_pilar(ARQUIVO_EXCESSO_VELOCIDADE)

def salvar_excesso_velocidade(df: pd.DataFrame):
    _salvar_eventos_pilar(df, ARQUIVO_EXCESSO_VELOCIDADE)

def carregar_controle_jornada() -> pd.DataFrame:
    return _carregar_eventos_pilar(ARQUIVO_CONTROLE_JORNADA)

def salvar_controle_jornada(df: pd.DataFrame):
    _salvar_eventos_pilar(df, ARQUIVO_CONTROLE_JORNADA)

def normalizar_categoria_evento(categoria: str) -> str:
    c = DataUtils.normalizar_texto(categoria)
    if "BITRUCK" in c: return "BITRUCK"
    if "CARRETA" in c: return "CARRETA"
    if "BITREM" in c: return "BITREM"
    if "RODOTREM" in c: return "RODOTREM"
    if "RODOENTREGA" in c or "RODO ENTREGA" in c: return "RODOENTREGA"
    if "TRUCK" in c: return "TRUCK"
    return c

def valor_ponto_categoria(categoria: str) -> float:
    return float(VALOR_PONTO_POR_CATEGORIA.get(normalizar_categoria_evento(categoria), 0.0))

def aplicar_descontos_eventos(df_resumo: pd.DataFrame, df_excesso: pd.DataFrame, df_jornada: pd.DataFrame) -> pd.DataFrame:
    res = df_resumo.copy()
    if res.empty:
        return res
    for c in ["EVENTOS_EXCESSO_VELOCIDADE","DESCONTO_EXCESSO_VELOCIDADE","EVENTOS_CONTROLE_JORNADA","DESCONTO_CONTROLE_JORNADA"]:
        res[c] = 0

    mapa_categoria = {}
    if "MOTORISTA" in res.columns and "CATEGORIA" in res.columns:
        for _, rr in res[["MOTORISTA", "CATEGORIA"]].dropna(subset=["MOTORISTA"]).iterrows():
            chave = DataUtils.normalizar_texto(rr["MOTORISTA"])
            categoria = DataUtils.normalizar_texto(rr["CATEGORIA"])
            if chave and categoria:
                mapa_categoria[chave] = categoria

    for df, prefix in [(df_excesso, "EXCESSO"), (df_jornada, "JORNADA")]:
        if df is None or df.empty:
            continue
        tmp = df.copy()
        tmp["MOTORISTA_N"] = tmp["MOTORISTA"].apply(DataUtils.normalizar_texto)
        tmp["EVENTOS"] = pd.to_numeric(tmp["EVENTOS"], errors="coerce").fillna(0)
        if "CATEGORIA" not in tmp.columns:
            tmp["CATEGORIA"] = ""
        tmp["CATEGORIA"] = tmp.apply(
            lambda r: r["CATEGORIA"] if DataUtils.normalizar_texto(r.get("CATEGORIA", ""))
            else mapa_categoria.get(r["MOTORISTA_N"], ""),
            axis=1,
        )
        tmp["VALOR_PONTO"] = tmp["CATEGORIA"].apply(valor_ponto_categoria)
        tmp["DESCONTO"] = tmp["EVENTOS"] * tmp["VALOR_PONTO"]
        agg = tmp.groupby("MOTORISTA_N").agg(EVENTOS=("EVENTOS","sum"), DESCONTO=("DESCONTO","sum"))
        for idx in res.index:
            m = DataUtils.normalizar_texto(res.at[idx,"MOTORISTA"])
            if m in agg.index:
                if prefix=="EXCESSO":
                    res.at[idx,"EVENTOS_EXCESSO_VELOCIDADE"] = int(agg.at[m,"EVENTOS"])
                    res.at[idx,"DESCONTO_EXCESSO_VELOCIDADE"] = float(agg.at[m,"DESCONTO"])
                else:
                    res.at[idx,"EVENTOS_CONTROLE_JORNADA"] = int(agg.at[m,"EVENTOS"])
                    res.at[idx,"DESCONTO_CONTROLE_JORNADA"] = float(agg.at[m,"DESCONTO"])
    for idx in res.index:
        ex_n = int(res.at[idx,"EVENTOS_EXCESSO_VELOCIDADE"])
        ex_d = float(res.at[idx,"DESCONTO_EXCESSO_VELOCIDADE"])
        jo_d = float(res.at[idx,"DESCONTO_CONTROLE_JORNADA"])
        premio = float(res.at[idx,"PREMIO"])
        if ex_n > 30:
            res.at[idx,"PREMIO"] = 0.0
            res.at[idx,"STATUS_PREMIO"] = "DESCLASSIFICADO"
            res.at[idx,"MOTIVO_DESCLASSIFICACAO"] = f"Excesso de velocidade: {ex_n} eventos (>30) - prêmio perdido integralmente"
        elif int(res.at[idx,"EVENTOS_CONTROLE_JORNADA"]) >= 130:
            res.at[idx,"PREMIO"] = 0.0
            res.at[idx,"STATUS_PREMIO"] = "DESCLASSIFICADO"
            res.at[idx,"MOTIVO_DESCLASSIFICACAO"] = f"Controle de jornada: {int(res.at[idx,'EVENTOS_CONTROLE_JORNADA'])} eventos (>=130) - prêmio perdido integralmente"
        else:
            novo = max(0.0, premio - ex_d - jo_d)
            res.at[idx,"PREMIO"] = novo
            partes=[]
            if ex_n: partes.append(f"Excesso de velocidade: -R$ {ex_d:.2f} ({ex_n} eventos)")
            jo_n=int(res.at[idx,"EVENTOS_CONTROLE_JORNADA"])
            if jo_n: partes.append(f"Controle de jornada: -R$ {jo_d:.2f} ({jo_n} eventos)")
            if partes:
                base=str(res.at[idx,"MOTIVO_DESCLASSIFICACAO"])
                if base=="Elegível / Em conformidade": base=""
                res.at[idx,"MOTIVO_DESCLASSIFICACAO"]=" | ".join([x for x in [base,*partes] if x])
    return res

def recalcular_saida_dashboard(current_tuple, df_excesso, df_jornada):
    vals=list(current_tuple)
    res_f=aplicar_descontos_eventos(vals[-1], df_excesso, df_jornada)
    premio=float(res_f["PREMIO"].sum()) if not res_f.empty else 0.0
    vals[0]=f"R$ {premio:,.2f}".replace(",","X").replace(".",",").replace("X",".")
    rv=res_f.copy()
    if not rv.empty and "PREMIO" in rv.columns:
        rv["PREMIO"]=rv["PREMIO"].map(lambda x:f"R$ {float(x):,.2f}".replace(",","X").replace(".",",").replace("X","."))
    vals[6]=rv
    vals[7]=gerar_tabela_rh(res_f)
    vals[-1]=res_f
    return tuple(vals)

def carregar_inativos() -> dict:
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

def carregar_datas_motoristas() -> dict:
  if os.path.exists(ARQUIVO_DATAS_MOTORISTAS):
    try:
      df = pd.read_csv(ARQUIVO_DATAS_MOTORISTAS, dtype=str, encoding="utf-8-sig")
      if "MOTORISTA" in df.columns and "DATA_CONTRATACAO" in df.columns:
        return {
            DataUtils.normalizar_texto(m): str(d or "").strip()
            for m, d in zip(df["MOTORISTA"], df["DATA_CONTRATACAO"])
            if str(m).strip()
        }
    except Exception as e:
      print(f"Erro ao carregar datas de motoristas: {e}")
  return {}

def salvar_data_contratacao_motorista(motorista: str, data_contratacao: str):
  garantir_diretorio()
  motorista = DataUtils.normalizar_texto(motorista)
  data_contratacao = str(data_contratacao or "").strip()
  df = pd.DataFrame(columns=["MOTORISTA", "DATA_CONTRATACAO"])
  if os.path.exists(ARQUIVO_DATAS_MOTORISTAS):
    try:
      df = pd.read_csv(ARQUIVO_DATAS_MOTORISTAS, dtype=str, encoding="utf-8-sig")
      for c in ["MOTORISTA", "DATA_CONTRATACAO"]:
        if c not in df.columns:
          df[c] = ""
      df = df[["MOTORISTA", "DATA_CONTRATACAO"]].copy()
    except Exception:
      df = pd.DataFrame(columns=["MOTORISTA", "DATA_CONTRATACAO"])

  df["MOTORISTA"] = df["MOTORISTA"].apply(DataUtils.normalizar_texto)
  mask = df["MOTORISTA"] == motorista
  if mask.any():
    df.loc[mask, "DATA_CONTRATACAO"] = data_contratacao
  else:
    df = pd.concat([df, pd.DataFrame([{
        "MOTORISTA": motorista,
        "DATA_CONTRATACAO": data_contratacao,
    }])], ignore_index=True)
  df.to_csv(ARQUIVO_DATAS_MOTORISTAS, index=False, encoding="utf-8-sig")

def carregar_codigos_funcionais() -> dict:
  if os.path.exists(ARQUIVO_CODIGOS_FUNCIONAIS):
    try:
      df = pd.read_csv(ARQUIVO_CODIGOS_FUNCIONAIS, dtype=str, encoding="utf-8-sig")
      if "MOTORISTA" in df.columns and "CODIGO_FUNCIONAL" in df.columns:
        return {
            DataUtils.normalizar_texto(m): str(c or "").strip()
            for m, c in zip(df["MOTORISTA"], df["CODIGO_FUNCIONAL"])
            if str(m).strip()
        }
    except Exception as e:
      print(f"Erro ao carregar códigos funcionais: {e}")
  return {}

def salvar_codigo_funcional_motorista(motorista: str, codigo: str):
  garantir_diretorio()
  motorista = DataUtils.normalizar_texto(motorista)
  codigo = str(codigo or "").strip()
  df = pd.DataFrame(columns=["MOTORISTA", "CODIGO_FUNCIONAL"])
  if os.path.exists(ARQUIVO_CODIGOS_FUNCIONAIS):
    try:
      df = pd.read_csv(ARQUIVO_CODIGOS_FUNCIONAIS, dtype=str, encoding="utf-8-sig")
      for c in ["MOTORISTA", "CODIGO_FUNCIONAL"]:
        if c not in df.columns:
          df[c] = ""
      df = df[["MOTORISTA", "CODIGO_FUNCIONAL"]].copy()
    except Exception:
      df = pd.DataFrame(columns=["MOTORISTA", "CODIGO_FUNCIONAL"])

  df["MOTORISTA"] = df["MOTORISTA"].apply(DataUtils.normalizar_texto)
  mask = df["MOTORISTA"] == motorista
  if mask.any():
    df.loc[mask, "CODIGO_FUNCIONAL"] = codigo
  else:
    df = pd.concat([df, pd.DataFrame([{
        "MOTORISTA": motorista,
        "CODIGO_FUNCIONAL": codigo,
    }])], ignore_index=True)
  df.to_csv(ARQUIVO_CODIGOS_FUNCIONAIS, index=False, encoding="utf-8-sig")

def atualizar_data_inativacao_motorista(motorista: str, data_inativacao: str):
  motorista = DataUtils.normalizar_texto(motorista)
  garantir_diretorio()
  df = pd.DataFrame(columns=["TIPO", "VALOR", "DATA_INATIVACAO"])
  if os.path.exists(ARQUIVO_INATIVOS):
    try:
      df = pd.read_csv(ARQUIVO_INATIVOS, dtype=str, encoding="utf-8-sig")
    except Exception:
      pass
  for c in ["TIPO", "VALOR", "DATA_INATIVACAO"]:
    if c not in df.columns:
      df[c] = ""
  mask = (df["TIPO"] == "MOTORISTA") & (df["VALOR"] == motorista)
  if mask.any():
    df.loc[mask, "DATA_INATIVACAO"] = str(data_inativacao or "").strip()
    df.to_csv(ARQUIVO_INATIVOS, index=False, encoding="utf-8-sig")

def alternar_inativo(tipo: str, valor: str, inativar: bool = True, data_inativacao: str = ""):
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
      data_atual = str(data_inativacao or "").strip() or datetime.now().strftime("%d/%m/%Y")
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

    if (not os.path.isfile(self.CAMINHO_ABASTECIMENTOS)
        and not _listar_arquivos_abastecimentos_mensais()
        and not any(name.startswith("historico_abastecimentos_") and name.endswith(".csv")
                    for name in os.listdir(DATA_DIR if DATA_DIR and DATA_DIR != "." else "."))):
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
    if os.path.isfile(self.CAMINHO_ABASTECIMENTOS):
      self.CAMINHO_ABASTECIMENTOS = self._resolver_caminho_real(self.CAMINHO_ABASTECIMENTOS)


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


def _parse_datetime_flex(valor) -> pd.Timestamp:
  if valor is None:
    return pd.NaT
  try:
    if pd.isna(valor):
      return pd.NaT
  except Exception:
    pass

  if isinstance(valor, (pd.Timestamp, datetime, np.datetime64)):
    try:
      return pd.Timestamp(valor)
    except Exception:
      return pd.NaT

  if isinstance(valor, (int, float, np.integer, np.floating)):
    try:
      n = float(valor)
      if n > 20000:
        return pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
      if 0 <= n < 1:
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(n, unit="D")
    except Exception:
      return pd.NaT

  texto = str(valor).strip()
  if not texto:
    return pd.NaT
  try:
    dt = pd.to_datetime(texto, dayfirst=True, errors="coerce")
    if pd.notna(dt):
      return pd.Timestamp(dt)
  except Exception:
    pass
  return pd.NaT


def criar_data_filtro(valor) -> pd.Timestamp:
  parsed = parse_data_filtro(valor)
  return parsed if parsed is not None else pd.NaT


def calcular_dias_ausencia(data_inicio, data_fim):
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
    df_custom = carregar_motoristas_customizados().copy()

    if df_custom.empty:
      return pd.DataFrame(
          columns=[
              "MOTORISTA_CADASTRO",
              "TIPO_CADASTRO",
              "BASE_CADASTRO",
              "EH_FOLGUISTA",
              "CODIGO_FUNCIONAL",
              "DATA_CONTRATACAO",
              "STATUS",
              "DATA_INATIVACAO",
          ]
      )

    for c in ["MOTORISTAS", "TIPO", "BASE"]:
      if c not in df_custom.columns:
        df_custom[c] = ""

    cadastro = pd.DataFrame({
        "MOTORISTA_CADASTRO": df_custom["MOTORISTAS"].apply(DataUtils.normalizar_texto),
        "TIPO_CADASTRO": df_custom["TIPO"].apply(DataUtils.normalizar_texto).replace({"TOCO": "TRUCK"}),
        "BASE_CADASTRO": df_custom["BASE"].apply(DataUtils.normalizar_texto),
    })

    cadastro = cadastro[
        (cadastro["MOTORISTA_CADASTRO"] != "")
        & (cadastro["TIPO_CADASTRO"] != "")
    ].copy()

    cadastro["EH_FOLGUISTA"] = cadastro["TIPO_CADASTRO"].eq("FOLGUISTA")
    cadastro = cadastro.drop_duplicates("MOTORISTA_CADASTRO", keep="last").reset_index(drop=True)

    inativos_dict = carregar_inativos().get("MOTORISTA", {})
    datas_contratacao = carregar_datas_motoristas()
    codigos_funcionais = carregar_codigos_funcionais()

    cadastro["CODIGO_FUNCIONAL"] = cadastro["MOTORISTA_CADASTRO"].apply(
        lambda x: str(codigos_funcionais.get(x, "") or "").strip()
    )
    cadastro["DATA_CONTRATACAO"] = cadastro["MOTORISTA_CADASTRO"].apply(
        lambda x: str(datas_contratacao.get(x, "") or "").strip()
    )
    cadastro["STATUS"] = cadastro["MOTORISTA_CADASTRO"].apply(
        lambda x: "INATIVO" if x in inativos_dict else "ATIVO"
    )
    cadastro["DATA_INATIVACAO"] = cadastro["MOTORISTA_CADASTRO"].apply(
        lambda x: str(inativos_dict.get(x, "") or "").strip()
    )

    return cadastro

  def carregar_viagens(self) -> pd.DataFrame:
    arquivos = _listar_arquivos_viagens()
    if not arquivos:
      return pd.DataFrame()

    registros = []
    for caminho in arquivos:
      try:
        bruto = pd.read_excel(caminho, sheet_name=0, dtype=object, keep_default_na=False)
        if bruto.empty:
          continue

        col_dt_ini = DataUtils.encontrar_coluna(bruto, ["Dt Macro Inicial", "DATA INICIAL", "INICIO VIAGEM"])
        col_dt_fim = DataUtils.encontrar_coluna(bruto, ["Dt Macro Final", "DATA FINAL", "FIM VIAGEM"])
        col_odm_ini = DataUtils.encontrar_coluna(bruto, ["Odm Inicial", "ODM INICIAL", "KM INICIAL"])
        col_odm_fim = DataUtils.encontrar_coluna(bruto, ["Odm Final", "ODM FINAL", "KM FINAL"])
        col_km_total = DataUtils.encontrar_coluna(bruto, ["Km Total", "KM TOTAL"])
        col_motorista = DataUtils.encontrar_coluna(bruto, ["Motorista", "CONDUTOR", "MOTORISTA"])
        col_placa = DataUtils.encontrar_coluna(bruto, ["Dim Veiculo Enterprise - Sk Veiculo → Placa", "PLACA", "CAVALO"])
        col_origem = DataUtils.encontrar_coluna(bruto, ["Cidade Inicial", "ORIGEM", "CIDADE ORIGEM"])
        col_destino = DataUtils.encontrar_coluna(bruto, ["Cidade Final", "DESTINO", "CIDADE DESTINO"])

        obrigatorias = [col_dt_ini, col_dt_fim, col_odm_ini, col_odm_fim, col_motorista, col_placa]
        if any(c is None for c in obrigatorias):
          print(f"Arquivo de viagens ignorado por colunas ausentes: {os.path.basename(caminho)}")
          continue

        out = pd.DataFrame(index=bruto.index)
        out["VIAGEM_DATA_INICIO"] = bruto[col_dt_ini].apply(_parse_datetime_flex)
        out["VIAGEM_DATA_FIM"] = bruto[col_dt_fim].apply(_parse_datetime_flex)
        out["VIAGEM_ODM_INICIAL"] = bruto[col_odm_ini].apply(DataUtils.converter_numero)
        out["VIAGEM_ODM_FINAL"] = bruto[col_odm_fim].apply(DataUtils.converter_numero)
        out["VIAGEM_KM_TOTAL"] = bruto[col_km_total].apply(DataUtils.converter_numero) if col_km_total else np.nan
        out["MOTORISTA_VIAGEM"] = bruto[col_motorista].apply(DataUtils.normalizar_texto)
        out["PLACA_VIAGEM"] = bruto[col_placa].apply(DataUtils.padronizar_placa)
        out["VIAGEM_ORIGEM"] = bruto[col_origem].astype(str).str.strip() if col_origem else ""
        out["VIAGEM_DESTINO"] = bruto[col_destino].astype(str).str.strip() if col_destino else ""
        out["VIAGEM_ARQUIVO"] = os.path.basename(caminho)

        out = out[
            (out["PLACA_VIAGEM"] != "")
            & out["VIAGEM_ODM_INICIAL"].notna()
            & out["VIAGEM_ODM_FINAL"].notna()
            & (out["MOTORISTA_VIAGEM"] != "")
        ].copy()
        if not out.empty:
          out["_ODM_MIN"] = out[["VIAGEM_ODM_INICIAL", "VIAGEM_ODM_FINAL"]].min(axis=1)
          out["_ODM_MAX"] = out[["VIAGEM_ODM_INICIAL", "VIAGEM_ODM_FINAL"]].max(axis=1)
          registros.append(out)
      except Exception as exc:
        print(f"Erro ao carregar viagens {os.path.basename(caminho)}: {exc}")

    if not registros:
      return pd.DataFrame()
    return pd.concat(registros, ignore_index=True)

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
    col_hora = DataUtils.encontrar_coluna(df, ["HORA", "Hora", "HORARIO", "HORA ABASTECIMENTO"])

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
      resultado["DATA_HORA_ABASTECIMENTO"] = resultado[col_data].apply(_parse_datetime_flex)
      if col_hora:
        hora_series = resultado[col_hora].apply(_parse_datetime_flex)
        mascara_hora = hora_series.notna() & resultado["DATA_HORA_ABASTECIMENTO"].notna()
        resultado.loc[mascara_hora, "DATA_HORA_ABASTECIMENTO"] = (
            resultado.loc[mascara_hora, "DATA_HORA_ABASTECIMENTO"].dt.normalize()
            + pd.to_timedelta(
                hora_series.loc[mascara_hora].dt.hour * 3600
                + hora_series.loc[mascara_hora].dt.minute * 60
                + hora_series.loc[mascara_hora].dt.second,
                unit="s",
            )
        )
    else:
      resultado["DATA_ORIGINAL"] = pd.NaT
      resultado["DATA_NUM"] = pd.NaT
      resultado["DATA_FILTRO"] = pd.NaT
      resultado["DATA_HORA_ABASTECIMENTO"] = pd.NaT

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
# VALIDAÇÃO DO MOTORISTA PELO HISTÓRICO DE VIAGENS
# ================================================================
def associar_motorista_viagem(abastecimentos: pd.DataFrame, viagens: pd.DataFrame) -> pd.DataFrame:
  """Concilia cada abastecimento com o ciclo de viagens entre dois abastecimentos."""
  base = abastecimentos.copy()
  base["MOTORISTA_ABASTECIMENTO_ORIGINAL"] = base.get("CONDUTOR_NORMALIZADO", "").astype(str)
  base["MOTORISTA_VIAGEM"] = ""
  base["MOTORISTA_CONSIDERADO"] = base["MOTORISTA_ABASTECIMENTO_ORIGINAL"]
  base["STATUS_VALIDACAO_VIAGEM"] = "SEM ARQUIVO DE VIAGENS" if viagens is None or viagens.empty else "SEM ABASTECIMENTO ANTERIOR"
  for c in ["VIAGEM_ORIGEM", "VIAGEM_DESTINO", "VIAGEM_ARQUIVO", "VIAGENS_CICLO_MOTORISTAS"]:
    base[c] = ""
  for c in ["VIAGEM_ODM_INICIAL", "VIAGEM_ODM_FINAL", "VIAGEM_KM_TOTAL", "KM_ABASTECIMENTO_ANTERIOR"]:
    base[c] = np.nan
  for c in ["VIAGEM_DATA_INICIO", "VIAGEM_DATA_FIM", "DATA_HORA_ABASTECIMENTO_ANTERIOR"]:
    base[c] = pd.Series(pd.NaT, index=base.index, dtype="datetime64[ns]")
  base["VIAGENS_CICLO_QTD"] = 0

  if base.empty or viagens is None or viagens.empty:
    return base

  viagens = viagens.copy()
  viagens = viagens[
      viagens["PLACA_VIAGEM"].notna()
      & (viagens["PLACA_VIAGEM"].astype(str).str.strip() != "")
      & viagens["_ODM_MIN"].notna()
      & viagens["_ODM_MAX"].notna()
      & viagens["MOTORISTA_VIAGEM"].notna()
      & (viagens["MOTORISTA_VIAGEM"].astype(str).str.strip() != "")
  ].copy()
  if viagens.empty:
    return base

  viagens["VIAGEM_DATA_INICIO"] = viagens["VIAGEM_DATA_INICIO"].apply(_parse_datetime_flex)
  viagens["VIAGEM_DATA_FIM"] = viagens["VIAGEM_DATA_FIM"].apply(_parse_datetime_flex)
  por_placa = {placa: grp for placa, grp in viagens.groupby("PLACA_VIAGEM", sort=False)}

  base["_TMP_DT"] = pd.to_datetime(base.get("DATA_HORA_ABASTECIMENTO", base.get("DATA_NUM")), errors="coerce")
  base["_TMP_DT"] = base["_TMP_DT"].apply(lambda x: _parse_datetime_flex(x) if pd.notna(x) else pd.NaT)
  base = base.sort_values(["PLACA_PADRONIZADA", "KM_ATUAL_NUM", "_TMP_DT"], kind="stable")

  for placa, idxs in base.groupby("PLACA_PADRONIZADA", sort=False).groups.items():
    idx_list = list(idxs)
    idx_list = sorted(idx_list, key=lambda i: (
        pd.Timestamp(base.at[i, "_TMP_DT"]) if pd.notna(base.at[i, "_TMP_DT"]) else pd.Timestamp("1900-01-01"),
        float(base.at[i, "KM_ATUAL_NUM"]) if pd.notna(base.at[i, "KM_ATUAL_NUM"]) else float("inf"),
    ))
    cand_viagens = por_placa.get(placa)
    if cand_viagens is None or cand_viagens.empty:
      for i in idx_list:
        base.at[i, "STATUS_VALIDACAO_VIAGEM"] = "NÃO LOCALIZADO"
      continue

    for pos, idx in enumerate(idx_list):
      km_atual = pd.to_numeric(base.at[idx, "KM_ATUAL_NUM"], errors="coerce")
      if pd.isna(km_atual) or km_atual <= 0:
        continue
      if pos == 0:
        base.at[idx, "STATUS_VALIDACAO_VIAGEM"] = "SEM ABASTECIMENTO ANTERIOR"
        continue

      idx_prev = idx_list[pos - 1]
      km_prev = pd.to_numeric(base.at[idx_prev, "KM_ATUAL_NUM"], errors="coerce")
      if pd.isna(km_prev) or km_prev >= km_atual:
        base.at[idx, "STATUS_VALIDACAO_VIAGEM"] = "CICLO KM INVÁLIDO"
        continue

      dt_atual = base.at[idx, "_TMP_DT"]
      dt_prev = base.at[idx_prev, "_TMP_DT"]
      if pd.isna(dt_atual):
        dt_atual = base.at[idx, "DATA_NUM"] if "DATA_NUM" in base.columns else pd.NaT
      if pd.isna(dt_prev):
        dt_prev = base.at[idx_prev, "DATA_NUM"] if "DATA_NUM" in base.columns else pd.NaT
      dt_atual = _parse_datetime_flex(dt_atual)
      dt_prev = _parse_datetime_flex(dt_prev)

      base.at[idx, "KM_ABASTECIMENTO_ANTERIOR"] = float(km_prev)
      if pd.notna(dt_prev):
        base.at[idx, "DATA_HORA_ABASTECIMENTO_ANTERIOR"] = dt_prev

      ciclo_ini = float(km_prev)
      ciclo_fim = float(km_atual)
      cand = cand_viagens.copy()
      cand = cand[(cand["_ODM_MAX"] > ciclo_ini) & (cand["_ODM_MIN"] <= ciclo_fim)].copy()

      if pd.notna(dt_prev):
        cand = cand[cand["VIAGEM_DATA_FIM"].isna() | (cand["VIAGEM_DATA_FIM"] >= pd.Timestamp(dt_prev) - pd.Timedelta(days=1))]
      if pd.notna(dt_atual):
        limite_fim = pd.Timestamp(dt_atual) + pd.Timedelta(days=2)
        cand = cand[cand["VIAGEM_DATA_INICIO"].isna() | (cand["VIAGEM_DATA_INICIO"] <= limite_fim)]

      if cand.empty:
        base.at[idx, "STATUS_VALIDACAO_VIAGEM"] = "NÃO LOCALIZADO NO CICLO"
        continue

      cand["_MOTORISTA_N"] = cand["MOTORISTA_VIAGEM"].apply(DataUtils.normalizar_texto)
      motoristas = sorted([m for m in cand["_MOTORISTA_N"].dropna().unique().tolist() if m])
      base.at[idx, "VIAGENS_CICLO_QTD"] = int(len(cand))
      base.at[idx, "VIAGENS_CICLO_MOTORISTAS"] = " | ".join(motoristas)

      coberturas = {}
      for _, v in cand.iterrows():
        inicio = max(float(v["_ODM_MIN"]), ciclo_ini)
        fim = min(float(v["_ODM_MAX"]), ciclo_fim)
        cobertura = max(0.0, fim - inicio)
        mot = v["_MOTORISTA_N"]
        coberturas[mot] = coberturas.get(mot, 0.0) + cobertura

      if len(motoristas) == 1:
        motorista_viagem = motoristas[0]
        escolhido = cand[cand["_MOTORISTA_N"] == motorista_viagem].sort_values(["VIAGEM_DATA_FIM", "_ODM_MAX"], kind="stable").iloc[-1]
        base.at[idx, "MOTORISTA_VIAGEM"] = motorista_viagem
        base.at[idx, "MOTORISTA_CONSIDERADO"] = motorista_viagem
        original = DataUtils.normalizar_texto(base.at[idx, "MOTORISTA_ABASTECIMENTO_ORIGINAL"])
        base.at[idx, "STATUS_VALIDACAO_VIAGEM"] = "VALIDADO PELO CICLO" if original == motorista_viagem else "CORRIGIDO PELO CICLO"
        base.at[idx, "VIAGEM_ORIGEM"] = str(escolhido.get("VIAGEM_ORIGEM", ""))
        base.at[idx, "VIAGEM_DESTINO"] = str(escolhido.get("VIAGEM_DESTINO", ""))
        base.at[idx, "VIAGEM_ARQUIVO"] = str(escolhido.get("VIAGEM_ARQUIVO", ""))
        base.at[idx, "VIAGEM_ODM_INICIAL"] = escolhido.get("VIAGEM_ODM_INICIAL", np.nan)
        base.at[idx, "VIAGEM_ODM_FINAL"] = escolhido.get("VIAGEM_ODM_FINAL", np.nan)
        base.at[idx, "VIAGEM_KM_TOTAL"] = escolhido.get("VIAGEM_KM_TOTAL", np.nan)
        base.at[idx, "VIAGEM_DATA_INICIO"] = escolhido.get("VIAGEM_DATA_INICIO", pd.NaT)
        base.at[idx, "VIAGEM_DATA_FIM"] = escolhido.get("VIAGEM_DATA_FIM", pd.NaT)
      else:
        base.at[idx, "STATUS_VALIDACAO_VIAGEM"] = "PENDENTE DE VALIDAÇÃO — MÚLTIPLOS MOTORISTAS"
        mot_principal = max(coberturas, key=coberturas.get) if coberturas else ""
        if mot_principal:
          cand_principal = cand[cand["_MOTORISTA_N"] == mot_principal]
          if not cand_principal.empty:
            escolhido = cand_principal.sort_values(["VIAGEM_DATA_FIM", "_ODM_MAX"], kind="stable").iloc[-1]
            base.at[idx, "VIAGEM_ORIGEM"] = str(escolhido.get("VIAGEM_ORIGEM", ""))
            base.at[idx, "VIAGEM_DESTINO"] = str(escolhido.get("VIAGEM_DESTINO", ""))
            base.at[idx, "VIAGEM_ARQUIVO"] = str(escolhido.get("VIAGEM_ARQUIVO", ""))
            base.at[idx, "VIAGEM_ODM_INICIAL"] = escolhido.get("VIAGEM_ODM_INICIAL", np.nan)
            base.at[idx, "VIAGEM_ODM_FINAL"] = escolhido.get("VIAGEM_ODM_FINAL", np.nan)
            base.at[idx, "VIAGEM_KM_TOTAL"] = escolhido.get("VIAGEM_KM_TOTAL", np.nan)
            base.at[idx, "VIAGEM_DATA_INICIO"] = escolhido.get("VIAGEM_DATA_INICIO", pd.NaT)
            base.at[idx, "VIAGEM_DATA_FIM"] = escolhido.get("VIAGEM_DATA_FIM", pd.NaT)

  base = base.drop(columns=["_TMP_DT"], errors="ignore")
  return base
