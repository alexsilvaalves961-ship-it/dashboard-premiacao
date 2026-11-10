import os
import base64
import hmac
import hashlib
import re
import io
import zipfile
from pathlib import Path
import unicodedata
from dataclasses import dataclass
from datetime import datetime, date
from typing import Optional, Tuple
import textwrap
import time

import streamlit as st
import numpy as np
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt

# ================================================================
# AUTOMAÇÃO DE EXTRAÇÃO DE ABASTECIMENTOS (PORTAL CIAPETRO)
# ================================================================
def executar_automacao_abastecimentos_ciapetro(mes_ano_str="08_2026"):
  """Acessa o portal da frota da Ciapetro, realiza o login e baixa

  os dados de abastecimento salvando no formato esperado pelo app
  (ex: abastecimentos_MMYYYY.xlsx).
  """
  url_login = "https://portalfrota.uauh.com.br/app/_abastecimento/"
  cliente = "ciapetro"
  usuario = "administrativo.transporte@ciapetro.com.br"
  senha = "123456"

  options = webdriver.ChromeOptions()
  # Descomente a linha abaixo caso queira rodar em segundo plano (sem abrir o navegador visualmente)
  # options.add_argument("--headless")

  driver = webdriver.Chrome(options=options)
  driver.maximize_window()

  try:
    print("1. Acessando o portal da frota...")
    driver.get(url_login)
    time.sleep(3)

    wait = WebDriverWait(driver, 15)

    # Preenche o login (ajuste os atributos name/id se necessário após inspecionar o site)
    try:
      campo_cliente = driver.find_element(By.NAME, "client")
      campo_cliente.send_keys(cliente)
    except Exception:
      pass

    campo_usuario = wait.until(
        EC.presence_of_element_located((By.NAME, "username"))
    )
    campo_usuario.send_keys(usuario)

    campo_senha = driver.find_element(By.NAME, "password")
    campo_senha.send_keys(senha)
    campo_senha.send_keys(Keys.RETURN)

    print("2. Login efetuado com sucesso! Aguardando painel...")
    time.sleep(6)

    # Navega para a página de abastecimento
    driver.get(url_login)
    time.sleep(5)

    print("3. Coletando os dados da tabela de abastecimento...")
    tabelas = pd.read_html(driver.page_source)

    if tabelas:
      df_abastecimentos = tabelas[0]

      # Salva no padrão que o aplicativo já reconhece automaticamente (ex: abastecimentos_082026.xlsx)
      nome_arquivo = f"abastecimentos_{mes_ano_str}.xlsx"
      caminho_salvo = os.path.join(
          DATA_DIR if DATA_DIR and DATA_DIR != "." else ".", nome_arquivo
      )

      df_abastecimentos.to_excel(caminho_salvo, index=False)
      print(f"Sucesso! Dados salvos em '{caminho_salvo}'.")
      return True
    else:
      print("Aviso: Nenhuma tabela estruturada encontrada na página.")
      return False

  except Exception as exc:
    print(f"Erro na automação Selenium: {exc}")
    return False

  finally:
    driver.quit()


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

# Abastecimentos mensais: arquivos do tipo abastecimentos_MMYYYY.xlsx
# O sistema arquiva internamente cada competência em um CSV separado, evitando
# manter uma única planilha gigante. O arquivo legado uah_abastecimentos_3.xlsx
# é usado somente para uma migração inicial, quando ainda não existirem os arquivos
# históricos internos.
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
  # arquivo 082026 = competência 26/07/2026 a 25/08/2026
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
  # Reutiliza exatamente a mesma normalização do DataLoader, mas recebe o DataFrame
  # já lido para permitir arquivamento mensal sem alterar as colunas originais.
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
  # Identificador estável do registro, necessário para ordenar os abastecimentos
  # quando os dados vêm dos históricos mensais.
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
  """Importa cada arquivo mensal para um histórico por competência.

  Regras:
  - O arquivo mensal da competência é a fonte oficial daquela competência.
  - Ao existir o Excel mensal, o histórico daquela competência é RECONSTRUÍDO
    (não concatenado), evitando duplicações e resíduos de versões anteriores.
  - O histórico de competências anteriores permanece salvo mesmo que o Excel
    correspondente seja removido da hospedagem.
  - Cada linha recebe ID_ABASTECIMENTO único por competência/arquivo.
  """
  garantir_diretorio()
  cols = _colunas_historico_abastecimentos()

  # Primeiro, processa os arquivos mensais presentes.
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
      # Se o arquivo mensal estiver vazio/incompatível, não destrói um histórico já válido.
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
    # Fonte mensal substitui completamente o snapshot anterior da mesma competência.
    _salvar_csv_com_backup(part, arq)
    mensal_processados.add((mm, yyyy))

  # Migração do legado somente para competências ainda sem histórico mensal.
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
          # Não reprocessa o legado sobre um histórico já consolidado.
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
          _salvar_csv_com_backup(part[cols], arq)

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
      # Histórico antigo sem ID: gera identificador único usando competência + posição.
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
  """Garante a existência do diretório de dados antes de operações de escrita."""
  if DATA_DIR and DATA_DIR != "." and not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR, exist_ok=True)


# ================================================================
# BACKUP / PROTEÇÃO DOS DADOS INSERIDOS NO APLICATIVO
# ================================================================
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
BACKUP_RETENTION = 20
_BACKUP_AUTO_LAST_TS = 0.0
_BACKUP_AUTO_MIN_INTERVAL = 60.0  # um snapshot antes de uma sequência de gravações


def _arquivos_de_dados_para_backup():
    """Retorna somente os arquivos de dados persistentes do aplicativo.

    Entradas mensais Excel e código-fonte não entram no backup de dados do aplicativo.
    Históricos CSV, cadastros e demais arquivos de estado entram.
    """
    if not DATA_DIR or not os.path.isdir(DATA_DIR):
        return []
    extensoes = {".csv", ".parquet", ".json", ".db", ".sqlite", ".sqlite3", ".txt"}
    arquivos = []
    for root, dirs, files in os.walk(DATA_DIR):
        # O diretório de backups não pode entrar dentro de outro backup.
        dirs[:] = [d for d in dirs if os.path.abspath(os.path.join(root, d)) != os.path.abspath(BACKUP_DIR)]
        for nome in files:
            caminho = os.path.join(root, nome)
            if Path(nome).suffix.lower() in extensoes:
                arquivos.append(caminho)
    return sorted(arquivos)


def _nome_backup(prefixo="backup"):
    return f"{prefixo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"


def criar_backup_automatico(motivo="auto"):
    """Cria um snapshot local antes de uma gravação, sem interromper o salvamento."""
    global _BACKUP_AUTO_LAST_TS
    agora = datetime.now().timestamp()
    if agora - _BACKUP_AUTO_LAST_TS < _BACKUP_AUTO_MIN_INTERVAL:
        return None
    try:
        garantir_diretorio()
        os.makedirs(BACKUP_DIR, exist_ok=True)
        caminho_zip = os.path.join(BACKUP_DIR, _nome_backup(f"auto_{motivo}"))
        arquivos = _arquivos_de_dados_para_backup()
        with zipfile.ZipFile(caminho_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            manifest = [
                f"Data/hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                f"Motivo: {motivo}",
                f"DATA_DIR: {os.path.abspath(DATA_DIR or '.')}",
                "Arquivos:",
            ]
            base = os.path.abspath(DATA_DIR or ".")
            for caminho in arquivos:
                rel = os.path.relpath(caminho, base)
                zf.write(caminho, arcname=rel)
                manifest.append(rel)
            zf.writestr("MANIFESTO.txt", "\n".join(manifest))
        _BACKUP_AUTO_LAST_TS = agora
        _limpar_backups_antigos()
        return caminho_zip
    except Exception as exc:
        print(f"Aviso: não foi possível criar backup automático: {exc}")
        return None


def _limpar_backups_antigos():
    try:
        if not os.path.isdir(BACKUP_DIR):
            return
        arquivos = [
            os.path.join(BACKUP_DIR, n) for n in os.listdir(BACKUP_DIR)
            if n.lower().endswith(".zip")
        ]
        arquivos.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        for caminho in arquivos[BACKUP_RETENTION:]:
            try:
                os.remove(caminho)
            except OSError:
                pass
    except Exception:
        pass


def criar_backup_manual_bytes():
    """Monta um backup completo dos dados atuais em memória para download."""
    memoria = io.BytesIO()
    arquivos = _arquivos_de_dados_para_backup()
    base = os.path.abspath(DATA_DIR or ".")
    with zipfile.ZipFile(memoria, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        manifest = [
            f"Data/hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            f"DATA_DIR: {base}",
            "Arquivos:",
        ]
        for caminho in arquivos:
            rel = os.path.relpath(caminho, base)
            zf.write(caminho, arcname=rel)
            manifest.append(rel)
        zf.writestr("MANIFESTO.txt", "\n".join(manifest))
    memoria.seek(0)
    return memoria.getvalue()


def _listar_backups_salvos():
    try:
        if not os.path.isdir(BACKUP_DIR):
            return []
        dados = []
        for nome in os.listdir(BACKUP_DIR):
            if not nome.lower().endswith(".zip"):
                continue
            caminho = os.path.join(BACKUP_DIR, nome)
            try:
                dados.append({
                    "ARQUIVO": nome,
                    "DATA": datetime.fromtimestamp(os.path.getmtime(caminho)).strftime('%d/%m/%Y %H:%M:%S'),
                    "TAMANHO": f"{os.path.getsize(caminho)/1024:.1f} KB",
                    "CAMINHO": caminho,
                })
            except OSError:
                pass
        return sorted(dados, key=lambda x: x["DATA"], reverse=True)
    except Exception:
        return []


def _salvar_csv_com_backup(df: pd.DataFrame, caminho: str):
    """Grava um arquivo de dados criando snapshot antes da primeira gravação da sequência."""
    criar_backup_automatico("antes_salvar")
    garantir_diretorio()
    df.to_csv(caminho, index=False, encoding="utf-8-sig")


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
    _salvar_csv_com_backup(df, ARQUIVO_AUSENCIAS)
  except Exception as e:
    print(f"Erro ao salvar ausências: {e}")


def carregar_desclassificacoes() -> pd.DataFrame:
  """Carrega as desclassificações com DATA_EVENTO para respeitar a competência 26-25."""
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
        # Compatibilidade com lançamentos antigos: não atribuir uma data falsa.
        # Eles ficam marcados como legado e não entram em uma competência nova
        # até receberem uma data pelo administrador.
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
  """Salva as desclassificações no arquivo CSV com codificação UTF-8."""
  try:
    garantir_diretorio()
    _salvar_csv_com_backup(df, ARQUIVO_DESCLASSIFICACOES)
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
    _salvar_csv_com_backup(df, ARQUIVO_CATEGORIAS_CUSTOM)
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
  try: _salvar_csv_com_backup(df, ARQUIVO_CATEGORIAS_VIGENCIA)
  except Exception as e: print(f"Erro ao criar arquivo de vigência de categorias: {e}")
  return df[cols] if not df.empty else pd.DataFrame(columns=cols)

def salvar_categorias_vigencia(df: pd.DataFrame):
  try:
    garantir_diretorio(); cols=["MOTORISTA_CHAVE","CATEGORIA_ESCOLHIDA","DATA_INICIO","DATA_FIM"]; out=df.copy()
    for c in cols:
      if c not in out.columns: out[c] = ""
    _salvar_csv_com_backup(out[cols].fillna(""), ARQUIVO_CATEGORIAS_VIGENCIA)
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
    _salvar_csv_com_backup(df, ARQUIVO_FROTA_CUSTOM)
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
    _salvar_csv_com_backup(df, ARQUIVO_MOTORISTAS_CUSTOM)
  except Exception as e:
    print(f"Erro ao salvar motoristas customizados: {e}")




# ================================================================
# MIGRAÇÃO ÚNICA DO CADASTRO LEGADO -> GESTÃO DE CADASTROS
# ================================================================
# A partir desta versão, a Pasta4 não é fonte de dados em runtime.
# O snapshot abaixo é somente uma fotografia do cadastro legado usada UMA VEZ
# para povoar o cadastro persistente da Gestão de Cadastros.
# Depois da primeira gravação, o aplicativo passa a trabalhar exclusivamente
# com motoristas_customizados.csv.
LEGACY_CADASTRO_SNAPSHOT = [{'MOTORISTAS': 'ADEILSON DE OLIVEIRA ANGELINO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'AIRTON ANTONIO GONÇALVES', 'TIPO': 'BITRUCK', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'CLAUDINEI FRANCISCO FERREIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'CLAUDIO JOSE KREGENSKI', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'DANILO CASSIANO FERREIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'DIEISON APARECIDO DA CRUZ', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'DOUGLAS ENRIQUE DA SILVA LUIZ', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EDILSON LEITE DE CAMARGO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EDINEI MARCOS CORDEIRO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EDISON VIEIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EDSON RECOFKA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'EMERSON APARECIDO PEREIRA DA SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'FABIANO CASTILHO CALEGARI', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'GEDIVALDO SOUZA LUZ ALVES', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'GILMAR LOPACINSKI', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'INACIO DOUTOR', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JOAO PAULO LISNIOWSKI', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JONAS GOGOLA DE ANDRADE', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JOSIVAN DA SILVA OLIVEIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JOSUE LOPES DE SENE', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'JULIANA COQUES PAZ', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'LEOMAR MOREIRA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'LIDIOMAR DA SILVA DE SOUZA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'MARCELO DA SILVA E SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'MARCIO LEMOS MACHADO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'NELSON SOBOTHE', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'NILSON APARECIDO SAMPAIO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'NILSON RODRIGUES DE SOUZA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'NILTON DE JESUS RODRIGUES DE SOUZA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'ODAIR GONÇALVES MIRANDA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'PAULO DE MELO SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'PEDRO VANDERLEI BRASILINO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'RICARDO SERGIO DA SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'RODRIGO DE SOUZA MACHADO', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'VALDECI CARVALHO DA SILVA JUNIOR', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'VALDECI FERREIRA DA SILVA JUNIOR', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'WANDERLEY LOPES SILVA', 'TIPO': 'CARRETA', 'BASE': 'ARAUCARIA'}, {'MOTORISTAS': 'FELIPE TELES DA CRUZ', 'TIPO': 'BITRUCK', 'BASE': 'CAMPO GRANDE'}, {'MOTORISTAS': 'RENATO RIEFF MARIN', 'TIPO': 'CARRETA', 'BASE': 'CAMPO GRANDE'}, {'MOTORISTAS': 'ELICAR JUSTINO', 'TIPO': 'TRUCK', 'BASE': 'CHAPECO'}, {'MOTORISTAS': 'ANTONIO CARLOS BRAMBILA', 'TIPO': 'BITRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'JHONATAN ALVES DOS SANTOS', 'TIPO': 'BITRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'LINCOLN FRANCEL PIMENTA', 'TIPO': 'BITRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'RODRIGO LORENTINO', 'TIPO': 'BITRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'APARECIDO DIAMARAES', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'APARECIDO JOEL SANT ANA', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'APARECIDO RODRIGUES DA SILVA', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'CARLOS ELIER PIEROLI', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'ELISANGELA APARECIDA GOMES COELHO', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'FELIPE COMAR DIAS', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'MAURILIO FERREIRA DAS NEVES', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'RODOLFO MOZELLI SPAGOLLA', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'VALBER JUNIOR COSTA', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'WESLEI RIBEIRO JACOMINI', 'TIPO': 'CARRETA', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'GILBERTO BEZERRA PINTO', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'JOSE CARLOS RODRIGUES', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'JOSE DOS SANTOS', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'NIVALMIR ANTUNES', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'REGINALDO MENDES OLIVEIRA', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'SERGIO APARECIDO GIRALDELLO', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'VILSON TOMACHAK', 'TIPO': 'RODOTREM', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'DENNER DOS SANTOS', 'TIPO': 'TRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'DIEGO FRANCISCO DE SOUZA', 'TIPO': 'TRUCK', 'BASE': 'CIANORTE'}, {'MOTORISTAS': 'ALEX DOUGLAS LOPES ALONSO', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'ANDERSON DE SOUZA SOARES GOMES', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'ANDERSON NUBIATO RODRIGUES DA SILVA', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'ANGELA MARIA GONÇALVES', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'ANTONIO ROBERTO BELTRAMINI', 'TIPO': 'CARRETA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'CELSO RICARDO RODRIGUES', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'CRISTIAN FABIANO LUIZ DA SILVA', 'TIPO': 'TOCO', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'EDE WILSON RODRIGUES', 'TIPO': 'CARRETA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'EDVALDO GONCALVES', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'FABIO CARLOS ARAUJO DO CARMO', 'TIPO': 'CARRETA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'FERNANDO EMIDIO DE SOUZA LIMA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'GEDIELCIO CARVALHO COSTA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'GILMAR DA SILVA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'GILMAR FERREIRA NEVES', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'GUSTAVO ROBERTO PEREIRA', 'TIPO': 'CARRETA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'JHONE GIMENES SANTOS', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'JOAO VITOR DOS SANTOS', 'TIPO': 'RODOTREM', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'JOSE NILSON MARTINS DE ARAUJO', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'LEANDRO DE OLIVEIRA FERREIRA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'LUIS HENRIQUE SANTIAGO FIALHO', 'TIPO': 'FOLGUISTA', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'MICHEL ANTONIOLI', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'PAULO CESAR VICENTINI', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'TATIANE CAXIMIRO PEREIRA', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'VALDINEY FERREIRA PRIMO', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'WESLEY ZANETTI DE OLIVEIRA', 'TIPO': 'TRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'WILLIAM ANDRADE DE MOURA', 'TIPO': 'BITRUCK', 'BASE': 'GUARARAPES'}, {'MOTORISTAS': 'FRANCISCO DAS CHAGAS CORREA CRISPIM', 'TIPO': 'TRUCK', 'BASE': 'ITAJAI'}, {'MOTORISTAS': 'ROGERIO FRANÇA DOS SANTOS', 'TIPO': 'TRUCK', 'BASE': 'ITAJAI'}, {'MOTORISTAS': 'SILVANO DA SILVA FREITAS', 'TIPO': 'TRUCK', 'BASE': 'ITAJAI'}, {'MOTORISTAS': 'ANTONIO APARECIDO PEREIRA', 'TIPO': 'TRUCK', 'BASE': 'PAULINIA'}, {'MOTORISTAS': 'JOSE AUGUSTO DOS SANTOS', 'TIPO': 'TRUCK', 'BASE': 'PAULINIA'}, {'MOTORISTAS': 'RENATO PEREIRA FRANÇA', 'TIPO': 'RODOTREM', 'BASE': 'PAULINIA'}, {'MOTORISTAS': 'AGUINALDO DOS SANTOS TEIXEIRA', 'TIPO': 'RODO ENTREGA', 'BASE': 'SAO JOSE DOS CAMPOS'}, {'MOTORISTAS': 'KERLEI MIRANDA MARTINS', 'TIPO': 'TRUCK', 'BASE': 'SAO JOSE DOS CAMPOS'}, {'MOTORISTAS': 'TADEU JOSE CAETANO DE SOUZA', 'TIPO': 'TRUCK', 'BASE': 'SAO JOSE DOS CAMPOS'}, {'MOTORISTAS': 'RONAN ROMULO ANTUNES', 'TIPO': 'RODO ENTREGA', 'BASE': 'SAO JOSE DOS CAMPOS'}, {'MOTORISTAS': 'SIDNEI DE OLIVEIRA MARIANO', 'TIPO': 'CARRETA', 'BASE': 'SARANDI'}, {'MOTORISTAS': 'NIVALDO REIS MACHADO', 'TIPO': 'BITRUCK', 'BASE': 'UBERABA'}, {'MOTORISTAS': 'SIDNEY RODRIGUES FERREIRA', 'TIPO': 'BITRUCK', 'BASE': 'UBERABA'}, {'MOTORISTAS': 'WELLINGTON DE MELO BATISTA', 'TIPO': 'BITRUCK', 'BASE': 'UBERABA'}, {'MOTORISTAS': 'HIGOR GABRIEL OLIVEIRA BITU', 'TIPO': 'BITRUCK', 'BASE': 'UBERLANDIA'}, {'MOTORISTAS': 'JOSE DONIZETE FERREIRA GOMES', 'TIPO': 'BITRUCK', 'BASE': 'UBERLANDIA'}, {'MOTORISTAS': 'ANTONIO JOSE DE SOUZA MARTINS', 'TIPO': 'BITREM', 'BASE': 'VARZEA GRANDE'}, {'MOTORISTAS': 'JORGE SANTOS DA SILVA', 'TIPO': 'RODOTREM', 'BASE': 'VARZEA GRANDE'}, {'MOTORISTAS': 'MARCOS ROBERTO DOS SANTOS', 'TIPO': 'RODOTREM', 'BASE': 'VARZEA GRANDE'}, {'MOTORISTAS': 'OTAVIO ROSA FRANCO', 'TIPO': 'RODOTREM', 'BASE': 'VARZEA GRANDE'}]

# Alguns cadastros atuais foram apenas correções de nome de pessoas já existentes
# no cadastro legado. A chave canônica evita duplicar essas pessoas durante a migração.
ALIASES_MOTORISTAS_MIGRACAO = {
    "WANDERLEY LOPES DA SILVA": "WANDERLEY LOPES SILVA",
    "MARCOS ROBERTO DOS SANTOS ROSA": "MARCOS ROBERTO DOS SANTOS",
}

def _chave_migracao_motorista(nome: str) -> str:
    nome_n = DataUtils.normalizar_texto(nome)
    return ALIASES_MOTORISTAS_MIGRACAO.get(nome_n, nome_n)


def migrar_cadastro_legado_uma_vez() -> int:
    """Povoa o cadastro persistente com o snapshot legado e os cadastros atuais.

    Não lê Pasta4.xlsx. O snapshot está incorporado no código exclusivamente para
    permitir a transição única. Os dados já existentes na Gestão de Cadastros têm
    prioridade sobre nome/categoria/filial do legado.
    """
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

        # Snapshot legado; não depende do arquivo Pasta4.
        legado = pd.DataFrame(LEGACY_CADASTRO_SNAPSHOT, columns=["MOTORISTAS", "TIPO", "BASE"])
        legado["MOTORISTAS"] = legado["MOTORISTAS"].apply(DataUtils.normalizar_texto)
        legado["TIPO"] = legado["TIPO"].apply(DataUtils.normalizar_texto).replace({"TOCO":"TRUCK"})
        legado["BASE"] = legado["BASE"].apply(DataUtils.normalizar_texto)

        # Cadastro legado primeiro.
        mesclado = {}
        for _, r in legado.iterrows():
            chave = _chave_migracao_motorista(r["MOTORISTAS"])
            if chave:
                mesclado[chave] = {
                    "MOTORISTAS": r["MOTORISTAS"],
                    "TIPO": r["TIPO"],
                    "BASE": r["BASE"],
                }

        # Cadastro atual tem prioridade e também pode corrigir o nome do legado.
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
        # Faz a migração quando ainda não há o conjunto completo do cadastro persistente.
        # Depois que chegar ao conjunto consolidado, reexecutar é idempotente.
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
        _salvar_csv_com_backup(df, caminho)
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

    # Fonte oficial da categoria para os eventos: resultado do cálculo do motorista.
    # Isso evita VALOR/EVENTO = R$ 0,00 quando o extrato automático não traz categoria.
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


def carregar_datas_motoristas() -> dict:
  """Carrega datas de contratação dos motoristas."""
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
  """Salva/atualiza a data de contratação de um motorista."""
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
  _salvar_csv_com_backup(df, ARQUIVO_DATAS_MOTORISTAS)


def carregar_codigos_funcionais() -> dict:
  """Carrega os códigos funcionais persistidos por motorista."""
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
  """Salva/atualiza o código funcional de um motorista."""
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
  _salvar_csv_com_backup(df, ARQUIVO_CODIGOS_FUNCIONAIS)


def atualizar_data_inativacao_motorista(motorista: str, data_inativacao: str):
  """Atualiza a data de inativação sem alterar o status."""
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
    _salvar_csv_com_backup(df, ARQUIVO_INATIVOS)


def alternar_inativo(tipo: str, valor: str, inativar: bool = True, data_inativacao: str = ""):
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
      data_atual = str(data_inativacao or "").strip() or datetime.now().strftime("%d/%m/%Y")
      novo_reg = pd.DataFrame(
          [{"TIPO": tipo, "VALOR": valor, "DATA_INATIVACAO": data_atual}]
      )
      df = pd.concat([df, novo_reg], ignore_index=True)
  else:
    df = df[~mask]

  try:
    _salvar_csv_com_backup(df, ARQUIVO_INATIVOS)
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

    # A base fictícia de abastecimentos só é criada quando não existe
    # absolutamente nenhuma fonte real (legado, mensal ou histórico).
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
    # O abastecimento pode vir do legado ou, preferencialmente, dos arquivos
    # mensais abastecimentos_MMYYYY.xlsx. Não exigimos o legado quando já existe
    # qualquer arquivo mensal ou histórico interno.
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
  """Converte datas/horas, inclusive serial do Excel, preservando o horário."""
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
    """
    Cadastro oficial dos motoristas: a fonte única é a Gestão de Cadastros.

    A partir desta versão, o antigo arquivo-base de motoristas não participa mais
    do carregamento, cálculo, filtros ou relatório RH. O cadastro persistente é o arquivo
    interno ARQUIVO_MOTORISTAS_CUSTOM, alimentado pela tela Gestão de Cadastros.
    """
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

    # Garante estrutura mínima do cadastro persistente.
    for c in ["MOTORISTAS", "TIPO", "BASE"]:
      if c not in df_custom.columns:
        df_custom[c] = ""

    cadastro = pd.DataFrame({
        "MOTORISTA_CADASTRO": df_custom["MOTORISTAS"].apply(DataUtils.normalizar_texto),
        "TIPO_CADASTRO": df_custom["TIPO"].apply(DataUtils.normalizar_texto).replace({"TOCO": "TRUCK"}),
        "BASE_CADASTRO": df_custom["BASE"].apply(DataUtils.normalizar_texto),
    })

    # Elimina linhas vazias e mantém um único cadastro por motorista.
    cadastro = cadastro[
        (cadastro["MOTORISTA_CADASTRO"] != "")
        & (cadastro["TIPO_CADASTRO"] != "")
    ].copy()

    cadastro["EH_FOLGUISTA"] = cadastro["TIPO_CADASTRO"].eq("FOLGUISTA")
    cadastro = cadastro.drop_duplicates("MOTORISTA_CADASTRO", keep="last").reset_index(drop=True)

    # Dados complementares administrados pela própria Gestão de Cadastros.
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
    """Carrega automaticamente os arquivos mensais "viagens MM_YYYY.xlsx".

    O cruzamento usa placa + faixa de odômetro e prioriza a janela de data/hora da viagem.
    Arquivos de todos os meses disponíveis são aceitos para cobrir viagens iniciadas em
    um mês e abastecimentos ocorridos no mês seguinte.
    """
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
          # Normaliza viagens que eventualmente estejam registradas com ODM invertido.
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
            + (hora_series.loc[mascara_hora] - pd.Timestamp("1899-12-30")).where(
                hora_series.loc[mascara_hora] < pd.Timestamp("1900-01-01"),
                pd.to_timedelta(hora_series.loc[mascara_hora].dt.hour * 3600 + hora_series.loc[mascara_hora].dt.minute * 60 + hora_series.loc[mascara_hora].dt.second, unit="s"),
            )
        )
        # Caso a coluna de hora tenha vindo como datetime-base do Excel, usa somente o horário.
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
  """Concilia cada abastecimento com o ciclo de viagens entre dois abastecimentos.

  Regra operacional da ficha manual:
  - o abastecimento atual normalmente fecha o ciclo iniciado no abastecimento anterior;
  - o motorista responsável é identificado pelas viagens realizadas entre os dois KM de
    abastecimento, usando placa + faixa de odômetro e janela temporal;
  - se todas as viagens candidatas do ciclo pertencem a um único motorista, a atribuição
    é automática;
  - se houver mais de um motorista possível no mesmo ciclo, o sistema não "chuta":
    mantém o motorista original e marca PENDENTE DE VALIDAÇÃO;
  - o KM/média continuam sendo calculados pela sequência da placa, independentemente
    do motorista que ficou associado ao abastecimento.
  """
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

  # A sequência por placa é a referência do ciclo. Ordenamos todas as linhas para
  # encontrar o abastecimento imediatamente anterior, inclusive de outra competência.
  base["_TMP_DT"] = pd.to_datetime(base.get("DATA_HORA_ABASTECIMENTO", base.get("DATA_NUM")), errors="coerce")
  base["_TMP_DT"] = base["_TMP_DT"].apply(lambda x: _parse_datetime_flex(x) if pd.notna(x) else pd.NaT)
  base = base.sort_values(["PLACA_PADRONIZADA", "KM_ATUAL_NUM", "_TMP_DT"], kind="stable")

  for placa, idxs in base.groupby("PLACA_PADRONIZADA", sort=False).groups.items():
    idx_list = list(idxs)
    # Mantém a ordem temporal/odômetro para identificar o abastecimento anterior.
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

      # O ciclo é o intervalo entre o KM do abastecimento anterior e o KM do atual.
      # Uma viagem candidata deve cruzar esse intervalo e ficar temporalmente entre
      # os abastecimentos, admitindo até 2 dias após o fim da viagem para o fechamento
      # do tanque (caso mostrado na ficha: viagem termina dia 30 e abastece dia 31).
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

      # Soma a cobertura de KM do ciclo por motorista. Se somente um motorista
      # executou o ciclo, a associação é segura mesmo com várias pernas (carregado/vazio).
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
        # Não atribuir automaticamente quando o mesmo ciclo contém mais de um motorista.
        # O abastecimento permanece com o motorista original, mas fica explícito para conferência.
        base.at[idx, "STATUS_VALIDACAO_VIAGEM"] = "PENDENTE DE VALIDAÇÃO — MÚLTIPLOS MOTORISTAS"
        mot_principal = max(coberturas, key=coberturas.get) if coberturas else ""
        # Mostra a maior cobertura apenas como referência, sem alterar o responsável.
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
  # O cálculo de prêmio usa o motorista considerado; nos casos pendentes, permanece o original.
  base["CONDUTOR_NORMALIZADO"] = base["MOTORISTA_CONSIDERADO"].apply(DataUtils.normalizar_texto)
  return base


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

      manual_mask = grupo["CATEGORIA_MANUAL"].astype(str).str.strip() != ""
      categorias_abastecimento = [
          str(x).strip().upper()
          for x in grupo.loc[grupo["CATEGORIA_ABASTECIMENTO"].notna(), "CATEGORIA_ABASTECIMENTO"]
          if str(x).strip()
      ]
      categorias_unicas = sorted(set(categorias_abastecimento))

      if manual_mask.any():
        # Mapeamento manual por placa tem prioridade absoluta.
        grupo["CATEGORIA_ELEGIVEL"] = np.where(
            manual_mask,
            grupo["CATEGORIA_MANUAL"],
            grupo["CATEGORIA_MANUAL"].replace("", np.nan).ffill().bfill().fillna(
                tipo_cad if tipo_cad else (categorias_unicas[0] if categorias_unicas else "TRUCK")
            ),
        )
        grupo["USA_CATEGORIA_MANUAL"] = manual_mask
      elif len(categorias_unicas) == 1:
        # Quando todos os abastecimentos do motorista são de uma única categoria,
        # usamos a categoria real da frota/abastecimento. Não deixamos o
        # TIPO_CADASTRO sobrescrever a categoria efetivamente trabalhada.
        grupo["CATEGORIA_ELEGIVEL"] = categorias_unicas[0]
        grupo["USA_CATEGORIA_MANUAL"] = False
      elif eh_folguista:
        # Folguista sem mapeamento manual: usa a categoria em que rodou mais KM.
        soma = grupo.groupby("CATEGORIA_ABASTECIMENTO")["KM_CONSUMO"].sum()
        cat_elegivel = (
            soma.idxmax()
            if not soma.empty
            else (tipo_cad if tipo_cad else "TRUCK")
        )
        grupo["CATEGORIA_ELEGIVEL"] = cat_elegivel
        grupo["USA_CATEGORIA_MANUAL"] = False
      else:
        # Para múltiplas categorias sem mapeamento manual, mantemos a categoria
        # cadastral como fallback para não alterar a regra já existente.
        grupo["CATEGORIA_ELEGIVEL"] = tipo_cad if tipo_cad else grupo["CATEGORIA_ABASTECIMENTO"].iloc[0]
        grupo["USA_CATEGORIA_MANUAL"] = False

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
def _dias_ausencia_por_competencia(
    df_ausencias: pd.DataFrame,
    data_inicio_comp,
    data_fim_comp,
) -> dict:
    """Distribui cada afastamento somente pelos dias que caem na competência 26-25.

    Registros antigos sem DATA_FIM usam DATA_INICIO + DIAS - 1 como fim.
    Retorna um mapa MOTORISTA normalizado -> dias de ausência dentro da competência.
    """
    if df_ausencias is None or df_ausencias.empty or "MOTORISTA" not in df_ausencias.columns:
        return {}

    ini_comp = parse_data_filtro(data_inicio_comp)
    fim_comp = parse_data_filtro(data_fim_comp)
    if ini_comp is None or fim_comp is None:
        return {}

    ini_comp = pd.Timestamp(ini_comp).normalize()
    fim_comp = pd.Timestamp(fim_comp).normalize()
    acumulado = {}

    for _, row in df_ausencias.iterrows():
        mot = DataUtils.normalizar_texto(row.get("MOTORISTA", ""))
        if not mot:
            continue

        dt_ini = parse_data_filtro(row.get("DATA_INICIO", ""))
        if dt_ini is None or pd.isna(dt_ini):
            continue
        dt_ini = pd.Timestamp(dt_ini).normalize()

        dt_fim = parse_data_filtro(row.get("DATA_FIM", ""))
        if dt_fim is None or pd.isna(dt_fim):
            dias = pd.to_numeric(row.get("DIAS", 0), errors="coerce")
            try:
                dias = int(dias)
            except Exception:
                dias = 0
            if dias <= 0:
                continue
            dt_fim = dt_ini + pd.Timedelta(days=dias - 1)
        else:
            dt_fim = pd.Timestamp(dt_fim).normalize()

        if dt_fim < dt_ini:
            continue

        inicio_intersecao = max(dt_ini, ini_comp)
        fim_intersecao = min(dt_fim, fim_comp)
        if inicio_intersecao <= fim_intersecao:
            dias_intersecao = int((fim_intersecao - inicio_intersecao).days + 1)
            acumulado[mot] = acumulado.get(mot, 0) + dias_intersecao

    return acumulado


def filtrar_desclassificacoes_competencia(df_desclassificacoes: pd.DataFrame, data_inicio_comp=None, data_fim_comp=None) -> pd.DataFrame:
  """Retorna somente desclassificações lançadas dentro da competência 26-25.

  A competência é inclusiva no início e no fim. Registros antigos sem DATA_EVENTO
  são mantidos no histórico, mas não participam do cálculo mensal até receberem
  uma data, evitando que uma desclassificação de um mês seja carregada para outro.
  """
  if df_desclassificacoes is None or df_desclassificacoes.empty:
    return pd.DataFrame() if df_desclassificacoes is None else df_desclassificacoes.copy()
  df = df_desclassificacoes.copy()
  if data_inicio_comp is None or data_fim_comp is None or "DATA_EVENTO" not in df.columns:
    return df
  ini = pd.Timestamp(data_inicio_comp).normalize()
  fim = pd.Timestamp(data_fim_comp).normalize()
  datas = df["DATA_EVENTO"].apply(parse_data_filtro)
  mask = datas.notna() & (datas >= ini) & (datas <= fim)
  return df.loc[mask].copy()


def aplicar_regras_gerais(
    df_resumo_original: pd.DataFrame,
    df_ausencias: pd.DataFrame,
    df_desclassificacoes: pd.DataFrame,
    data_inicio_comp=None,
    data_fim_comp=None,
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
    if data_inicio_comp is not None and data_fim_comp is not None:
        soma_dias = _dias_ausencia_por_competencia(
            df_ausencias, data_inicio_comp, data_fim_comp
        )
    else:
        # Compatibilidade com chamadas antigas: soma total dos afastamentos.
        df_aus_tmp = df_ausencias.copy()
        df_aus_tmp["DIAS"] = pd.to_numeric(
            df_aus_tmp["DIAS"], errors="coerce"
        ).fillna(0)
        soma_dias = {
            DataUtils.normalizar_texto(k): int(v)
            for k, v in df_aus_tmp.groupby("MOTORISTA")["DIAS"].sum().to_dict().items()
        }

    res["DIAS_AUSENCIA"] = (
        res["MOTORISTA"].apply(DataUtils.normalizar_texto).map(soma_dias).fillna(0).astype(int)
    )
    # A régua de prêmio continua sendo baseada em 30 dias.
    # O diferencial é que os dias de afastamento agora pertencem somente à competência correta.
    res["DIAS_EFETIVOS"] = np.maximum(0, 30 - res["DIAS_AUSENCIA"])
    res["PREMIO"] = res.apply(
        lambda r: max(0.0, r["PREMIO_BRUTO"] * (r["DIAS_EFETIVOS"] / 30.0)),
        axis=1,
    )
  else:
    res["DIAS_AUSENCIA"] = 0
    res["DIAS_EFETIVOS"] = 30
    res["PREMIO"] = res["PREMIO_BRUTO"]

  # Desclassificações são eventos mensais: somente lançamentos da competência
  # atual podem afetar o prêmio. Isso impede que um evento de agosto permaneça
  # ativo na competência seguinte.
  df_desclassificacoes_comp = filtrar_desclassificacoes_competencia(
      df_desclassificacoes, data_inicio_comp, data_fim_comp
  )

  if (
      not df_desclassificacoes_comp.empty
      and "MOTORISTA" in df_desclassificacoes_comp.columns
  ):
    for idx in res.index:
      m_nome = res.at[idx, "MOTORISTA"]
      g = df_desclassificacoes_comp[df_desclassificacoes_comp["MOTORISTA"] == m_nome]
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
  """Gera o relatório do RH a partir do cadastro oficial completo.

  Regra: todo motorista ATIVO no cadastro oficial deve aparecer no RH,
  mesmo sem abastecimento, sem consumo válido, em férias ou atestado.
  Quando não houver prêmio calculado para a competência, o valor fica R$ 0,00.
  A filial e o código funcional vêm exclusivamente da Gestão de Cadastros.
  """
  colunas_saida = ["CÓDIGO FUNCIONAL", "NOME", "FILIAL", "VALOR PAGO"]

  try:
    cad_rh = (cadastro_all.copy() if isinstance(globals().get("cadastro_all"), pd.DataFrame)
              else cadastro.copy() if isinstance(globals().get("cadastro"), pd.DataFrame)
              else pd.DataFrame())
    if cad_rh.empty:
      return pd.DataFrame(columns=colunas_saida)

    # Somente motoristas ativos entram no RH. Férias/atestado não alteram o status cadastral.
    if "STATUS" in cad_rh.columns:
      cad_rh = cad_rh[
          cad_rh["STATUS"].fillna("ATIVO").astype(str).str.strip().str.upper().eq("ATIVO")
      ].copy()

    # Usuário de consulta vê somente os ativos da filial autorizada.
    if not is_admin and FILIAL_ACESSO not in ("", "TODAS") and "BASE_CADASTRO" in cad_rh.columns:
      cad_rh = cad_rh[
          cad_rh["BASE_CADASTRO"].apply(DataUtils.normalizar_texto) == FILIAL_ACESSO
      ].copy()

    cad_rh["_MOTORISTA_RH"] = cad_rh["MOTORISTA_CADASTRO"].apply(DataUtils.normalizar_texto)
    cad_rh = cad_rh[cad_rh["_MOTORISTA_RH"].ne("")].copy()
    cad_rh = cad_rh.drop_duplicates("_MOTORISTA_RH", keep="last")

    # Resultado da competência, quando existir. Ausentes do cálculo recebem zero.
    calc = df_resumo.copy() if isinstance(df_resumo, pd.DataFrame) else pd.DataFrame()
    if not calc.empty and "MOTORISTA" in calc.columns:
      calc["_MOTORISTA_RH"] = calc["MOTORISTA"].apply(DataUtils.normalizar_texto)
      if "PREMIO" not in calc.columns:
        calc["PREMIO"] = 0.0
      calc["PREMIO"] = pd.to_numeric(calc["PREMIO"], errors="coerce").fillna(0.0)
      # O prêmio já é o valor final após regras/descontos.
      premio_por_motorista = calc.groupby("_MOTORISTA_RH", as_index=False)["PREMIO"].sum()
    else:
      premio_por_motorista = pd.DataFrame(columns=["_MOTORISTA_RH", "PREMIO"])

    rh = cad_rh[["_MOTORISTA_RH", "MOTORISTA_CADASTRO", "BASE_CADASTRO", "CODIGO_FUNCIONAL"]].copy()
    rh = rh.merge(premio_por_motorista, on="_MOTORISTA_RH", how="left")
    rh["PREMIO"] = pd.to_numeric(rh["PREMIO"], errors="coerce").fillna(0.0)
    rh["CODIGO_FUNCIONAL"] = rh["CODIGO_FUNCIONAL"].fillna("").astype(str).str.strip()

    # Fallback somente para códigos já cadastrados no arquivo auxiliar durante a transição.
    codigos_fallback = carregar_codigos_funcionais()
    rh["CODIGO_FUNCIONAL"] = rh.apply(
        lambda r: r["CODIGO_FUNCIONAL"] if r["CODIGO_FUNCIONAL"] else codigos_fallback.get(r["_MOTORISTA_RH"], ""),
        axis=1,
    )

    def formatar_valor_pago(x):
      try:
        return f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
      except Exception:
        return "R$ 0,00"

    out = pd.DataFrame({
        "CÓDIGO FUNCIONAL": rh["CODIGO_FUNCIONAL"],
        "NOME": rh["MOTORISTA_CADASTRO"],
        "FILIAL": rh["BASE_CADASTRO"].fillna("").astype(str),
        "VALOR PAGO": rh["PREMIO"].map(formatar_valor_pago),
    })
    return out.sort_values(["FILIAL", "NOME"], kind="stable").reset_index(drop=True)
  except Exception as exc:
    print(f"Erro ao gerar tabela RH completa: {exc}")
    return pd.DataFrame(columns=colunas_saida)


def estilizar_rh_zerados(df: pd.DataFrame):
  """Destaca em vermelho os valores zerados no relatório RH."""
  if df is None or df.empty:
    return df

  def _zero(v):
    try:
      texto = str(v).strip().upper().replace("R$", "").replace(" ", "")
      texto = texto.replace(".", "").replace(",", ".")
      return float(texto) == 0.0
    except Exception:
      return False

  def _estilo(v):
    return "color: red; font-weight: 700" if _zero(v) else ""

  styler = df.style
  if "VALOR PAGO" in df.columns:
    styler = styler.map(_estilo, subset=["VALOR PAGO"])
  return styler


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
      resumo_periodo,
      df_ausencias,
      df_desclassificacoes,
      data_inicio_comp=d_i,
      data_fim_comp=d_f,
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

  # KPI de motoristas: fonte exclusiva = cadastro oficial completo.
  # Não depende de abastecimento, consumo, férias, atestado ou prêmio calculado.
  cad_kpi = cadastro_all.copy() if isinstance(globals().get("cadastro_all"), pd.DataFrame) else (cadastro.copy() if isinstance(cadastro, pd.DataFrame) else pd.DataFrame())
  if not is_admin and FILIAL_ACESSO not in ("", "TODAS"):
    cad_kpi = cad_kpi[
        cad_kpi["BASE_CADASTRO"].apply(DataUtils.normalizar_texto) == FILIAL_ACESSO
    ].copy()
  if not cad_kpi.empty:
    if "STATUS" in cad_kpi.columns:
      cad_kpi = cad_kpi[
          cad_kpi["STATUS"].fillna("ATIVO").astype(str).str.strip().str.upper().eq("ATIVO")
      ].copy()
    if motorista and motorista != "TODOS":
      m_kpi = DataUtils.normalizar_texto(motorista)
      cad_kpi = cad_kpi[
          cad_kpi["MOTORISTA_CADASTRO"].apply(DataUtils.normalizar_texto) == m_kpi
      ]
    if filial and filial != "TODAS" and (is_admin or FILIAL_ACESSO in ("", "TODAS")):
      f_kpi = DataUtils.normalizar_texto(filial)
      cad_kpi = cad_kpi[
          cad_kpi["BASE_CADASTRO"].apply(DataUtils.normalizar_texto) == f_kpi
      ]
    tot_mots = int(cad_kpi["MOTORISTA_CADASTRO"].nunique())
  else:
    tot_mots = 0

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

  eventos_jornada = int(pd.to_numeric(row_data.get("EVENTOS_CONTROLE_JORNADA", 0), errors="coerce") or 0)
  eventos_excesso = int(pd.to_numeric(row_data.get("EVENTOS_EXCESSO_VELOCIDADE", 0), errors="coerce") or 0)
  desconto_jornada = float(pd.to_numeric(row_data.get("DESCONTO_CONTROLE_JORNADA", 0), errors="coerce") or 0)
  desconto_excesso = float(pd.to_numeric(row_data.get("DESCONTO_EXCESSO_VELOCIDADE", 0), errors="coerce") or 0)
  total_controles = eventos_jornada + eventos_excesso
  valor_total_controles = desconto_jornada + desconto_excesso
  percentual_controles = (100.0 if eventos_jornada >= 130 or eventos_excesso > 30 else 0.0)
  valor_total_controles_str = (
      f"R$ {valor_total_controles:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
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
        </table>
    </div>
    """
