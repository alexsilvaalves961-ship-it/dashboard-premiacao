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
# IMPORTAÇÕES DE AUTOMAÇÃO (SELENIUM) CONFIGURADAS PARA A NUVEM
# ================================================================
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ================================================================
# AUTOMAÇÃO DE EXTRAÇÃO DE ABASTECIMENTOS (PORTAL CIAPETRO)
# ================================================================
def executar_automacao_abastecimentos_ciapetro(mes_ano_str="08_2026"):
  """Acessa o portal da frota da Ciapetro em modo headless (nuvem),

  realiza o login e baixa os dados de abastecimento salvando no formato esperado.
  """
  url_login = "https://portalfrota.uauh.com.br/app/_abastecimento/"
  cliente = "ciapetro"
  usuario = "administrativo.transporte@ciapetro.com.br"
  senha = "123456"

  options = webdriver.ChromeOptions()
  options.add_argument("--headless")
  options.add_argument("--no-sandbox")
  options.add_argument("--disable-dev-shm-usage")
  options.add_argument("--disable-gpu")
  options.add_argument("--remote-debugging-port=9222")

  try:
    driver = webdriver.Chrome(options=options)
  except Exception as e:
    print(f"Erro ao iniciar o ChromeDriver: {e}")
    return False

  try:
    print("1. Acessando o portal da frota...")
    driver.get(url_login)
    time.sleep(3)

    wait = WebDriverWait(driver, 15)

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

    driver.get(url_login)
    time.sleep(5)

    print("3. Coletando os dados da tabela de abastecimento...")
    tabelas = pd.read_html(driver.page_source)

    if tabelas:
      df_abastecimentos = tabelas[0]
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
# CONFIGURAÇÃO DE CAMINHOS E PERSISTÊNCIA
# ================================================================
DATA_DIR = os.getenv("DATA_DIR", ".")


def garantir_diretorio():
  if DATA_DIR and DATA_DIR != "." and not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR, exist_ok=True)


# ================================================================
# UTILITÁRIOS E TRATAMENTO DE TEXTO / PLACAS / DATAS
# ================================================================
class DataUtils:

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
  dias_ef = row_data.get("DIAS_EFETIVOS", 30)
  dias_efetivos_val = 30 if pd.isna(dias_ef) else int(dias_ef)

  return f"""
    <div class="recibo-card" style="background-color: #FFFFFF; padding: 28px; border-radius: 12px; max-width: 650px; margin: 0 auto 20px auto; font-family: Arial, sans-serif; color: #000000; border: 1px solid #CBD5E1; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);">
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
        </table>
    </div>
    """


# ================================================================
# INTERFACE PRINCIPAL DO STREAMLIT
# ================================================================
st.set_page_config(page_title="Dashboard de Premiação", layout="wide")

st.title("Dashboard de Premiação - Ciapetro")
st.sidebar.header("Painel de Controlo")

if st.sidebar.button("Executar Automação da Ciapetro"):
  with st.spinner("A aceder ao portal e a descarregar dados..."):
    sucesso = executar_automacao_abastecimentos_ciapetro("08_2026")
    if sucesso:
      st.sidebar.success("Dados de abastecimento atualizados com sucesso!")
    else:
      st.sidebar.error(
          "Falha na automação. Verifique os logs para mais detalhes."
      )

st.write(
    "Bem-vindo ao sistema de cálculo de premiações de frota. Utilize o menu"
    " lateral para acionar as rotinas automáticas."
)
