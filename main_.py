import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import time
import re
from dotenv import load_dotenv
# from decouple import Config
# config = Config()
import tempfile
import pandas as pd
import numpy as np
import zipfile
import warnings
from pandas.api.types import is_float_dtype
import locale
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import numbers
from openpyxl.styles import Border, Side
from pandas.tseries.offsets import BDay
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')  # Tenta definir para pt_BR
except locale.Error:
    locale.setlocale(locale.LC_ALL, 'en_US.UTF-8') 

temp_dir = tempfile.gettempdir()
download_path = f"{temp_dir}/Teste_FPR"
st.markdown(
    """
    <style>

    /* Fundo da página */
    .stApp {
        background-color: #064635; 
        color: #ffffff;
    }
    /* Centralizar o título */
    h1 {
        text-align: center;
        color: #ffffff;
    }

    /* Estilizar o texto acima das labels */
    label {
        color: #ffffff !important;
        font-weight: bold;
    }

    /* Estilizar o botão */
    div.stButton > button {
        color: #000000; 
        background-color: #ffffff;
        padding: 10px 20px; 
        font-weight: bold; 
        border-radius: 8px;
    }

    /* Alterar a cor do ícone de ajuda para branco */
    div[data-testid="stTooltipIcon"] svg {
        color: #ffffff; 
        fill: #ffffff; 
    }
    /* Estilizar mensagens de sucesso */
    div.stAlert {
        background-color: #ffffff;
        color: #ffffff !important;
        border-radius: 10px; 
        padding: 10px; 
        font-weight: bold; 
        box-shadow: 2px 2px 5px rgba(0, 0, 0, 0.1);
    }

    div.stDownloadButton > button {
        color: #000000; /* Texto preto */
        background-color: #ffffff; /* Fundo branco */
        padding: 10px 20px; /* Margens internas */
        font-weight: bold; /* Texto em negrito */
        border-radius: 8px; /* Bordas arredondadas */
    }
    
    </style>
    <div>
        <h1>Extração dos fundos no <br> ambiente AMPLIS</h1>
    </div>
    """,
    unsafe_allow_html=True
)

#### ---- Funções ---- ####
def tratar_valores(valor):
        if '(' in valor and ')' in valor:
            valor = valor.replace('(', '').replace(')', '')
            return -float(valor.replace('.', '').replace(',', '.'))
        return float(valor.replace('.', '').replace(',', '.'))

def ler (caminho):
        if caminho.name.endswith('.csv'):
            temp1 = pd.read_csv(
                caminho,
                encoding='latin-1',
                low_memory=False,
                decimal=',',
                sep=';'
                )
        elif caminho.name.endswith('.xlsx'):
            temp1 = pd.read_excel(
                caminho,
                engine='openpyxl'  # Certifique-se de ter instalado o openpyxl
            )
        else:
            raise ValueError("Formato de arquivo não suportado. Envie um arquivo CSV ou XLSX.")
        return temp1
        
def calcular_atv_probl(row):
        if row['ATRASO'] > 90:
            proporcao = row['PDD_PROPORCIONAL'] / row['VP_PROPORCIONAL']
            if proporcao < 0.2:
                return 1.5
            elif proporcao < 0.5:
                return 1
            else:
                return 0.5
        else:
            return "N"

def classificar_pessoa(doc_sacado):
        if len(doc_sacado) == 18:
            return "PJ"
        else:
            return "PF"

def atualizar_rwa(dataframe):
        dataframe['RWA'] = dataframe['SALDO'] * dataframe['FPR']
        return dataframe

def ativ_probl(row, data):
    data = pd.to_datetime(data, format='%d/%m/%Y', dayfirst=True)
    if (data - pd.to_datetime(row['DATA_EMISSAO_2'])).days > 90:
        return 'S'
    else:
        return 'N'

# Carrega o arquivo .env
load_dotenv()

# ----- Inputs ----- #
data = st.text_input(
    "Informe a data de extração (dd/mm/YYYY):",
    placeholder="Exemplo: 31/12/2024")

def valores_únicos(caminho_arquivo, coluna):
    
    try:
        df = pd.read_csv(
            caminho_arquivo,
            sep=';', 
            encoding='latin1', 
            on_bad_lines='skip' 
            )
        valores = df[coluna].dropna().unique()
        return valores
    
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        return []

@st.cache_data
# Validação de datas:
def validar_data(data):
    '''
    Verificação se, a string informada pelo usuário, segue o padrão
    do ambiente AMPLIS. Caso a data informada não siga o padrão dd/mm/YYYY,
    a expressão retorna None. 
    '''
    return bool(re.match(r"^\d{2}/\d{2}/\d{4}$", data))

# ----- WebScraping ----- #
def extração_amplis(data):
    # Verificação se o diretório existe e, caso ele não exista, cria o diretório.
    if not os.path.exists(download_path):
        # Criação do diretório.
        os.makedirs(download_path)

    # Configurações específicas para o navegador Google Chrome no contexto do Selenium WebDriver.
    # Personalização do comportamento do navegador durante a automação.
    chrome_options = webdriver.ChromeOptions() # Criação do objeto, que permite configurar as opções do navegador.
    prefs = {
        "download.default_directory": download_path, # Definição do diretório padrão onde os arquivos baixados serão salvos
        "download.prompt_for_download": False, # Desativação do prompt que aparece ao iniciar um download
        "profile.default_content_settings.popups": 0, # Desativação dos pop-ups relacionados ao download
        "safebrowsing.enabled": True, # Proteção contra sites maliciosos
    }
    chrome_options.add_experimental_option(
        "prefs", # O Chrome possui diversas "opções experimentais", e "prefs" é uma delas,
                 # usada para definir preferências de comportamento do navegador.
          prefs
          ) # Adição das preferências
    
    # Inicialização
    serviço = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(
        service=serviço,
        options=chrome_options
    )

    # Acesso a página do AMPLIS:
    driver.get("https://idltrust.totvs.amplis.com.br/amplis/login/SEC_00001.jsf")
    driver.maximize_window() # Maximização da tela.

    # Criação de uma espera explícita; aguarda até uma condição específica.
    WebDriverWait(
        driver=driver,
        timeout=10 # Tempo máximo aguardando
        ).until(EC.presence_of_all_elements_located((By.ID, 'loginForm:userLoginInput:campo'))) # Condição que precisa ser atendida.
    # Localização do elemento na página com base no ID.
    login_input = driver.find_element(By.ID, 
                                      'loginForm:userLoginInput:campo')
    # Configuração do Login
    login = "JBENETON"
    login_input.send_keys(login)
    
    time.sleep(1)
    
    # Configuração do Password
    password_input = driver.find_element(
        By.ID,
        'loginForm:userPasswordInput'
    )
    password = "Qista@2025"
    password_input.send_keys(password)

    # Configuração de acesso
    entrar_button = WebDriverWait(
        driver=driver,
        timeout=10
    ).until(EC.element_to_be_clickable((By.ID, 'loginForm:botaoOk'))) # Condição que o botão esteja "clicável"
    # Clicando no Botão usando JavaScript
    # entrar_button.click()
    driver.execute_script("arguments[0].click();", entrar_button)
    time.sleep(5)

    # ----- Ir até a Carteira Diária ----- #
    relatório_menu = WebDriverWait(
        driver=driver,
        timeout=10
    ).until(
        EC.presence_of_element_located((By.ID, 'menuRelatoriosButtonSpan'))
    )
    patrimônio_menu = WebDriverWait(
        driver=driver,
        timeout=10
    ).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="dropdownMenuRelatorios"]/li[2]'))
    )

    # Criar um ActionChain para mover o mouse até o elemento "Relatórios"
    action = ActionChains(
        driver=driver
        ).move_to_element(relatório_menu).perform()
    time.sleep(1)

    # Criar um ActionChain para mover o mouse até o elemento "Patrimônio"
    action=ActionChains(
        driver=driver
    ).move_to_element(
        patrimônio_menu
    ).perform()
    time.sleep(1)

    # Criar um ActionChain para mover o mouse até o elemento "Carteira Diária"
    carteira_diária = WebDriverWait(
        driver=driver,
        timeout=10
    ).until(
        EC.element_to_be_clickable((By.ID, "mainForm:j_id_62:1:j_id_6n:0:j_id_82"))
    ).click()

    # ----- Alteração da Data ----- #
    
    # Localizar o ambiente da data inicial.
    período_input = WebDriverWait(
        driver=driver,
        timeout=10
    ).until(
        EC.presence_of_element_located((By.ID, 'mainForm:calendarDateBegin:campoInputDate'))
    )
    
    # Clicar no ambiente.
    período_input.click()
    time.sleep(1)

    # Limpar o ambiente.
    período_input.clear()
    time.sleep(1)

    # Adicionando a data exposta pelo usuário.
    # período_input.send_keys(data)
    driver.execute_script("arguments[0].value = arguments[1];", período_input, data)

    # Localizar o ambiente da data final.
    final_input = WebDriverWait(
        driver=driver,
        timeout=10
    ).until(
        EC.presence_of_element_located((By.ID, 'mainForm:calendarDateEnd:campoInputDate'))
    )
    # Clicar no ambiente.
    final_input.click()
    time.sleep(1)

    # Limpar o ambiente.
    final_input.clear()
    time.sleep(1)

    # Adicionando a data exposta pelo usuário.
    # período_input.send_keys(data)
    driver.execute_script("arguments[0].value = arguments[1];", final_input, data)

    # Enter para a confirmação da escolha
    final_input.send_keys(Keys.RETURN)

    print(f'Data {data} selecionada com sucesso!')
    time.sleep(2)

    # ----- Seleção de todos os Fundos ----- #
    carteira = WebDriverWait(
        driver=driver,
        timeout=10
    ).until(
        EC.presence_of_element_located((By.ID, 'mainForm:portfolioPickList:firstSelect'))
    )
    time.sleep(3)

    # Cria uma instância Select para interagir com a lista
    select_carteira = Select(carteira)
    
    # Seleciona a primeira opção
    primeira = select_carteira.options[0]
    select_carteira.select_by_value(primeira.get_attribute("value"))
    time.sleep(3)
    
    # Click no Botão para atualizar os fundos selecionados
    WebDriverWait(
        driver=driver,
        timeout=10
    ).until(
        EC.element_to_be_clickable((By.ID, 'mainForm:portfolioPickList:includeAll'))
    ).click()
    time.sleep(1)

    # ----- Seleção da Opção: CSV ----- #
    dropdown = driver.find_element(
        By.ID,
        "mainForm:saida:campo"
    )

    # Usar a classe Select para interagir com o dropdown
    select = Select(dropdown)

    # Selecionar a opção "CSV"
    select.select_by_visible_text("CSV")
    time.sleep(5)

    # ----- Download ----- #
    select_button = WebDriverWait(
        driver, 
        10
        ).until(
            EC.presence_of_element_located((By.ID, "mainForm:confirmButton"))
        ).click()

    timeout = 100000 # Tempo máximo em segundos
    start_time = time.time()

    while True:
        # Verificação se há arquivos no diretório
        files = [f for f in os.listdir(download_path) if f.endswith(".zip")]
        if files: # Se encontrar o arquivo CSV.
            print("Download realizado")
            break
        if time.time() - start_time > timeout:
            print("Tempo Limite excedido. Download não concluído.")
            driver.quit()
            exit()
        
        # Aguarda 3 segundo antes de verificar novamente.
        time.sleep(3)
    
    driver.quit()

    ############### FIM ###############
    
    dict = {}
    arquivos = os.listdir(download_path)
    uploaded_file = os.path.join(download_path, arquivos[0])

    if uploaded_file is not None: # atribuir o conteúdo do zip nesta variável.
        try:
            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                zip_ref.extractall(download_path)
            
            # Listar os arquivos extraídos
            extracted_files = os.listdir(download_path)
            st.write("Arquivos Extraídos com sucesso!")
            # st.write(extracted_files)

            for arquivo in extracted_files:
                try:
                    if re.match(r"^\d+", arquivo) and arquivo.endswith('.csv'):
                        key = re.search(r"-(.*)\.csv$", arquivo).group(1)
                        caminho_arquivo = os.path.join(download_path, arquivo)
                        df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin1', on_bad_lines='skip')
                        dict[key] = df
                        print(f"DataFrame para '{key}' adicionado ao dicionário.")
                except Exception as e:
                    print(f"Erro ao processar o arquivo '{arquivo}': {e}")
            
            # st.write("Resumo dos DataFrames carregados:")
            # for key, df in dict.items():
            #     st.write(f"**{key}**: {df.shape[0]} linhas, {df.shape[1]} colunas")
            #     st.dataframe(df)
        
        except zipfile.BadZipFile:
            st.error("O arquivo carregado não é válido.")

def main():
    st.title("Processamento dos arquivos")
    DOWNLOAD_PATH = os.path.join(tempfile.gettempdir(), "Teste_FPR")
    ARQUIVO_RENDA_FIXA = "01-Renda_Fixa.csv"

    # Verificar se o arquivo "01-Renda_Fixa.csv" existe dentro da pasta.
    caminho_renda_fixa = os.path.join(DOWNLOAD_PATH, ARQUIVO_RENDA_FIXA)
    if os.path.exists(caminho_renda_fixa):
        print('Caminho Existente.')

        valores = valores_únicos(caminho_renda_fixa, 'CARTEIRA')
        valor_inicial = "FIDC REAG H Y" if "FIDC REAG H Y" in valores else valores[0]

        if valores.size > 0:
            col_fund1, col_fund2 = st.columns(2)
            with col_fund1:
                opção_selecionada = st.selectbox(
                    'Selecione um fundo:',
                    valores,
                    index=list(valores).index(valor_inicial) if valor_inicial in valores else 0
                )
            with col_fund2:
                data_mes = st.text_input(
                    'Informe a data desejada:',
                    placeholder="Exemplo: 31/12/2024"
                )

            dict = {}
            download_path = f"{temp_dir}/Teste_FPR"
            arquivos = os.listdir(download_path)
            uploaded_file = os.path.join(download_path, arquivos[0])

            if uploaded_file is not None: # atribuir o conteúdo do zip nesta variável.

                extracted_files = os.listdir(download_path)
                # st.write("Arquivos Extraídos: ")
                # st.write(extracted_files)

                for arquivo in extracted_files:
                    try:
                        if re.match(r"^\d+", arquivo) and arquivo.endswith('.csv'):
                            key = re.search(r"-(.*)\.csv$", arquivo).group(1)
                            caminho_arquivo = os.path.join(download_path, arquivo)
                            df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin1', on_bad_lines='skip')
                            df_filtrado = df[df["CARTEIRA"] == opção_selecionada]
                            df_filtrado = df_filtrado.reset_index(drop=True)
                            dict[key] = df_filtrado
                            print(f"DataFrame para '{key}' adicionado ao dicionário.")
                    except Exception as e:
                        print(f"Erro ao processar o arquivo '{arquivo}': {e}")
                
                # st.write("Resumo dos DataFrames carregados:")
                # for key, df_filtrado in dict.items():
                #     st.write(f"**{key}**: {df_filtrado.shape[0]} linhas, {df_filtrado.shape[1]} colunas")
                #     st.dataframe(df)
        else:
            st.warning('Nenhum valor encontrado.')
    else:
        st.error("Arquivo '01-Renda_Fixa.csv' não encontrado. Clique no botão 'Execução' para gerar os arquivos.")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        uploaded_csv = st.file_uploader("Faça o upload do arquivo CSV (Estoque):", type=["csv", "xlsx"])
    with col_up2:
        uploaded_excel = st.file_uploader("Faça o upload do arquivo Posição:", type=["xlsx", "csv", "xls"])


    if "cota" not in st.session_state or "investimento" not in st.session_state:
        col_in1, col_in2 = st.columns(2)
        with col_in1:
            cota = st.selectbox("Selecione o tipo de cota:", options=['SUB', 'SR'], index=0)
        with col_in2:
            investimento = st.number_input("Informe o valor do Investimento:",
                                            min_value=0.0,
                                            step=0.01,
                                            format="%.2f",
                                            # help="Digite o valor do Investimento."
                                            )

        if st.button("Executar Análise"):
            st.session_state["cota"] = cota
            st.session_state["investimento"] = investimento

    if "cota" in st.session_state and "investimento" in st.session_state:
        cota = st.session_state["cota"]
        investimento = st.session_state["investimento"]      

        #### ---- Obtenção do PL ---- ####
        # RENDA FIXA
        renda_fixa = dict['Renda_Fixa']
        
        # Verifica se o fundo que estamos utilizando para o cálculo envolve o FIDC CONSIG PUB
        if renda_fixa['CARTEIRA'].str.contains('FIDC CONSIG PUB').any():
            # Caso seja, realiza a renomeação da nomeclatura.
            renda_fixa['TITULO'] = renda_fixa['TITULO'].replace({'COTA MEZANINO': 'COTA SENIOR'})

        renda_fixa = renda_fixa[['TITULO', 'VALORLIQUIDO']]
        tabela_tr = str.maketrans({'.':'', ',':'.', '(':'', ')':''})
        renda_fixa.loc[:, 'VALORLIQUIDO'] = renda_fixa['VALORLIQUIDO'].astype(str).str.translate(tabela_tr).astype(float)
        renda_fixa_grouped = renda_fixa.groupby('TITULO', as_index=False)['VALORLIQUIDO'].sum()

        patrimonio = dict['Patrimonio-Totais']
        patrimonio['VALORPATRIMONIOLIQUIDO'] = patrimonio['VALORPATRIMONIOLIQUIDO'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
        junior = patrimonio['VALORPATRIMONIOLIQUIDO'][0]

        titulos_desejados = ['COTA MEZANINO', 'COTA SENIOR']
        renda_fixa_filtrada = renda_fixa_grouped[renda_fixa_grouped['TITULO'].isin(titulos_desejados)]
        soma_total = renda_fixa_filtrada['VALORLIQUIDO'].sum() + junior

        print('Valor do PL Final:', soma_total)

        df_pl = renda_fixa_filtrada.set_index('TITULO')
        df_pl.loc['COTA JUNIOR'] = [junior]
        df_pl.loc['PL'] = [soma_total]
        print('DataFrame do PL:')
        print(df_pl)
        
        #### ---- Obtenção do Saldo ---- ####

        if uploaded_csv is not None:
            # Utiliza da Função para ler o arquivo em CSV.
            df1 = ler(uploaded_csv)
            st.success("Arquivo carregado com sucesso.")
        else:
            st.warning("Erro ao carregar o arquivo.")
        
        combined_df = df1

        Documento = combined_df['DOC_SACADO']
        columns_to_convert = ['VALOR_NOMINAL', 'VALOR_PRESENTE', 'VALOR_AQUISICAO', 'VALOR_PDD']

        for col in columns_to_convert:
            if not is_float_dtype(combined_df[col]):
                combined_df[col] = (
                    combined_df[col]
                    .astype(float)
                )

        soma_valores = combined_df
        soma_valores['DATA_REFERENCIA'] = pd.to_datetime(soma_valores["DATA_REFERENCIA"],
                                                        # format='%d/%m/%Y',
                                                        # dayfirst=True
                                                        )
        soma_valores['DATA_VENCIMENTO_AJUSTADA_2'] = pd.to_datetime(soma_valores['DATA_VENCIMENTO_AJUSTADA_2'],
                                                                    # format='%d/%m/%Y',
                                                                    # dayfirst=True
                                                                    )
        soma_valores['ATRASO'] = (soma_valores['DATA_REFERENCIA'] - soma_valores['DATA_VENCIMENTO_AJUSTADA_2']).dt.days
        soma_valores['DATA_EMISSAO_2'] = pd.to_datetime(soma_valores['DATA_EMISSAO_2'], 
                                                        # format='%d/%m/%Y', 
                                                        # dayfirst=True
                                                        )
        soma_valores['Ativo problemático'] = soma_valores.apply(ativ_probl, axis=1, data=data_mes)
        # CONTRATO:

        # Soma para valores com ATRASO > 90
        over_90_saldo = soma_valores.loc[soma_valores['ATRASO'] > 90, 'VALOR_PRESENTE'].sum()
        # Soma para valores com ATRASO < 90
        below_90_saldo = soma_valores.loc[soma_valores['ATRASO'] <= 90, 'VALOR_PRESENTE'].sum()
        total_saldo = over_90_saldo + below_90_saldo
        # Soma para valores com ATRASO > 90
        over_90_pdd = soma_valores.loc[soma_valores['ATRASO'] > 90, 'VALOR_PDD'].sum()
        # Soma para valores com ATRASO < 90
        below_90_pdd = soma_valores.loc[soma_valores['ATRASO'] <= 90, 'VALOR_PDD'].sum()
        total_pdd = over_90_pdd + below_90_pdd

        contrato = total_saldo - total_pdd
        print('Contratos: ', contrato.round(2))

        # CAIXA
        patrimonio['SALDOCAIXAATUAL'] = patrimonio['SALDOCAIXAATUAL'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
        caixa = patrimonio['SALDOCAIXAATUAL'].iloc[0] 
        print('Caixa: ', caixa)

        # CONTAS Á RECEBER E Á PAGAR
        try:
            contas = dict['CPR-Lancamentos']
            contas = contas[['MODALIDADE', 'VALOR']]
            contas.loc[:, 'VALOR'] = contas.loc[:, 'VALOR'].apply(tratar_valores)
            contas_resultados = contas.groupby(by='MODALIDADE')['VALOR'].sum()
            pagar = contas_resultados.get("Pagar", 0)
            receber = contas_resultados.get("Receber", 0)
        except KeyError:
            pagar = 0
            receber = 0
        
        print("Pagar:", pagar)
        print("Receber:", receber)

        # FUNDO

        if "Fundos-Fundos" in dict:
            fundos = dict['Fundos-Fundos']

            # Verifica se as colunas necessárias existem.
            if all(col in fundos.columns for col in ['CODIGO', 'VALORLIQUIDO']):
                df_fundos = fundos[['CODIGO', 'VALORLIQUIDO']]
                df_fundos.loc[:, 'VALORLIQUIDO'] = df_fundos.loc[:, 'VALORLIQUIDO'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
                fundo = df_fundos['VALORLIQUIDO'].sum()
            else:
                print("As colunas 'CODIGO' e 'VALORLIQUIDO' não existem no arquivo 'Fundos-Fundos'.")
                fundo = 0
        else:
            print("Não houve arquivo 'Fundos-Fundos' no diretório.")
            fundo = 0
        print('Valor Total envolvendo os fundos: ', fundo)

        # RENDA FIXA
        renda_fixa_nao_desejada = renda_fixa_grouped[~renda_fixa_grouped['TITULO'].isin(titulos_desejados)]
        renda = renda_fixa_nao_desejada['VALORLIQUIDO'].sum()
        print('Outras Rendas: ', renda)

        # OUTROS ATIVOS
        if 'Outros_Ativos' in dict:
            ativos = dict['Outros_Ativos']
            ativos_não_desejados = ['A VENCER', 'VENCIDOS', 'PDD']
            ativos_desejados = ativos[~ativos['ATIVO'].isin(ativos_não_desejados)]
            ativos_desejados.loc[:, 'VALOR'] = ativos_desejados.loc[:, 'VALOR'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
            outros_ativos = ativos_desejados['VALOR'].sum()
        else:
            print('Não há outros ativos á serem adicionados')
            outros_ativos = 0
        print('Soma dos outros ativos: ', outros_ativos)

        data = []
        data.append(('Contrato', contrato.round(2)))
        data.append(('Caixa', caixa))
        data.append(('Pagar', pagar))
        data.append(('Receber', receber))
        # Verificar se 'Fundos-Fundos' está presente no dicionário
        if 'Fundos-Fundos' in dict:
            for _, row in df_fundos.iterrows():
                data.append((row['CODIGO'], row['VALORLIQUIDO']))
        else:
            data.append(('Fundo', fundo))
        # Verifica se há 'Renda Fixa' á ser add no dataframe
        if not renda_fixa_nao_desejada.empty:
            for _, row in renda_fixa_nao_desejada.iterrows():
                data.append((row['TITULO'], row['VALORLIQUIDO']))
        else:
            data.append(('Renda Fixa', renda))

        if 'Outros_Ativos' in dict:
            if not ativos_desejados.empty:
                for _, row in ativos_desejados.iterrows():
                    data.append((row['ATIVO'], row['VALOR']))
            else:
                data.append(('Outros Ativos', outros_ativos))
        else:
                data.append(('Outros Ativos', outros_ativos))

        df_saldo = pd.DataFrame(data, columns=['Index', 'SALDO'])
        df_saldo.set_index('Index', inplace=True)

        print("DataFrame Saldo:")
        print(df_saldo)

        # COMPARAÇÃO ENTRE OS VALORES:
        print('# ---- Comparação entre valores ---- #')
        print()
        if df_saldo['SALDO'].sum().round(2) == soma_total.round(2):
            print("VALORES IGUAIS.")

        else:
            print("OS VALORES SÃO DIFERENTES.")
            diferenca = df_saldo['SALDO'].sum().round(2) - soma_total.round(2)
            print(f"A DIFERENÇA É DE: {diferenca:,.2f}")

    #### ---- Obtenção do DataBase ---- ####
        investimento = investimento
        valor_investido = investimento
        valor_investido = float(valor_investido)
        valor_investido

        porcentagem = valor_investido / df_saldo['SALDO'].sum()
        # PORCENTAGEM
        porcentagem = porcentagem 
                
        print('Valor da Porcentagem: ', porcentagem)
        dicionario = {}
        dicionario['Porcentagem'] = porcentagem

        # VALOR PRESENTE PROPORCIONAL
        soma_valores['TIPO_RECEBIVEL'] = soma_valores['TIPO_RECEBIVEL'].apply(lambda x: x.encode('latin1').decode('utf-8'))
        soma_valores['VP_PROPORCIONAL'] = soma_valores['VALOR_PRESENTE'] * porcentagem
        # VALOR PDD PROPORCIONAL
        soma_valores['PDD_PROPORCIONAL'] = soma_valores['VALOR_PDD'] * porcentagem
        # ATIVOS PROBLEMÁTICOS
        soma_valores.loc[:, 'ATV_PROBL'] = soma_valores.apply(calcular_atv_probl, axis=1)
        # PF/PJ
        soma_valores.loc[:, 'PF/PJ'] = soma_valores['DOC_SACADO'].astype(str).apply(classificar_pessoa)
        # FPR
        soma_valores['FPR'] = np.where(
                soma_valores['TIPO_RECEBIVEL'] == "Precatórios", 12.5,
                np.where(
                    soma_valores['TIPO_RECEBIVEL'] == "Ação Judicial", 12.5,
                    np.where(
                        soma_valores['ATV_PROBL'] != "N", soma_valores['ATV_PROBL'],
                        np.where(
                            soma_valores['VP_PROPORCIONAL'].astype(float) > 100000,
                            np.where(soma_valores['PF/PJ'] == "PJ", 0.85, 1.00),
                            0.75
                        )
                    )
                )
            )
        # RWA
        soma_valores['RWACPAD']  = (soma_valores['VP_PROPORCIONAL'] - soma_valores['PDD_PROPORCIONAL']) * soma_valores['FPR']

    #### ---- Cálculo do FPR e RWA ---- ####
        # FPR
        dicionário = {
            'Contrato': soma_valores['RWACPAD'].sum() / (soma_valores['VP_PROPORCIONAL'] - soma_valores['PDD_PROPORCIONAL']).sum(),
            'Caixa': 0.2,
            'Pagar': 0.0,
            'Receber': 1.0,
        }
        df_saldo['FPR'] = df_saldo.index.map(dicionário)

        # RWA
        df_saldo['RWA'] = df_saldo['SALDO'] * df_saldo['FPR']

        # DATAFRAME
        print("DataFrame Saldo Atualizado:")
        print(df_saldo)

    #### ---- Obtenção do FPR Final Apurado ---- ####
        def calcular_fpr_final(df_saldo, df_pl, soma_total, over_90_saldo, over_90_pdd, pagar, receber, cota):
            outros = 0
            if 'Ajuste' in df_saldo.index and pd.notna(df_saldo.loc['Ajuste', 'SALDO'] and df_saldo.loc['Ajuste', 'SALDO'] != 0):
                outros = float(df_saldo.loc['Ajuste', 'SALDO'])
            cota = cota
            sub = None
            sr = None

            # Definiação dos valores para sub e sr:

            if 'CARTEIRA' in patrimonio.columns:  # Verifica se a coluna existe para evitar erros.
                if patrimonio['CARTEIRA'].str.contains('FIDC REAG H Y').any():
                    if cota == 'SUB':
                        sub = df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total
                        sr = (df_pl.loc['COTA MEZANINO', 'VALORLIQUIDO'] + df_pl.loc['COTA SENIOR', 'VALORLIQUIDO']) / soma_total
                    else:
                        sub = (df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] + df_pl.loc['COTA MEZANINO', 'VALORLIQUIDO']) / soma_total
                        sr = df_pl.loc['COTA SENIOR', 'VALORLIQUIDO'] / soma_total
                else:
                    if 'COTA MEZANINO' in df_pl.index:
                        if 'COTA SENIOR' in df_pl.index:
                            if cota == 'SUB':
                                sub = (df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] + df_pl.loc['COTA MEZANINO', 'VALORLIQUIDO']) / soma_total
                                sr = (df_pl.loc['COTA MEZANINO', 'VALORLIQUIDO'] + df_pl.loc['COTA SENIOR', 'VALORLIQUIDO']) / soma_total
                            else:
                                sub = df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total
                                sr = df_pl.loc['COTA SENIOR', 'VALORLIQUIDO'] / soma_total
                        else:
                            if cota == 'SUB':
                                # sub = (df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] + df_pl.loc['COTA MEZANINO', 'VALORLIQUIDO']) / soma_total
                                sub = (df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO']) / soma_total
                                sr = df_pl.loc['COTA MEZANINO', 'VALORLIQUIDO'] / soma_total
                            else:
                                sub = df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total
                                sr = df_pl.loc['COTA MEZANINO', 'VALORLIQUIDO'] / soma_total
                    else:
                        if 'COTA SENIOR' in df_pl.index:
                            if cota == 'SUB':
                                sub = df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total
                                sr = df_pl.loc['COTA SENIOR', 'VALORLIQUIDO'] / soma_total
                            else:
                                sub = df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total
                                sr = df_pl.loc['COTA SENIOR', 'VALORLIQUIDO'] / soma_total
                        else:
                            if cota == 'SUB':
                                sub = df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total
                                sr = (df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total) - 1
                            else:
                                sub = df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total
                                sr = (df_pl.loc['COTA JUNIOR', 'VALORLIQUIDO'] / soma_total) - 1
            else:
                print("A coluna 'CARTEIRA' não existe no DataFrame.")

            if 'sub' in locals():
                print('SUB:', sub.round(2))
                dicionario['SUB'] = sub
            if 'sr' in locals():
                print('SR:', sr.round(2))
                dicionario['SR'] = sr

            # Cálculo da Razão da Inadimplência
            razao_inadimplencia = ((over_90_saldo - over_90_pdd) + outros) / df_saldo['SALDO'].sum()
            dicionario['Inadimplência'] = razao_inadimplencia

            # Valores para o Encaixe e Desencaixe
            if cota == 'SR':
                ponto_encaixe = sub
                ponto_desencaixe = 1.0
            elif cota == 'SUB':
                ponto_encaixe = 0.0
                ponto_desencaixe = sub
            
            dicionario['Encaixe'] = ponto_encaixe
            dicionario['Desencaixe'] = ponto_desencaixe

            print('Ponto de Encaixe: ', ponto_encaixe, 'Ponto de Desencaixe: ', ponto_desencaixe, 'Razão da Inadimplência: ', razao_inadimplencia.round(4))
            print('SUB:', sub, 'SR:', sr)

            # Variáveis para o Cálculo do FPR
            A = ponto_encaixe
            dicionario['A'] =  A
            D = ponto_desencaixe
            dicionario['D'] = D
            W = razao_inadimplencia
            dicionario['W'] = W
            RWAhip = df_saldo['RWA'].sum()
            dicionario['RWAhip'] = RWAhip
            V = (df_saldo['SALDO'].sum() - pagar).round(0)
            dicionario['V'] = V
            F = 0.08

            # Cálculo de Ksa
            Ksa = (RWAhip * F) / V
            dicionario['Ksa'] = Ksa
            print(f"Ksa: {Ksa:.2%}")

            # Cálculo do Ka
            Ka = ((1 - W) * Ksa + (W * 0.5))
            print(f"Ka: {Ka:.2%}")
            dicionario['Ka'] = Ka

            parte_1 = np.exp(((- 1 / Ka) * (D - Ka)))
            print(f"{parte_1:.2%}")
            parte_2 = np.exp((-1 / Ka) * np.maximum(A - Ka, 0))
            print(f"{parte_2:.2%}")
            parte_3 = - (1 / Ka) * (D - Ka - np.maximum(A - Ka, 0))
            print(f"{parte_3:.2%}")
            Kssfa = (parte_1 - parte_2) / parte_3
            print(f"{Kssfa:.2%}")
            dicionario['Kssfa'] = Kssfa
            parte_4 = ((Ka - A) / (D - A)) * (1/F)
            print(f"{parte_4:.2%}")
            parte_5 = ((D - Ka) / (D - A)) * (1/F) * Kssfa
            print(f"{parte_5:.2%}")
            FPR = parte_4 + parte_5
            print(f"{FPR:.2%}")
            dicionario['FPR'] = FPR

            Ka = Ka
            I = (1/F)
            dicionario['I'] = I
            print(f"{I:.2%}")
            II = (1/F) * Kssfa
            dicionario['II'] = II
            print(f"{II:.2%}")
            III = FPR
            dicionario['III'] = III
            print(f"{III:.2%}")

            A = ponto_encaixe
            D = ponto_desencaixe 
            W = razao_inadimplencia
            RWAhip = df_saldo['RWA'].sum()
            V = (df_saldo['SALDO'].sum() - pagar).round(0)

            print(f'Verifica se {D} é menor que {float(Ka)}')
            if D <= float(Ka):
                D6 = True
                print(f'Caso D6: o valor do D é igual á {D} e o valor do Ka é {Ka}, nesse caso {D} é menor que {float(Ka)}, D6 passar á ser {D6}')
            else:
                D6 = False
                print(f'Caso do {D} ser maior que {float(Ka)}, D6 é, portanto, {D6}')

            print()
            print(f'Verifica se {float(A)} é maior que {float(Ka)}')
            if float(A) >= float(Ka):
                D7 = True
                print(f'Caso D7: o valor do A é igual á {float(A)} e o valor do Ka é {Ka}, nesse caso {float(A)} é maior que {float(Ka)}, D7 passa a ser {D7}')
            else:
                D7 = False
                print(f'Caso do {float(Ka)} ser menor que {float(Ka)}, D7 é, portanto, {D7}')

            F7 = 0.25

            print()

            if D6 == True:
                F6 = float(I)
                print(f'Se {D6} == True, então F6 é igual á {float(I)}')
            else:
                if D7 == True:
                    F6 = float(II)
                    print(f'Se {D7} == True, então F6 é igual á {float(II)}')
                else:
                    F6 = float(III)
                    print(f'Se {D7} == False, então F6 é igual á {float(III)}')

            print(F6, F7)
            print('Pegar o valor máximo:')
            print()
            FPR_apurado = float(np.maximum(F6, F7))
            print()
            print()
            print('# ---- Valor Final para o FPR ---- #')
            print(f"{FPR_apurado:.0%}")
            return FPR_apurado
        
    #### ---- Interface do Streamlit ---- #####
        st.title("Editar DataFrame Dinamicamente no Streamlit")

        # Selecionar as colunas a serem exibidas.
        columns_to_display = ['Index', 'SALDO', 'FPR']
        editable_columns = df_saldo.reset_index()[columns_to_display]

        # Adicionar várias linhas em branco (exemplo: 5 linhas)
        num_blank_rows = 10
        blank_rows = pd.DataFrame([{'Index': '', 'SALDO': 0, 'FPR': None}] * num_blank_rows)
        editable_columns = pd.concat([editable_columns, blank_rows], ignore_index=True)

        # Exibir DataFrame editável
        st.markdown(
        """
        <hr style="border: 1px solid white; margin: 20px 0;">
        """, unsafe_allow_html=True)

        st.write("## 1) Edição dos Valores da Coluna FPR:")
        edited_subset = st.data_editor(editable_columns, use_container_width=True)

        # Processar edições realizadas pelo usuário:
        updated_df = edited_subset[edited_subset['Index'] != ''].set_index('Index')
        updated_df['SALDO'] = updated_df['SALDO'].astype(float)
        updated_df['FPR'] = updated_df['FPR'].astype(float)

        # Recalcular a coluna RWA
        updated_df['RWA'] = updated_df['SALDO'] * updated_df['FPR']

        # Atualizar o DataFrame original
        df_saldo = updated_df
        df_saldo = df_saldo[~df_saldo.index.isin(["Renda Fixa", "Outros Ativos"])]
        st.markdown(
        """
        <hr style="border: 1px solid white; margin: 20px 0;">
        """, unsafe_allow_html=True)
        st.write("## 2) Comparação entre DataFrames:")
        fpr_final = calcular_fpr_final(
            df_saldo, df_pl, soma_total, over_90_saldo, over_90_pdd, pagar, receber, cota
        )
        col1, col2 = st.columns([1, 2])

        with col1:
            if not df_pl.empty:
                st.write("#### DataFrame Patrimônio Líquido")
                st.dataframe(df_pl)
                st.write("#### DataFrame Saldo")
                st.dataframe(df_saldo)
                st.write(f'Total da coluna Saldo: {locale.format_string('%.2f', df_saldo['SALDO'].sum(), grouping=True)}')
            st.write(f'Toda da coluna RWA: {locale.format_string('%.2f', df_saldo['RWA'].sum(), grouping=True)}')
            expo = df_saldo['SALDO'].sum() - pagar
            st.write(f'EXPO: {locale.format_string('%.2f', expo, grouping=True)}')
            st.markdown(
            f"""
            <div style="text-align: center; font-size: 24px; font-weight: bold; color: white;">
            Comparação entre os valores:
            </div>
            """,
            unsafe_allow_html=True,)
            if df_saldo['SALDO'].sum().round(2) == soma_total.round(2):
                print("VALORES IGUAIS.")
                st.markdown(
                """
                <div style="text-align: center; font-size: 16px; font-weight: bold; color: green;">
                OS VALORES ENTRE OS DATAFRAMES SÃO IGUAIS.
                </div>
                """,
                unsafe_allow_html=True,)
            else:
                print("OS VALORES SÃO DIFERENTES.")
                diferenca = df_saldo['SALDO'].sum().round(2) - soma_total.round(2)
                print(f"A DIFERENÇA É DE: {diferenca:,.2f}")
                st.markdown(
                f"""
                <div style="text-align: center; font-size: 16px; font-weight: bold; color: red;">
                OS VALORES ENTRE OS DATAFRAMES SÃO DIFERENTES, DIFERENÇA DE: {diferenca:.2f}
                </div>
                """,
                unsafe_allow_html=True,)

        with col2:
            st.write('### Database:')
            st.dataframe(soma_valores)

        st.write('#### Saldo %')
        saldo_porcentagem = df_saldo['SALDO'] * porcentagem
        st.dataframe(saldo_porcentagem.to_frame().T)
        st.write('### Informações:')
        df_dicionário = pd.DataFrame(dicionario.items(), columns=['Item', 'Valor dos Itens'])
        df_dicionário.set_index('Item', inplace=True)
        df_dicionário['Valor dos Itens'] = df_dicionário['Valor dos Itens'].round(3)
        st.dataframe(df_dicionário.T)
        st.markdown(
        """
        <hr style="border: 1px solid white; margin: 20px 0;">
        """,
        unsafe_allow_html=True)
        st.markdown(
                f"""
                <div style="text-align: center; font-size: 30px; font-weight: bold;">
                Valor Final do FPR Apurado: {fpr_final:.2%}
                </div>
                """,
                unsafe_allow_html=True,)
        st.markdown(
        """
        <hr style="border: 1px solid white; margin: 20px 0;">
        """,
        unsafe_allow_html=True)

        st.write("## 3) Arquivo Posição:")
        st.write('### Arquivo Modelo:')
        st.write("*Utilize o DataFrame abaixo como modelo; em tese, é o arquivo enviado do mês anterior.")
        if uploaded_excel is not None:
            df_excel = pd.read_excel(uploaded_excel)
            st.dataframe(df_excel)
        else:
            st.warning("Nenhum arquivo Excel foi carregado.")
        
        # st.write('Arquivo CSV á ser exportado:')
        st.write('### DataFrame Editável:')
        segunda_coluna = df_excel.iloc[:, 1]

        # Alteração dos valores do DataFrame:
        posicao = pd.DataFrame()
        posicao['X'] = soma_valores['VP_PROPORCIONAL']
        posicao['Sistema Origem'] = opção_selecionada
        posicao['Localização'] = 'BR'
        posicao['Tipo de carteira'] = 0
        posicao['Produto'] = 'OPERACAO_CREDITO_FUNDO'
        posicao['Data de contratação'] = soma_valores['DATA_EMISSAO_2'].dt.strftime('%d/%m/%Y').astype(str)
        posicao['Data Vencimento'] = soma_valores['DATA_VENCIMENTO_AJUSTADA_2'].dt.strftime('%d/%m/%Y').astype(str)
        posicao['Moeda Operação'] = 'BRL'
        posicao['Forma de liquidacao'] = np.nan
        posicao['Indexador Ativo'] = np.nan
        posicao['Indexador Passivo'] = np.nan
        posicao['Valor atual'] = soma_valores['VP_PROPORCIONAL']
        posicao['Valor original'] = np.nan
        posicao['Sistema Registro'] = np.nan
        posicao['Tipo Controle'] = np.nan
        posicao['Tipo Pessoa'] = np.nan
        posicao['Grupo Econômico/Matriz/Conectada'] = np.nan
        posicao['Faturamentos'] = np.nan
        posicao['Provisoes'] = pd.to_numeric(soma_valores['PDD_PROPORCIONAL'])
        posicao['Notional'] = np.nan
        posicao['Contraparte'] = df_excel['Unnamed: 20'].mode()[0]
        posicao['Contraparte'] = posicao['Contraparte'].fillna('').astype(str).str.replace(',', '', regex=False) ####
        posicao['Compromissada Over ou operação marcada à mercado'] = np.nan
        posicao['Tipo acordo compensação (CEM)/ Indicador Margem (SA-CCR)'] = np.nan
        posicao['Ativo objeto'] = soma_valores['TIPO_RECEBIVEL']
        posicao['Emitente'] = soma_valores['DOC_SACADO']
        posicao['Emitente'] = posicao['Emitente'].str.replace(r'[./-]', '', regex=True)
        posicao['Indicador MTM Ativo'] = np.nan
        posicao['Data Vencimento Ativo'] = soma_valores['DATA_VENCIMENTO_AJUSTADA_2'].dt.strftime('%d/%m/%Y').astype(str)
        posicao['Valor Ativo Associado'] = soma_valores['VP_PROPORCIONAL']
        posicao['Moeda Ativo Objeto'] = np.nan
        posicao['Nome da Contraparte'] = posicao['Sistema Origem']
        posicao['Nome do Emissor'] = np.nan
        posicao['COSIF'] = np.nan
        posicao['Descrição  ativo objeto'] = np.nan
        posicao['Razão para conectar contraparte'] = np.nan
        posicao['P'] = np.nan
        posicao['Data do próximo fluxo'] = np.nan
        posicao['K'] = np.nan
        posicao['Lambda'] = np.nan
        posicao['Data início ou exercício do ativo'] = np.nan
        posicao['Data da Posição'] = np.nan
        posicao['Ativo problemático'] = soma_valores['Ativo problemático']
        posicao['Característica especial'] = np.nan
        posicao['Campo reservado'] = np.nan
        posicao['Valor atual cenário 1'] = np.nan
        posicao['Valor notional cenário 1'] = np.nan
        posicao['Valor atual cenário 2'] = np.nan
        posicao['Valor notional cenário 2'] = np.nan
        posicao['Valor atual cenário 3'] = np.nan
        posicao['Valor notional cenário 3'] = np.nan
        posicao['Valor atual cenário 4'] = np.nan
        posicao['Valor notional cenário 4'] = np.nan
        posicao['Valor atual cenário 5'] = np.nan
        posicao['Valor notional cenário 5'] = np.nan
        posicao['Valor atual cenário 6'] = np.nan
        posicao['Valor notional cenário 6'] = np.nan
        posicao = posicao.reset_index()
        posicao = posicao.rename({'index': 'Num operacao'}, axis=1)
        posicao['Num operacao'] = posicao['Num operacao'].apply(lambda x: x + 1) 
        posicao = posicao.drop('X', axis=1)

        moda_colunas = posicao.mode().iloc[0]

        for index, valor in df_saldo.iterrows():
            if index == "Contrato":
                continue
            nova_linha = {
                "Produto": index,                                
                "Valor atual": valor['SALDO'] * porcentagem,    
            }

            for coluna in posicao.columns:
                if coluna not in nova_linha: 
                    nova_linha[coluna] = moda_colunas[coluna]
                    
            nova_linha_df = pd.DataFrame([nova_linha], columns=posicao.columns)
            
            posicao = pd.concat([posicao, nova_linha_df], ignore_index=True)      
        
        ultimo_valor = posicao['Num operacao'].max()
        contador = ultimo_valor # + 1
        indices = posicao.index[posicao['Num operacao'] == 1].tolist()
        for i in indices:
            posicao.at[i, 'Num operacao'] = contador
            contador += 1
        posicao.loc[0, 'Num operacao'] = 1

        posicao['Produto'] = posicao['Produto'].replace("Caixa", "CAIXA_FIDC")
        posicao['Produto'] = posicao['Produto'].replace("Pagar", "VALOR_A_PAGAR_FIDC")
        posicao['Produto'] = posicao['Produto'].replace("Receber", "VALOR_A_RECEBER_FIDC")

        for i in posicao['Produto'].unique():
            if i != "OPERACAO_CREDITO_FUNDO":
                posicao.loc[posicao['Produto'] == i, 'Data de contratação'] = data_mes
                proximo_dia_util = (pd.to_datetime(data_mes) + BDay(1)).strftime('%d/%m/%Y')
                posicao.loc[posicao['Produto'] == i, 'Data Vencimento'] = proximo_dia_util
                posicao.loc[posicao['Produto'] == i, 'Ativo problemático'] = 'N'

        edited_posicao = st.data_editor(posicao, num_rows="dynamic", use_container_width=True, height=600)

        st.write("### Arquivo Final:")
        st.write('* Faço o download deste arquivo no formato CSV.')
        lista = [
            "Data do Ciclo",
            data_mes,
            "Tipo de Movimento",
            "I",
            "Empresa",
            df_excel.columns[5],
            "Alias",
            df_excel.columns[7],
            "Unnamed: 8","Unnamed: 9","Unnamed: 10",
            posicao['Valor atual'].sum(),
            posicao['Valor original'].sum(),
            "Unnamed: 13","Unnamed: 14","Unnamed: 15","Unnamed: 16","Unnamed: 17",
            posicao['Provisoes'].sum(),
            "Unnamed: 19","Unnamed: 20","Unnamed: 21","Unnamed: 22","Unnamed: 23","Unnamed: 24","Unnamed: 25","Unnamed: 26","Unnamed: 27","Unnamed: 28","Unnamed: 29",
            "Unnamed: 30","Unnamed: 31","Unnamed: 32","Unnamed: 33","Unnamed: 34","Unnamed: 35","Unnamed: 36","Unnamed: 37","Unnamed: 38","Unnamed: 39","Unnamed: 40",
            "Unnamed: 41","Unnamed: 42","Unnamed: 43","Unnamed: 44","Unnamed: 45","Unnamed: 46","Unnamed: 47","Unnamed: 48","Unnamed: 49","Unnamed: 50","Unnamed: 51",
            "Unnamed: 52","Unnamed: 53", "Unnamed: 54", "Unnamed: 55"
        ]

        # Alteração das colunas:
        num_colunas = len(edited_posicao.columns)
        if len(lista) < num_colunas:
            lista.extend([""] * (num_colunas - len(lista)))
        elif len(lista) > num_colunas:
            lista = lista[:num_colunas]

        # Adicionar a primeira linha com os nomes das colunas atuais
        edited_posicao.loc[-1] = edited_posicao.columns 
        edited_posicao.index = edited_posicao.index + 1 
        edited_posicao = edited_posicao.sort_index()  # Reorganiza o índice

        # Substituir os nomes das colunas pelo conteúdo da lista ajustada
        edited_posicao.columns = lista

        st.dataframe(edited_posicao)

        uploaded_csv = st.file_uploader("### Faça o upload do arquivo POSIÇÃO em csv", type=["csv"])

        if uploaded_csv is not None:
            df = pd.read_csv(uploaded_csv)
            df.columns = ["" if "Unnamed" in col else col for col in df.columns]

            # Converter o DataFrame para Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")
            excel_data = output.getvalue()

             # Salvar o arquivo temporariamente para modificações com a biblioteca openpyxl
            temp_excel_path = "temp_arquivo_convertido.xlsx"
            with open(temp_excel_path, "wb") as temp_file:
                temp_file.write(excel_data)

            # Modificar o arquivo gerado
            wb = load_workbook(temp_excel_path)
            ws = wb.active

            # Exclusão da primeira coluna
            ws.delete_cols(1)

            # Formatação da coluna "Data do Ciclo":
            for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
                for cell in row:
                    if cell.value is not None:
                        cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

            # Formatação da coluna "Alias":
            for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
                for cell in row:
                    if cell.value is not None:
                        cell.number_format = numbers.FORMAT_DATE_DDMMYY

            
            coluna_L = 12  # Coluna L é a 12ª coluna.
            soma_coluna = 0

            for row in ws.iter_rows(min_row=2, min_col=coluna_L, max_col=coluna_L):  # Iteração a partir da segunda linha
                for cell in row:
                    if cell.value is not None:  
                        try:
                            cell.value = float(cell.value) 
                            soma_coluna += cell.value  
                            cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1  # Formatar como numérico
                        except ValueError:
                            pass  # Ignorar erros de conversão, caso a célula tenha um valor inválido.

            # Renomeação da coluna L com a soma obtida:
            ws.cell(row=1, column=coluna_L).value = soma_coluna
            ws.cell(row=1, column=coluna_L).number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            
            # Tratamento de Dados da coluna D.
            coluna_D = "D"  
            coluna_D_index = ws[coluna_D + "1"].column  

            # Iterar pelas células a partir da segunda linha na coluna D:
            for row in ws.iter_rows(min_row=2, min_col=coluna_D_index, max_col=coluna_D_index):
                for cell in row:
                    if cell.value is not None: 
                        try:
                            valor = float(cell.value)
                            if valor == 0.0:  # Se for 0.0 (Geral), substitui por 0.
                                cell.value = 0
                            else:
                                cell.value = valor  
                            cell.number_format = numbers.FORMAT_NUMBER  
                        except ValueError:
                            pass
            
            # Tratamento de dados da coluna S:
            coluna_S = "S"
            coluna_S_index = ws[coluna_S + "1"].column  

            # Converte os valores da coluna "S" para números e formata com duas casas decimais
            soma_coluna_S = 0
            for row in ws.iter_rows(min_row=2, min_col=coluna_S_index, max_col=coluna_S_index):
                for cell in row:
                    if cell.value not in (None, "0", "0.0", 0.0):  # Ignora valores nulos ou zero
                        try:
                            cell.value = round(float(cell.value), 2)
                            soma_coluna_S += cell.value  
                            cell.number_format = "#,##0.00"  # Formato de número com duas casas decimais
                        except ValueError:
                            pass 

            # Atualiza o cabeçalho da coluna "S" com a soma
            ws[coluna_S + "1"].value = round(soma_coluna_S, 2)
            ws[coluna_S + "1"].number_format = "#,##0.00"

            # Acessa a coluna "AB"
            coluna_AB = "AB"
            coluna_AB_index = ws[coluna_AB + "1"].column  # Obtém o índice da coluna "AB"

            # Converte os valores da coluna "AB" para números e formata com duas casas decimais
            for row in ws.iter_rows(min_row=2, min_col=coluna_AB_index, max_col=coluna_AB_index):
                for cell in row:
                    if cell.value not in (None, "0", "0.0", 0.0):  
                        try:
                            cell.value = round(float(cell.value), 2)
                            cell.number_format = "#,##0.00" 
                        except ValueError:
                            pass 
            
            # Transformar a coluna AA em formato de data (a partir da segunda linha):
            for cell in ws['AA'][1:]:
                if cell.value:
                    cell.number_format = 'DD/MM/YYYY'

            ws['M1'] = '=L1-S1'

            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = max_length + 2  # Ajuste adicional
                ws.column_dimensions[col_letter].width = adjusted_width
            
            # Tratamento das bordas das coluna:
            borda_vazia = Border(left=Side(border_style=None),
                     right=Side(border_style=None),
                     top=Side(border_style=None),
                     bottom=Side(border_style=None))

            # Iterar pelas células na primeira linha
            for cell in ws[1]:
                if cell.value is None:  
                    cell.border = borda_vazia  # Remover as bordas

            # Salvar as alterações no arquivo Excel
            wb.save(temp_excel_path)
            wb.close()

            # Reabrir o arquivo modificado e preparar para download
            with open(temp_excel_path, "rb") as modified_file:
                modified_excel_data = modified_file.read()

            st.download_button(
                label="Baixar em Excel",
                data=modified_excel_data,
                file_name=f"arquivo_convertido_{str(opção_selecionada)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ----- Execução ----- #
if st.button("Execução"):
    if not validar_data(data=data):
        st.error("Insira uma data válida no formato dd/mm/YYYY")
    else:
        try:
            extração_amplis(data)
            st.success("Extração concluída!")
        except TimeoutError as e:
            st.error('Problemas com o AMPLIS, tente novamente após alguns instantes.')
        except Exception as e:
            st.error("Ocorreu um erro inesperado! Tente novamente após alguns instantes.")
    
    # if not validar_data(data=data):
    #     st.error("Insira uma data válida no formato dd/mm/YYYY")
    # else:
    #     extração_amplis(data=data)
main()




