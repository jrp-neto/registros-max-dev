import functions
import logs
import pandas as pd
import traceback
import warnings
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

# Configurações gerais
options = Options()
options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-usb"])
warnings.simplefilter("ignore", UserWarning)

def start(user: str, password: str, spreadsheet: str, run_mode: bool):
    # Verificar se o processo será executado em segundo plano
    options.add_argument("--headless") if run_mode else options.add_argument("--force-device-scale-factor=0.9")
    # Inicialização de variáveis
    url = "https://intranet.sereduc.com/CRA/ConsultarAluno.aspx"
    try:
        df_spreadsheet = pd.read_excel(spreadsheet, sheet_name=1, dtype=object)
    except PermissionError:
        logs.logging.info("Feche a planilha antes de inciar os registros!")
        return "Feche a planilha antes de inciar os registros!", "red", 10000
    except FileNotFoundError:
        logs.logging.info("Arquivo Excel não encontrado!")
        return "Arquivo Excel não encontrado!", "red", 10000
    register_counter = 0
    error_counter = 0
    e_mails = []
    try:
        # Abrir navegador na URL desejada e faz login
        chrome = webdriver.Chrome(options=options)
        chrome.get(url)
        login_error = functions.login(user, password, chrome)
        if login_error:
            return "Usuário ou senha estão incorretos!", "red", 10000
        for _, row in df_spreadsheet.iterrows():
            try:
                # Verificar se todas as informações necessárias estão OK
                if pd.isna(row["E-mail"]) or pd.isna(row["Matrícula"]) or pd.isna(row["Registro CRA"]):
                    if pd.isna(row["E-mail"]):
                        continue
                    logs.logging.info(f'As colunas "Matrícula" e "Registro CRA" devem estar preenchidas. E-mail: {row["E-mail"]}')
                    e_mails.append(row["E-mail"])
                    error_counter += 1
                    continue
                # Buscar discente
                functions.search(row["Matrícula"], chrome)
                # Registrar ocorrência
                functions.register(row["Registro CRA"], chrome)
                # Verificar se o registro foi salvo
                functions.verify(row["Registro CRA"], row["Matrícula"], chrome)
            except:
                error_message = traceback.format_exc()
                logs.logging.error(f"Ocorreu um problema ao registrar o discente de matrícula: {str(row["Matrícula"]).zfill(8)}. Erro: {error_message}.")
                logs.logging.error("Passando para o próximo discente...")
                if len(chrome.window_handles) > 1:
                    chrome.close()
                    chrome.switch_to.window(chrome.window_handles[0])
                error_counter += 1
                e_mails.append(row["E-mail"])
                chrome.get(url)
            else:
                logs.logging.info(f"Registro na matrícula {str(row["Matrícula"]).zfill(8)} salvo com sucesso!")
                register_counter += 1
                chrome.get(url)
        # Finalizar registros
        logs.logging.info("Registros finalizados!")
        logs.logging.info(f"Registros realizados com sucesso: {register_counter}!")
        logs.logging.info(f"Registros não realizados: {error_counter}")
        logs.logging.info(f"E-mail dos alunos que não foram registrados: ")
        for e_mail in e_mails:
           logs.logging.info(f"{e_mail}")
        return "Registros finalizados!", "green", 60000
    finally:
        chrome.quit()