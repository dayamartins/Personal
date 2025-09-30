import time
import os
import pandas as pd
import shutil
import schedule
import win32com.client
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from plyer import notification

print("Iniciando atualização do relatório FTTA")

notification.notify(
    title='Iniciando Automação',
    message='A atualização do relatório FTTA está começando.',
    app_name='Automação',
    timeout=5
)
print("Primeira notificação enviada.")

try:
    # Configura o do Selenium com preferências de download
    options = Options()
    options.add_argument("--start-maximized")  # Abre a Janela Maximizada
    # options.add_argument("--start-minimized") # Abre a Janela Minimizada
    # options.add_argument("--headless=new")  # Abre a janela no modo invisível
    prefs = {
        "download.default_directory": os.path.join(os.path.expanduser("~"), "Downloads"),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    USERNAME = "email"
    PASSWORD = "password"

    # 1. Acessa o site
    driver.get("URL")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "input")))

    # 2. Faz Login
    try:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.NAME, "loginfmt")))
        email_input = driver.find_element(By.NAME, "loginfmt")
        email_input.click()
        email_input.clear()
        email_input.send_keys(USERNAME)
        email_input.send_keys(Keys.RETURN)

        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.NAME, "passwd")))
        password_input = driver.find_element(By.NAME, "passwd")
        password_input.click()
        password_input.clear()
        password_input.send_keys(PASSWORD)
        password_input.send_keys(Keys.RETURN)

        # "Ficar conectado?" → clica em "Não"
        try:
            no_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "idBtn_Back"))
            )
            no_button.click()
            print("Clicked 'Não' on stay signed in prompt.")
        except:
            try:
                no_button_text = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@value='Não']"))
                )
                no_button_text.click()
                print("Clicked 'Não' on stay signed in prompt (via value).")
            except:
                print("Tela 'Ficar conectado?' não apareceu ou botão não encontrado.")

        print("✅ Login completo!")
        time.sleep(10)

    except Exception as e:
        print(f"❌ Erro no login: {e}")
        notification.notify(
            title='Erro no Login',
            message='Erro no login. Operação interrompida.',
            app_name='Automação',
            timeout=5
        )
        driver.quit()
        exit()

    # 3. Procura o botão "Mais opções"

    try:
        botao_menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@title='Mais opções']"))
        )
        botao_menu.click()
        print("Botão 'Mais opções' clicado diretamente.")
        botao_encontrado = True
    except:
        print("Botão não encontrado diretamente, tentando nos iframes...")

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    botao_encontrado = False

    for i, iframe in enumerate(iframes):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(iframe)
            botao_menu = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@title='Mais opções']"))
            )
            botao_menu.click()
            print(f"Botão 'Mais opções' clicado no iframe {i}")
            botao_encontrado = True
            break
        except Exception as e:
            print(f"Iframe {i} não tem o botão: {e}")
            continue

    driver.switch_to.default_content()  # Volta para o conteúdo principal

    if not botao_encontrado:
        print("Erro: Nenhum botão correto foi encontrado!")

    # 4. Clica em "Exportar"
    try:
        print("Procurando botão 'Exportar'...")
        botao_exportar = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Exportar')]"))
        )
        botao_exportar.click()
        print("Botão 'Exportar' clicado.")
    except Exception as e:
        print(f"Erro ao encontrar o botão 'Exportar': {e}")

    # 5. Aguarda o download
    time.sleep(15)

    # 6. Verifica se de fato ocorreu o download
    nome_arquivo = "FTTA_BASE.csv"
    origem = os.path.join(os.path.expanduser("~"), "Downloads", nome_arquivo)
    if os.path.exists(origem):
        print("Relatório atualizado com sucesso!")
    else:
        print("Erro: Arquivo não encontrado. O download pode não ter sido concluído.")
        notification.notify(
            title='Erro no Download',
            message='Erro no download. Operação interrompida.',
            app_name='Automação',
            timeout=5
        )
        driver.quit()
        exit()

except Exception as e:
    print(f"❌ Erro inesperado: {e}")

finally:
    driver.quit()
    print("Navegador fechado.")

# MOVENDO O ARQUIVO
destino = os.path.join(os.path.expanduser("~"), "User", "Folder1", "Folder2",
                       nome_arquivo)
if os.path.exists(origem):
    if os.path.exists(destino):
        os.remove(destino)
        print(f"Arquivo existente removido: {destino}")
    shutil.move(origem, destino)
    print(f"Arquivo movido para {destino}")
else:
    print("Erro: Arquivo CSV não encontrado para mover.")

# ATUALIZAÇÃO DO EXCEL
file_path = r"C:\Users\User\Folder1\Folder2\FTTA_DATA_AUTOMATION.xlsx"
print("Iniciando a atualização do Excel...")
if os.path.exists(file_path):
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.Visible = False

    workbook = excel.Workbooks.Open(file_path)
    try:
        workbook.RefreshAll()
        time.sleep(5)
        workbook.Save()
        print(f"Arquivo {file_path} atualizado com sucesso!")
    except Exception as e:
        print(f"Erro ao atualizar o arquivo: {e}")
    finally:
        workbook.Close(SaveChanges=True)
        excel.DisplayAlerts = True
        print(f"Arquivo {file_path} fechado!")
else:
    print(f"Erro: O arquivo {file_path} não foi encontrado!")

# NOTIFICAÇÃO FINAL
print("Tentando enviar a segunda notificação...")
notification.notify(
    title='Atualização Concluída',
    message='Relatório FTTA foi atualizado com sucesso!',
    app_name='Automação',
    timeout=5
)
print("Segunda notificação enviada.")
