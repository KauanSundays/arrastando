import time
import os
import json
import re
import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import config

usuario = config.USUARIO
senha = config.SENHA

# --- Criar pasta principal se não existir ---
pasta_principal = os.path.join(os.getcwd(), "processos_pdfs")
os.makedirs(pasta_principal, exist_ok=True)

options = Options()
options.add_experimental_option("prefs", {
    "download.default_directory": pasta_principal,  
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(options=options)
driver.maximize_window()
wait = WebDriverWait(driver, 20)

def esperar_continuar():
    root = tk.Tk()
    root.title("Login Completo")
    root.geometry("250x100")

    def continuar():
        root.destroy()

    botao = tk.Button(root, text="Continuar", command=continuar, width=20, height=2)
    botao.pack(expand=True)
    root.mainloop()

def esperar_e_encontrar(by, seletor, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, seletor))
    )

def login(usuario, senha):
    driver.get("https://esaj.tjsp.jus.br/sajcas/login?service=https%3A%2F%2Fesaj.tjsp.jus.br%2Fesaj%2Fj_spring_cas_security_check")
    esperar_e_encontrar(By.ID, "usernameForm").send_keys(usuario)
    esperar_e_encontrar(By.ID, "passwordForm").send_keys(senha)
    esperar_e_encontrar(By.ID, "pbEntrar").click()
    esperar_continuar()

def preencher_numero_processo(numero_processo, delay=0.5):
    digits = re.sub(r'\D', '', numero_processo)
    if len(digits) < 17:
        raise ValueError(f"Número incompleto após limpeza: '{digits}' (len={len(digits)})")
    parte1 = digits[:13]
    parte2 = digits[16:]

    campo1 = esperar_e_encontrar(By.ID, "numeroDigitoAnoUnificado")
    campo1.clear()
    campo1.click()
    time.sleep(delay)
    campo1.send_keys(parte1)
    time.sleep(0.2)
    campo1.send_keys(Keys.TAB)
    time.sleep(3)

    campo2 = esperar_e_encontrar(By.ID, "foroNumeroUnificado")
    try:
        campo2.click()
    except Exception:
        pass
    time.sleep(delay)
    campo2.send_keys(parte2)
    time.sleep(0.2)

def pesquisar_processo(numero_processo):
    driver.get("https://esaj.tjsp.jus.br/cpopg/open.do?gateway=true")
    time.sleep(4)
    preencher_numero_processo(numero_processo)
    time.sleep(2)
    esperar_e_encontrar(By.ID, "botaoConsultarProcessos").click()
    time.sleep(5)

def abrir_autos():
    visualizar_autos = esperar_e_encontrar(By.ID, "linkPasta")
    ActionChains(driver)\
        .key_down(Keys.CONTROL)\
        .click(visualizar_autos)\
        .key_up(Keys.CONTROL)\
        .perform()
    time.sleep(3)

def gerar_pdf(numero_processo):
    # Cria pasta do processo
    pasta_processo = os.path.join(pasta_principal, numero_processo)
    os.makedirs(pasta_processo, exist_ok=True)

    # Redefine a pasta de download para essa execução
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": pasta_processo}
    )

    abas = driver.window_handles
    driver.switch_to.window(abas[1])
    wait = WebDriverWait(driver, 60)
    div_botoes = wait.until(EC.presence_of_element_located((By.ID, "divBotoesInterna")))

    div_botoes.find_element(By.ID, "selecionarButton").click()
    time.sleep(2)
    div_botoes.find_element(By.ID, "salvarButton").click()
    time.sleep(3)
    botao_continuar = wait.until(EC.element_to_be_clickable((By.ID, "botaoContinuar")))
    botao_continuar.click()
    time.sleep(5)
    botao_download = wait.until(EC.element_to_be_clickable((By.ID, "btnDownloadDocumento")))
    botao_download.click()
    print(f"✅ Documento do processo {numero_processo} salvo em {pasta_processo}!")
    time.sleep(2)

def main():
    try:
        caminho_json = os.path.join('resultados.json')
        with open(caminho_json, 'r', encoding='utf-8') as f:
            dados_processos = json.load(f)

        login(usuario, senha)
        for processo in dados_processos[:3]:
            numero_processo = str(processo['processo']).strip()
            print(f"🔹 Processando: {numero_processo}")
            pesquisar_processo(numero_processo)
            time.sleep(3)
            abrir_autos()
            gerar_pdf(numero_processo)
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
    except Exception as e:
        print(f"⚠️ Erro: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
