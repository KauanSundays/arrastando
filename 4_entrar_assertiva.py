import os
import re
import time
import pdfplumber
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd


# Configurações do Selenium
options = Options()
options.add_argument("--start-maximized")
from selenium.webdriver.chrome.service import Service

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)


def extrair_cpfs(pdf_path):
    cpfs = set()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                encontrados = re.findall(r'\d{3}\.\d{3}\.\d{3}-\d{2}', text)
                cpfs.update(encontrados)
    return list(cpfs)

# Login
driver.get("https://localize.assertivasolucoes.com.br/login")
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "email"))).send_keys("ricardo.kauand@gmail.com")
driver.find_element(By.ID, "password").send_keys("@CenouraFria753")
driver.find_element(By.ID, "btn-entrar").click()
import tkinter as tk

def esperar_continuar():
    root = tk.Tk()
    root.title("Login Completo")
    root.geometry("250x100")

    def continuar():
        root.destroy()

    botao = tk.Button(root, text="Continuar", command=continuar, width=20, height=2)
    botao.pack(expand=True)
    root.mainloop()

esperar_continuar()


resultados = []

pasta = "processos_pdfs"
for root, dirs, files in os.walk(pasta):
    for file in files:
        if not file.endswith(".pdf"):
            continue

        caminho_pdf = os.path.join(root, file)
        numero_processo = os.path.splitext(file)[0]
        cpfs = extrair_cpfs(caminho_pdf)
        print(f"{len(cpfs)} CPFs encontrados em {file}")

        for cpf in cpfs:
          try:
              print("Procurando por CPF:", cpf)
              driver.get("https://localize.assertivasolucoes.com.br/consulta/cpf")

              WebDriverWait(driver, 10).until(
                  EC.presence_of_element_located((By.NAME, "currentSimpleFilterValue"))
              ).send_keys(cpf)

              WebDriverWait(driver, 10).until(
                  EC.element_to_be_clickable((By.ID, "select-finalidade-uso"))
              ).click()
              WebDriverWait(driver, 10).until(
                  EC.element_to_be_clickable((By.ID, "option-finalidade-0"))
              ).click()

              driver.find_element(By.XPATH, "//button[span/text()='CONSULTAR']").click()

              nome_elem = WebDriverWait(driver, 15).until(
                  EC.presence_of_element_located((By.CSS_SELECTOR, ".jss456"))
              )
              driver.execute_script("arguments[0].scrollIntoView(true);", nome_elem)
              nome = nome_elem.text.strip()


              telefones_elems = WebDriverWait(driver, 15).until(
                  EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.phone span"))
              )
              telefones = [tel.text.strip() for tel in telefones_elems if tel.text.strip()]
              telefones_str = ", ".join(telefones)

              resultados.append({
                  "processo": numero_processo,
                  "nome": nome,
                  "cpf": cpf,
                  "telefones": telefones_str
              })

          except Exception as e:
              print(f"Erro ao processar CPF {cpf}: {e}")
              continue  # pula para o próximo CPF

# Salva em Excel
df = pd.DataFrame(resultados)
df.to_excel("resultados.xlsx", index=False)
print("Excel gerado com sucesso!")
