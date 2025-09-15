import subprocess
import os
import shutil
import requests
import zipfile
import tempfile
import tkinter as tk
from tkcalendar import DateEntry
from tkinter import messagebox
import os
import json
import glob
import unicodedata
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
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
import datetime, sys

hoje = datetime.date.today()
if hoje > datetime.date(2025, 9, 20):
    # Cria janela do Tkinter
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    messagebox.showwarning("Licença expirada", "⚠️ Versão expirada, peça nova licença.")
    root.destroy()
    sys.exit()

def verificar_credencial_remota():
    print("Verificando credencial remota...")
    url = "https://raw.githubusercontent.com/KauanSundays/luciano_credencial/refs/heads/main/credendial.txt"
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    try:
        resp = requests.get(url, timeout=5)
        # Se status code não for 200, significa que o arquivo não existe ou não está acessível
        if resp.status_code != 200:
            raise Exception(f"Erro, favor procurar desenvolvedor responsavel Kauan! HTTP {resp.status_code}")

        conteudo = resp.text.strip()
        if "phillySpecial12" not in conteudo:
            raise Exception("Chave/licença não encontrada no arquivo remoto!")

    except Exception as e:
        messagebox.showerror("Verificação falhou", f"❌ Execução interrompida: {e}")
        root.destroy()
        sys.exit()

    root.destroy()
    print("✅ Licença válida, continuando o script...")

verificar_credencial_remota()
# subprocess.run(["python", "1_baixar_diario.py"])
# subprocess.run(["python", "2_procurar_arquivos.py"])
# subprocess.run(["python", "3_esaj.py"])
# subprocess.run(["python", "4_entrar_assertiva.py"])
def baixar_diario():
  def log(mensagem):
    print(mensagem)
    status_label.config(text=mensagem)
    root.update_idletasks()

  def copiar_conteudo(origem, destino):
      os.makedirs(destino, exist_ok=True)
      for item in os.listdir(origem):
          src = os.path.join(origem, item)
          dst = os.path.join(destino, item)
          try:
              if os.path.isdir(src):
                  shutil.copytree(src, dst, dirs_exist_ok=True)
              else:
                  shutil.copy2(src, dst)
          except Exception as e:
              log(f"⚠️ Erro ao copiar {src}: {e}")

  def baixar_caderno():
      try:
          data = campo_data.get()
          nome_pasta = campo_pasta.get().strip()

          if not data:
              messagebox.showwarning("Aviso", "Selecione uma data primeiro.")
              return
          if not nome_pasta:
              messagebox.showwarning("Aviso", "Digite o nome da pasta.")
              return

          log(f"🔎 Consultando API para {data}...")
          url_api = f"https://comunicaapi.pje.jus.br/api/v1/caderno/TJSP/{data}/D"
          r = requests.get(url_api)
          r.raise_for_status()
          dados = r.json()
          url_zip = dados["url"]

          nome_zip = f"TJSP-D-{data}_v1.zip"
          caminho_zip = os.path.join(os.getcwd(), nome_zip)

          if os.path.exists(caminho_zip):
              log("📂 Arquivo já existe, pulando download...")
          else:
              log("⬇️ Baixando arquivo ZIP...")
              resp = requests.get(url_zip, stream=True)
              with open(caminho_zip, "wb") as f:
                  for chunk in resp.iter_content(chunk_size=8192):
                      f.write(chunk)

          log("📦 Extraindo arquivo...")
          pasta_arquivos = os.path.join(os.getcwd(), nome_pasta)
          if os.path.exists(pasta_arquivos):
              log("🗑️ Removendo pasta existente...")
              shutil.rmtree(pasta_arquivos)

          temp_dir = tempfile.mkdtemp()
          with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
              zip_ref.extractall(temp_dir)

          log(f"📂 Copiando conteúdo extraído para '{pasta_arquivos}'...")
          copiar_conteudo(temp_dir, pasta_arquivos)

          log("🧹 Limpando temporários...")
          shutil.rmtree(temp_dir, ignore_errors=True)

          log("✅ Concluído!")
          messagebox.showinfo("Sucesso", f"Arquivos extraídos em: {pasta_arquivos}")
          root.destroy()  # Fecha a janela Tkinter

      except Exception as e:
          log(f"❌ Erro: {e}")
          messagebox.showerror("Erro", str(e))


  root = tk.Tk()
  root.title("Baixar Caderno TJSP")
  root.geometry("400x250")

  tk.Label(root, text="Selecione a Data:").pack(pady=5)
  campo_data = DateEntry(root, date_pattern="yyyy-mm-dd", width=12)
  campo_data.pack(pady=5)

  tk.Label(root, text="Nome da Pasta:").pack(pady=5)
  campo_pasta = tk.Entry(root, width=25)
  campo_pasta.insert(0, "arquivos")
  campo_pasta.pack(pady=5)

  btn_iniciar = tk.Button(root, text="Iniciar", command=baixar_caderno, width=20, height=2)
  btn_iniciar.pack(pady=10)

  status_label = tk.Label(root, text="Aguardando...", fg="blue")
  status_label.pack(pady=10)

  root.mainloop()
  pass

def procurar_arquivos():
    def normalizar(texto):
        return "".join(
            c for c in unicodedata.normalize("NFD", texto.lower())
            if unicodedata.category(c) != "Mn"
        )

    # grupos fixos iniciais
    grupos = [
        ["inss", "requisitorio"],
        ["pensão por morte", "requisitorio"]
    ]

    def executar_busca():
        pasta = os.path.join(os.getcwd(), "arquivos")
        if not os.path.exists(pasta):
            messagebox.showerror("Erro", f"Pasta '{pasta}' não encontrada.")
            return

        arquivos = sorted(glob.glob(os.path.join(pasta, "TJSP-D-*.json")))
        if not arquivos:
            messagebox.showwarning("Aviso", "Nenhum arquivo JSON encontrado na pasta.")
            return

        # pega grupos fixos marcados
        grupos_selecionados = [grupos[i] for i, var in enumerate(vars_grupos) if var.get() == 1]

        # pega grupos digitados no campo
        texto_extra = campo_grupos_extra.get("1.0", tk.END).strip()
        if texto_extra:
            for linha in texto_extra.splitlines():
                palavras = [p.strip().lower() for p in linha.split(",") if p.strip()]
                if palavras:
                    grupos_selecionados.append(palavras)

        if not grupos_selecionados:
            messagebox.showerror("Erro", "Selecione ou escreva ao menos um grupo.")
            return

        # pega limite de processos
        try:
            limite = int(campo_limite.get().strip())
        except:
            limite = None

        resultados = []
        total = len(arquivos)
        progress_bar["maximum"] = total
        processos_encontrados = 0

        for i, arquivo in enumerate(arquivos, start=1):
            with open(arquivo, "r", encoding="utf-8") as f:
                data = json.load(f)
                for item in data.get("items", []):
                    texto_norm = normalizar(str(item.get("texto", "")))
                    for grupo in grupos_selecionados:
                        if all(chave in texto_norm for chave in grupo):
                            processos_encontrados += 1
                            resultados.append({
                                "arquivo": os.path.basename(arquivo),
                                "processo": item.get("numero_processo"),
                                "texto": item.get("texto")
                            })
                            break

                    if limite and processos_encontrados >= limite:
                        break
            if limite and processos_encontrados >= limite:
                break

            progress_bar["value"] = i
            status_var.set(f"Analisados {i}/{total} arquivos | Processos encontrados: {processos_encontrados}")
            root.update_idletasks()

        if resultados:
            # Salva em JSON
            save_path_json = os.path.join(os.getcwd(), "resultados.json")
            with open(save_path_json, "w", encoding="utf-8") as f:
                json.dump(resultados, f, ensure_ascii=False, indent=4)

            # Salva em Excel
            save_path_excel = os.path.join(os.getcwd(), "resultados_processos.xlsx")
            df = pd.DataFrame(resultados)
            df.to_excel(save_path_excel, index=False)

            messagebox.showinfo(
                "Concluído",
                f"Busca concluída.\n{len(resultados)} ocorrências salvas em:\n\n{save_path_excel}"
            )
        else:
            messagebox.showinfo("Resultado", "Nenhuma ocorrência encontrada.")

    # --- GUI ---
    root = tk.Tk()
    root.title("Busca em JSON -> Excel")

    # tela cheia
    try:
        root.state("zoomed")  # Windows
    except:
        root.attributes("-fullscreen", True)  # Linux/Mac

    status_var = tk.StringVar(value="Aguardando início...")

    frame_grupos = ttk.LabelFrame(root, text="Grupos de Palavras Fixos", padding=10)
    frame_grupos.pack(fill="x", padx=10, pady=5)

    vars_grupos = []
    for i, grupo in enumerate(grupos):
        var = tk.IntVar(value=1)
        chk = ttk.Checkbutton(frame_grupos, text=", ".join(grupo), variable=var)
        chk.pack(anchor="w")
        vars_grupos.append(var)

    frame_extras = ttk.LabelFrame(root, text="Escreva mais combinações (1 por linha, separadas por vírgula)", padding=10)
    frame_extras.pack(fill="x", padx=10, pady=5)

    campo_grupos_extra = tk.Text(frame_extras, height=5)
    campo_grupos_extra.pack(fill="x", padx=5, pady=5)

    frame_limite = ttk.LabelFrame(root, text="Limite de Processos (opcional)", padding=10)
    frame_limite.pack(fill="x", padx=10, pady=5)

    campo_limite = tk.Entry(frame_limite, width=10)
    campo_limite.pack(padx=5, pady=5)

    ttk.Button(root, text="Executar Busca", command=executar_busca).pack(pady=10)

    progress_bar = ttk.Progressbar(root, length=400, mode="determinate")
    progress_bar.pack(pady=5)

    status_label = ttk.Label(root, textvariable=status_var)
    status_label.pack(pady=5)

    root.mainloop()
    
def esaj():
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

    def ler_credenciais(arquivo="credenciais.txt"):
        usuario = ""
        senha = ""
        with open(arquivo, "r", encoding="utf-8") as f:
            for linha in f:
                linha = linha.strip()
                if linha.startswith("usuario="):
                    usuario = linha.split("=", 1)[1].strip()
                elif linha.startswith("senha="):
                    senha = linha.split("=", 1)[1].strip()
            return usuario, senha

    # --- Criar pasta principal se não existir ---
    import shutil

    # --- Criar pasta principal se não existir ---
    pasta_principal = os.path.join(os.getcwd(), "processos_pdfs")
    os.makedirs(pasta_principal, exist_ok=True)

    # --- Limpar todas as subpastas e arquivos dentro de processos_pdfs ---
    for item in os.listdir(pasta_principal):
        item_path = os.path.join(pasta_principal, item)
        if os.path.isdir(item_path):
            shutil.rmtree(item_path)  # remove pasta inteira
        else:
            os.remove(item_path)  # remove arquivo solto

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

        # Abre a aba da extensão
        driver.execute_script("window.open('https://chromewebstore.google.com/detail/web-signer/bbafmabaelnnkondpfpjmdklbmfnbmol?hl=en', '_blank');")
        driver.switch_to.window(driver.window_handles[-1])

        # Chama a janela de confirmação
        esperar_continuar()

        # Após o usuário clicar em "Continuar", fecha a aba atual (Web Store)
        driver.close()

        # Volta para a primeira aba (ESAJ)
        driver.switch_to.window(driver.window_handles[0])
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
            usuario, senha = ler_credenciais("credenciais.txt")

            login(usuario, senha)
            for processo in dados_processos:
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

def entrar_assertiva():
    import os
    import re
    import pdfplumber
    import tkinter as tk
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service
    import pandas as pd

    # --- Configuração do Selenium ---
    options = Options()
    options.add_argument("--start-maximized")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 20)

    # --- Função para extrair CPFs de PDFs ---
    def extrair_cpfs(pdf_path):
        cpfs = set()
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    encontrados = re.findall(r'\d{3}\.\d{3}\.\d{3}-\d{2}', text)
                    cpfs.update(encontrados)
        return list(cpfs)

    # --- Login ---
    driver.get("https://localize.assertivasolucoes.com.br/login")

    # --- Janela de confirmação do login ---
    def esperar_continuar():
        root = tk.Tk()
        root.title("Login Completo")
        root.geometry("300x100")

        def continuar():
            root.destroy()

        botao = tk.Button(root, text="Continuar", command=continuar, width=20, height=2)
        botao.pack(expand=True)
        root.mainloop()

    esperar_continuar()

    # --- Janela mostrando que está processando os PDFs ---
    root_processando = tk.Tk()
    root_processando.title("Raspando PDFs")
    root_processando.geometry("300x100")

    label_processando = tk.Label(root_processando, text="Raspando PDFs...", font=("Arial", 12))
    label_processando.pack(expand=True, pady=20)

    # Atualiza a interface
    root_processando.update()

    # --- Extrair todos os CPFs dos PDFs ---
    todos_cpfs = []
    pasta = "processos_pdfs"
    for root_dir, dirs, files in os.walk(pasta):
        for file in files:
            if not file.endswith(".pdf"):
                continue
            caminho_pdf = os.path.join(root_dir, file)
            numero_processo = os.path.splitext(file)[0]
            cpfs = extrair_cpfs(caminho_pdf)
            for cpf in cpfs:
                todos_cpfs.append({"processo": numero_processo, "cpf": cpf})

    # Depois que termina, fecha a janela de processamento
    root_processando.destroy()

    # --- Janela mostrando quantidade de CPFs encontrados ---
    def mostrar_total_cpfs(total):
        root = tk.Tk()
        root.title("CPFs Encontrados")
        root.geometry("300x100")

        label = tk.Label(root, text=f"Foram encontrados {total} CPFs", font=("Arial", 12))
        label.pack(pady=10)

        def continuar():
            root.destroy()

        botao = tk.Button(root, text="OK", command=continuar, width=15)
        botao.pack(pady=10)

        root.mainloop()

    mostrar_total_cpfs(len(todos_cpfs))

    # --- Checklist agrupado por processo ---
    def selecionar_cpfs_por_processo(cpfs_por_processo):
        selecionados = []

        # Organiza os CPFs por processo
        processos = {}
        for item in cpfs_por_processo:
            processos.setdefault(item['processo'], []).append(item['cpf'])

        root = tk.Tk()
        root.title("Selecionar CPFs para consultar")
        root.geometry("600x400")

        canvas = tk.Canvas(root)
        frame = tk.Frame(canvas)
        scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0,0), window=frame, anchor='nw')

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame.bind("<Configure>", on_frame_configure)

        vars_check = []
        for processo, cpfs in processos.items():
            # Label do processo
            lbl = tk.Label(frame, text=f"Processo: {processo}", font=("Arial", 10, "bold"), anchor="w", justify="left")
            lbl.pack(fill="x", padx=5, pady=2)

            # Checkboxes dos CPFs
            for cpf in cpfs:
                var = tk.BooleanVar(value=True)
                chk = tk.Checkbutton(frame, text=f"CPF: {cpf}", variable=var, anchor="w", justify="left")
                chk.pack(fill="x", padx=25, pady=1)
                vars_check.append((var, processo, cpf))

        def continuar():
            for var, processo, cpf in vars_check:
                if var.get():
                    selecionados.append({"processo": processo, "cpf": cpf})
            root.destroy()

        btn = tk.Button(root, text="Continuar", command=continuar, width=20, height=2)
        btn.pack(pady=10)

        root.mainloop()
        return selecionados

    cpfs_para_consultar = selecionar_cpfs_por_processo(todos_cpfs)

    # --- Consultar apenas os CPFs selecionados ---
    resultados = []
    for item in cpfs_para_consultar:
        cpf = item["cpf"]
        numero_processo = item["processo"]
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
                EC.presence_of_element_located((By.ID, "identificacao-cadastral-name"))
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
            continue

    # --- Salvar em Excel ---
    if resultados:
        df = pd.DataFrame(resultados)
        df.to_excel("resultados.xlsx", index=False)
        print("Excel gerado com sucesso!")
    else:
        print("Nenhum resultado encontrado.")

    # --- Fechar driver ao final ---
    driver.quit()
    
def main():
    # Aqui você decide se quer interface GUI para todas ou execução sequencial
    # baixar_diario()
    # procurar_arquivos()
    # esaj()
    entrar_assertiva()

if __name__ == "__main__":
    main()
