def normalizar(texto):
    return "".join(
        c for c in unicodedata.normalize("NFD", texto.lower())
        if unicodedata.category(c) != "Mn"
    )

# grupos de palavras
grupos = [
    ["inss", "requisitorio"],
    ["pensão por morte", "requisitorio"]
]

def executar_busca():
    pasta = os.path.join(os.getcwd(), "arquivos")  # pasta fixa
    if not os.path.exists(pasta):
        messagebox.showerror("Erro", f"Pasta '{pasta}' não encontrada.")
        return

    arquivos = sorted(glob.glob(os.path.join(pasta, "TJSP-D-*.json")))
    if not arquivos:
        messagebox.showwarning("Aviso", "Nenhum arquivo JSON encontrado na pasta.")
        return

    grupos_selecionados = [grupos[i] for i, var in enumerate(vars_grupos) if var.get() == 1]
    if not grupos_selecionados:
        messagebox.showerror("Erro", "Selecione ao menos um grupo.")
        return

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

        progress_bar["value"] = i
        status_var.set(f"Analisados {i}/{total} arquivos | Processos encontrados: {processos_encontrados}")
        root.update_idletasks()

    if resultados:
        save_path = os.path.join(os.getcwd(), "resultados.json")
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(resultados, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("Concluído", f"Busca concluída.\n{len(resultados)} ocorrências salvas em:\n{save_path}")
    else:
        messagebox.showinfo("Resultado", "Nenhuma ocorrência encontrada.")


# --- GUI ---
root = tk.Tk()
root.title("Busca em JSON -> Excel")

status_var = tk.StringVar(value="Aguardando início...")

frame_grupos = ttk.LabelFrame(root, text="Grupos de Palavras", padding=10)
frame_grupos.pack(fill="x", padx=10, pady=5)

vars_grupos = []
for i, grupo in enumerate(grupos):
    var = tk.IntVar(value=1)  # marcados por padrão
    chk = ttk.Checkbutton(frame_grupos, text=", ".join(grupo), variable=var)
    chk.pack(anchor="w")
    vars_grupos.append(var)

ttk.Button(root, text="Executar Busca", command=executar_busca).pack(pady=10)

progress_bar = ttk.Progressbar(root, length=400, mode="determinate")
progress_bar.pack(pady=5)

status_label = ttk.Label(root, textvariable=status_var)
status_label.pack(pady=5)

root.mainloop()
