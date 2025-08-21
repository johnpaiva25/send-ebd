# Importação de bibliotecas necessárias
import requests
import json
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import threading
import sys
import os

# Caminho do arquivo contendo a lista de CPFs dos membros
ARQUIVO_CPF = "cpfMembros.txt"

# Função para obter o ID da EBD automaticamente a partir do site
def obter_ebd_id():
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')  # Executa o navegador sem interface gráfica

        # Ajuste de caminho para compatibilidade com programas empacotados (ex: PyInstaller)
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        # Caminho do driver do Chrome
        chromedriver_path = os.path.join(base_path, "chromedriver.exe")
        service = Service(chromedriver_path)

        # Inicializa o navegador
        driver = webdriver.Chrome(service=service, options=options)
        driver.get("https://www.igrejacristamaranata.org.br/ebd/participacoes/")
        driver.implicitly_wait(5)

        # Captura o valor do campo de input que contém o ebd_id
        input_tag = driver.find_element(By.NAME, "icm_member_ebd")
        ebd_id = input_tag.get_attribute("value")
        driver.quit()
        return ebd_id
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao obter EBD ID: {e}")
        return None

# Função responsável por enviar a contribuição para todos os CPFs cadastrados
def enviar_contribuicoes(contribuicao_html, ebd_id, progressbar, status_label, root):
    url_envia = 'https://intregracao-site.presbiterio.org.br/api-ebd/cadastro-contribuicao'
    resultados = []

    # Exibe no terminal o conteúdo da contribuição em HTML
    print("\n======= Contribuição enviada (HTML) =======\n")
    print(contribuicao_html)
    print("===========================================\n")

    try:
        # Lê os CPFs do arquivo
        with open(ARQUIVO_CPF, "r") as file:
            cpfs = [linha.strip() for linha in file if linha.strip()]

        total_cpfs = len(cpfs)
        for i, cpf in enumerate(cpfs, start=1):
            # Consulta os dados do membro com base no CPF
            url_dados = f"https://intregracao-site.presbiterio.org.br/api-ebd/consultacpf/{cpf}"
            response = requests.get(url_dados)
            if response.status_code != 200:
                resultados.append((cpf, None, "Falhou"))
                continue

            data = response.json()
            nome = data.get('nome', 'Nome não encontrado')

            # Atualiza o status na interface
            status_label.config(text=f"Enviando ({i}/{total_cpfs}): {nome}")

            # Monta o payload com os dados do membro e a contribuição
            payload = json.dumps({
                "nome": nome,
                "cpf": cpf,
                "denominacao_id": 21,
                "celular": data.get('celular'),
                "email": data.get('email'),
                "cidade": data.get('cidade'),
                "uf": data.get('uf'),
                "ebd_id": ebd_id,
                "categoria_id": "2",
                "contribuicao": contribuicao_html,
                "aceite_termo": True
            })

            headers = {'Content-Type': 'application/json'}

            # Envia os dados para a API
            response_envio = requests.post(url_envia, headers=headers, data=payload)
            status = "Sucesso" if response_envio.status_code == 200 else "Falhou"
            status = "Sucesso"
            resultados.append((cpf, nome, status))

            # Atualiza a barra de progresso
            progressbar['value'] = (i / total_cpfs) * 100
            root.update_idletasks()

        # Finaliza a operação com mensagem e exibição dos resultados
        status_label.config(text="Envio finalizado.")
        mostrar_resultados(resultados)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar contribuições: {e}")

# Função que exibe uma nova janela com o resumo dos envios realizados
def mostrar_resultados(resultados):
    janela = tk.Toplevel()
    janela.title("Resultados dos Envios")
    janela.geometry("600x400")

    label = tk.Label(janela, text="Resumo dos envios", font=("Arial", 12, "bold"))
    label.pack(pady=10)

    # Tabela para mostrar nome, CPF e status
    columns = ("nome", "cpf", "status")
    tree = ttk.Treeview(janela, columns=columns, show="headings", height=15)

    tree.heading("nome", text="Nome do Membro")
    tree.heading("cpf", text="CPF")
    tree.heading("status", text="Status do Envio")

    tree.column("nome", width=250)
    tree.column("cpf", width=120)
    tree.column("status", width=100)

    for cpf, nome, status in resultados:
        tree.insert("", tk.END, values=(nome or "Desconhecido", cpf, status))

    tree.pack(expand=True, fill="both", padx=10, pady=10)

    # Scrollbar vertical
    scrollbar = ttk.Scrollbar(janela, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side="right", fill="y")

# Função usada quando o arquivo de CPFs está ausente ou vazio
def solicitar_cpfs(callback_iniciar_interface):
    def salvar_cpfs():
        conteudo = text_area.get("1.0", tk.END).strip()
        if not conteudo:
            messagebox.showwarning("Aviso", "Nenhum CPF inserido.")
            return
        with open(ARQUIVO_CPF, "w") as f:
            for linha in conteudo.split("\n"):
                cpf = linha.strip()
                if cpf:
                    f.write(cpf + "\n")
        janela.destroy()
        callback_iniciar_interface()

    janela = tk.Tk()
    janela.title("Cadastro de CPFs")
    janela.geometry("500x350")

    tk.Label(janela, text="Arquivo de CPFs não encontrado ou vazio.\nInsira os CPFs abaixo (um por linha):", font=("Arial", 12)).pack(pady=10)

    text_area = scrolledtext.ScrolledText(janela, wrap=tk.WORD, width=60, height=12, font=("Arial", 10))
    text_area.pack(padx=10, pady=10)

    btn_salvar = tk.Button(janela, text="Salvar CPFs", command=salvar_cpfs, bg="blue", fg="white", font=("Arial", 11))
    btn_salvar.pack(pady=10)

    janela.mainloop()

# Função principal que cria a interface principal da aplicação
def iniciar_interface():
    def on_enviar():
        texto = text_area.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Nenhuma contribuição inserida.")
            return

        # Converte o texto em HTML (parágrafos)
        linhas = texto.split("\n")
        texto_html = "".join(f"<p>{linha.strip()}</p>" for linha in linhas if linha.strip())

        root.withdraw()  # Oculta a janela principal

        ebd_id = obter_ebd_id()
        if not ebd_id:
            root.deiconify()
            return

        # Cria uma nova janela de progresso
        progress_win = tk.Toplevel()
        progress_win.title("Enviando Contribuições")
        progress_win.geometry("400x150")

        status_label = tk.Label(progress_win, text="Iniciando envio...")
        status_label.pack(pady=10)

        progressbar = ttk.Progressbar(progress_win, orient="horizontal", length=300, mode="determinate")
        progressbar.pack(pady=10)

        # Inicia o envio em uma thread separada para não travar a interface
        threading.Thread(target=lambda: enviar_contribuicoes(texto_html, ebd_id, progressbar, status_label, progress_win)).start()

    def on_close():
        root.destroy()
        sys.exit(0)

    # Criação da interface principal
    global root
    root = tk.Tk()
    root.title("Contribuição EBD")
    root.geometry("600x450")
    root.protocol("WM_DELETE_WINDOW", on_close)

    tk.Label(root, text="Digite a contribuição abaixo:", font=("Arial", 12)).pack(pady=10)

    global text_area
    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=15, font=("Arial", 10))
    text_area.pack(padx=10, pady=10)

    btn_enviar = tk.Button(root, text="Enviar Contribuições", command=on_enviar, bg="green", fg="white", font=("Arial", 12))
    btn_enviar.pack(pady=10)

    root.mainloop()

# Execução principal do programa
if __name__ == "__main__":
    # Verifica se o arquivo de CPFs é inexistente ou está vazio
    def arquivo_cpf_invalido():
        if not os.path.exists(ARQUIVO_CPF):
            return True
        with open(ARQUIVO_CPF, "r") as f:
            linhas = [linha.strip() for linha in f if linha.strip()]
            return len(linhas) == 0

    # Decide entre solicitar CPFs ou iniciar a interface principal
    if arquivo_cpf_invalido():
        solicitar_cpfs(iniciar_interface)
    else:
        iniciar_interface()
