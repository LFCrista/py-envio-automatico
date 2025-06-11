import os
import time
import re
import threading
from tkinter import Tk, Label, Button, filedialog, Text, ttk, messagebox
from playwright.sync_api import sync_playwright
from docx import Document

# --- Funções Utilitárias ---
def get_pdf_files(pasta):
    return [os.path.join(pasta, f) for f in os.listdir(pasta)
            if os.path.isfile(os.path.join(pasta, f)) and f.lower().endswith('.pdf')]

def extrair_indice_final(nome_arquivo):
    match = re.search(r'-(\d+)\.pdf$', nome_arquivo)
    return int(match.group(1)) if match else 0

def salvar_texto_docx(respostas_dict, destino):
    doc = Document()
    for nome_arquivo, texto in respostas_dict.items():
        doc.add_heading(f"Arquivo: {nome_arquivo}", level=1)
        doc.add_paragraph(texto)
        doc.add_paragraph("\n")
    doc.save(destino)

def houve_erro_visual(page):
    # Verificação por classe CSS de erro no DOM
    elementos = page.locator(".text-token-text-error")
    if elementos.count() > 0:
        for i in range(elementos.count()):
            color = page.evaluate("(el) => getComputedStyle(el).color", elementos.nth(i))
            if "rgb" in color or "var(--text-error)" in color:
                print("🚫 Erro visual detectado: classe 'text-token-text-error' com cor de erro.")
                return True
    return False

def chat_esta_gerando(page):
    return page.locator("button:has(svg[aria-label='Stop generating'])").is_visible()

def ha_arquivo_anexado(page):
    return page.locator("div[role='listitem']").is_visible()

def esperar_resposta_gpt(page, tempo_maximo=180, intervalo_check=1.5, tempo_estavel=4):
    conteudo_anterior = ""
    tentativas_estaveis = 0
    tempo_decorrido = 0

    while tempo_decorrido < tempo_maximo:
        if page.url.endswith("/api/auth/error"):
            return "__ERRO_AUTENTICACAO__"
        if houve_erro_visual(page):
            return "__ERRO_GPT__"
        if chat_esta_gerando(page):
            time.sleep(intervalo_check)
            tempo_decorrido += intervalo_check
            continue

        elementos = page.locator(".markdown")
        respostas = elementos.all_text_contents()
        conteudo_atual = respostas[-1] if respostas else ""

        if conteudo_atual == conteudo_anterior:
            tentativas_estaveis += 1
        else:
            tentativas_estaveis = 0

        if tentativas_estaveis >= tempo_estavel:
            return conteudo_atual if conteudo_atual else "__ERRO_GPT__"

        conteudo_anterior = conteudo_atual
        time.sleep(intervalo_check)
        tempo_decorrido += intervalo_check

    return conteudo_anterior if conteudo_anterior else "__ERRO_GPT__"

def enviar_pdf_para_gpt(page, caminho_pdf):
    if ha_arquivo_anexado(page):
        return "__ARQUIVO_JA_ANEXADO__"

    try:
        # Verifica se o botão de upload existe e está habilitado
        botao_upload = page.locator("button:has(svg[aria-label='Upload a file'])")
        if botao_upload.count() == 0:
            print("⚠️ Botão de upload não encontrado no DOM.")
            return "__BOTAO_UPLOAD_INEXISTENTE__"
        if not botao_upload.is_enabled():
            print("⚠️ Botão de upload está desabilitado no momento.")
            return "__BOTAO_UPLOAD_DESABILITADO__"

        # Define o arquivo diretamente no input file
        page.set_input_files("input[type='file']", caminho_pdf)
        time.sleep(4)
        page.keyboard.type("T2")
        page.keyboard.press("Enter")
        return esperar_resposta_gpt(page)

    except Exception as e:
        print(f"⚠️ Erro inesperado ao enviar o arquivo '{os.path.basename(caminho_pdf)}': {e}")
        return "__ERRO_ENVIO_DESCONHECIDO__"

# --- Interface Gráfica ---
class App:
    def __init__(self, master):
        self.master = master
        master.title("Automatizador de PDFs para ChatGPT")

        self.label = Label(master, text="Selecione a pasta com PDFs:")
        self.label.pack(pady=5)

        self.select_button = Button(master, text="Selecionar Pasta", command=self.selecionar_pasta)
        self.select_button.pack(pady=5)

        self.run_button = Button(master, text="Iniciar Processamento", command=self.iniciar_processamento)
        self.run_button.pack(pady=10)

        self.progress = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        self.log = Text(master, height=15, width=70)
        self.log.pack(pady=10)

        self.pasta = None

    def selecionar_pasta(self):
        self.pasta = filedialog.askdirectory()
        if self.pasta:
            self.log.insert("end", f"📁 Pasta selecionada: {self.pasta}\n")
        else:
            self.log.insert("end", "❌ Nenhuma pasta selecionada.\n")

    def iniciar_processamento(self):
        if not self.pasta:
            messagebox.showwarning("Aviso", "Selecione uma pasta antes de continuar.")
            return
        threading.Thread(target=self.executar_processamento).start()

    def executar_processamento(self):
        arquivos_pdf = get_pdf_files(self.pasta)
        if not arquivos_pdf:
            self.log.insert("end", "❌ Nenhum PDF encontrado na pasta.\n")
            return

        arquivos_pdf = sorted(arquivos_pdf, key=lambda x: extrair_indice_final(os.path.basename(x)))
        self.progress["maximum"] = len(arquivos_pdf)
        respostas = {}

        self.log.insert("end", f"🚀 Iniciando envio de {len(arquivos_pdf)} arquivos...\n")

        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            page = browser.contexts[0].pages[0]

            for idx, pdf in enumerate(arquivos_pdf):
                nome = os.path.basename(pdf)
                self.log.insert("end", f"📎 Processando: {nome}\n")
                self.log.see("end")

                try:
                    resposta = enviar_pdf_para_gpt(page, pdf)
                    if resposta.startswith("__ERRO"):
                        self.log.insert("end", f"⚠️ Erro ao processar {nome}: {resposta}\n")
                    else:
                        self.log.insert("end", f"✅ {nome} processado com sucesso.\n")
                    respostas[nome] = resposta
                except Exception as e:
                    self.log.insert("end", f"❌ Erro inesperado em {nome}: {e}\n")
                    print(e)
                    respostas[nome] = str(e)

                self.progress["value"] = idx + 1

        destino_docx = os.path.join(self.pasta, "texto_corrigido_final.docx")
        salvar_texto_docx(respostas, destino_docx)
        self.log.insert("end", f"\n✅ Documento salvo em: {destino_docx}\n")
        messagebox.showinfo("Concluído", f"Processamento finalizado. Arquivo salvo:\n{destino_docx}")

# --- Execução ---
if __name__ == "__main__":
    root = Tk()
    app = App(root)
    root.mainloop()
