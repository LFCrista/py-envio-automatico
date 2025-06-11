import os
import sys
import time
import re
from playwright.sync_api import sync_playwright
from tkinter import Tk, filedialog
from docx import Document
from tqdm import tqdm

# --- Constantes de Erros ---
ERRO_GPT = "__ERRO_GPT__"
ERRO_AUTENTICACAO = "__ERRO_AUTENTICACAO__"
ARQUIVO_JA_ANEXADO = "__ARQUIVO_JA_ANEXADO__"
ERRO_ENVIO = "__ERRO_ENVIO__"
UPLOAD_DESABILITADO = "__UPLOAD_DESABILITADO__"

# --- Funções Utilitárias ---
def log(msg):
    print(f"[LOG] {msg}")

def selecionar_pasta():
    root = Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Selecione a pasta com arquivos PDF")

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
    erro_elements = page.locator(".text-token-text-error")
    if erro_elements.count() > 0:
        mensagens = erro_elements.all_text_contents()
        for msg in mensagens:
            log(f"❌ Erro detectado: {msg}")
        return True
    return False

def resposta_contem_erro(conteudo):
    erros = ["aguarde", "tente novamente", "espera", "erro", "carregar mais tarde"]
    return any(p in conteudo.lower() for p in erros)

def chat_esta_gerando(page):
    return page.locator("button:has(svg[aria-label='Stop generating'])").is_visible()

def ha_arquivo_anexado(page):
    return page.locator("div[role='listitem']").is_visible()

def upload_esta_desabilitado(page):
    try:
        upload_element = page.locator("#upload-file")
        if upload_element.count() == 0:
            log("⚠️ Elemento #upload-file não encontrado na página.")
            return False
        if upload_element.is_disabled():
            log("❌ O input de upload (#upload-file) está desabilitado. Cancelando envio.")
            return True
        return False
    except Exception as e:
        log(f"⚠️ Erro ao verificar estado do upload: {e}")
        return True

def esperar_resposta_gpt(page, tempo_maximo=180, intervalo_check=1.5, tempo_estavel=4):
    conteudo_anterior = ""
    tentativas_estaveis = 0
    tempo_decorrido = 0

    log("⏳ Aguardando resposta do GPT...")

    while tempo_decorrido < tempo_maximo:
        if page.url.endswith("/api/auth/error"):
            log("⚠️ Erro de autenticação detectado.")
            return ERRO_AUTENTICACAO

        if houve_erro_visual(page):
            log("❌ Erro crítico detectado na interface. Encerrando o script.")
            sys.exit(1)

        if chat_esta_gerando(page):
            log("⌛ Ainda gerando resposta...")
            time.sleep(1.5)
            tempo_decorrido += 1.5
            continue

        elementos = page.locator(".markdown")
        respostas = elementos.all_text_contents()
        conteudo_atual = respostas[-1] if respostas else ""

        if conteudo_atual == conteudo_anterior:
            tentativas_estaveis += 1
        else:
            tentativas_estaveis = 0

        if tentativas_estaveis >= tempo_estavel:
            log("✅ Resposta estabilizada.")
            if resposta_contem_erro(conteudo_atual):
                log("⚠️ Erro no conteúdo detectado.")
                return ERRO_GPT
            return conteudo_atual

        conteudo_anterior = conteudo_atual
        time.sleep(intervalo_check)
        tempo_decorrido += intervalo_check

    log("⚠️ Tempo máximo atingido.")
    if resposta_contem_erro(conteudo_anterior):
        log("⚠️ Erro na resposta final.")
        return ERRO_GPT
    return conteudo_anterior

def enviar_pdf_para_gpt(page, caminho_pdf):
    log(f"📎 Enviando PDF: {caminho_pdf}")

    if upload_esta_desabilitado(page):
        return UPLOAD_DESABILITADO

    tentativas = 0
    while ha_arquivo_anexado(page) and tentativas < 30:
        log("⏳ Aguardando remoção do arquivo anterior...")
        time.sleep(1.5)
        tentativas += 1

    if ha_arquivo_anexado(page):
        log("⚠️ Arquivo ainda anexado. Abortando.")
        return ARQUIVO_JA_ANEXADO

    try:
        page.click("button:has(svg[aria-label='Upload a file'])", timeout=5000)
        page.wait_for_selector("input[type='file']", timeout=5000)
    except Exception as e:
        log(f"⚠️ Falha ao clicar no botão de upload: {e}")
        if "Timeout 5000ms exceeded" in str(e):
            log("❌ Timeout crítico ao clicar no botão de upload. Encerrando o script.")
            sys.exit(1)
        return ERRO_ENVIO

    if upload_esta_desabilitado(page):
        return UPLOAD_DESABILITADO

    try:
        page.set_input_files("input[type='file']", caminho_pdf)
    except Exception as e:
        log(f"❌ Falha ao anexar o arquivo: {e}")
        return ERRO_ENVIO

    log("⌛ Aguardando reconhecimento do upload...")
    time.sleep(4)

    if upload_esta_desabilitado(page):
        return UPLOAD_DESABILITADO

    log("🛑 Upload desabilitado. Comando 'T2' não será enviado.")
    return UPLOAD_DESABILITADO

def processar_arquivos(arquivos_pdf, page):
    respostas = {}
    for pdf in tqdm(arquivos_pdf, desc="📄 Processando PDFs", unit="arquivo"):
        resposta = enviar_pdf_para_gpt(page, pdf)
        nome = os.path.basename(pdf)
        if resposta == UPLOAD_DESABILITADO:
            respostas[nome] = "[ERRO] Campo de upload desabilitado na interface."
        else:
            respostas[nome] = resposta
    return respostas

def main():
    pasta = selecionar_pasta()
    if not pasta:
        log("Nenhuma pasta selecionada. Encerrando.")
        return

    arquivos_pdf = get_pdf_files(pasta)
    if not arquivos_pdf:
        log("Nenhum PDF encontrado na pasta.")
        return

    arquivos_pdf = sorted(arquivos_pdf, key=lambda x: extrair_indice_final(os.path.basename(x)))

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
        except Exception as e:
            log(f"❌ Falha ao conectar ao navegador: {e}")
            sys.exit(1)

        context = browser.contexts[0]
        page = context.pages[0] if context.pages else context.new_page()

        respostas = processar_arquivos(arquivos_pdf, page)

    destino_docx = os.path.join(pasta, "texto_corrigido_final.docx")
    salvar_texto_docx(respostas, destino_docx)
    log(f"✅ Texto final salvo em: {destino_docx}")

if __name__ == "__main__":
    main()