import asyncio
import sys
import os
import time
import tempfile
import subprocess
import streamlit as st
from playwright.sync_api import sync_playwright
from docx import Document

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# --- Configurações do Chrome Debug ---
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_USER_DATA_DIR = r"C:\temp\chrome"
CHROME_REMOTE_DEBUGGING_PORT = 9222
CHROME_DEBUG_URL = f"http://localhost:{CHROME_REMOTE_DEBUGGING_PORT}"

# --- Constantes de Erros ---
ERRO_GPT = "__ERRO_GPT__"
ERRO_AUTENTICACAO = "__ERRO_AUTENTICACAO__"
ARQUIVO_JA_ANEXADO = "__ARQUIVO_JA_ANEXADO__"
ERRO_ENVIO = "__ERRO_ENVIO__"
UPLOAD_DESABILITADO = "__UPLOAD_DESABILITADO__"

# --- Funções ---
def abrir_chrome_com_debug():
    subprocess.Popen([
        CHROME_PATH,
        f"--remote-debugging-port={CHROME_REMOTE_DEBUGGING_PORT}",
        f"--user-data-dir={CHROME_USER_DATA_DIR}"
    ])

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
            print(f"[ERRO VISUAL] {msg}")
        return True
    return False

def upload_esta_desabilitado(page):
    try:
        upload_element = page.locator("input[type='file']")
        if upload_element.count() == 0:
            print("[AVISO] Campo input[type='file'] não encontrado.")
            return False
        if upload_element.is_disabled():
            print("[ERRO] Campo de upload está desabilitado.")
            return True
        return False
    except Exception as e:
        print(f"[ERRO] Falha ao verificar estado do upload: {e}")
        return True

def esperar_resposta_gpt(page, tempo_maximo=180, intervalo_check=1.5, tempo_estavel=4):
    conteudo_anterior = ""
    tentativas_estaveis = 0
    tempo_decorrido = 0

    while tempo_decorrido < tempo_maximo:
        if page.url.endswith("/api/auth/error"):
            return ERRO_AUTENTICACAO

        if houve_erro_visual(page):
            return ERRO_GPT

        if page.locator("button:has(svg[aria-label='Stop generating'])").is_visible():
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
            return conteudo_atual if conteudo_atual else ERRO_GPT

        conteudo_anterior = conteudo_atual
        time.sleep(intervalo_check)
        tempo_decorrido += intervalo_check

    return conteudo_anterior if conteudo_anterior else ERRO_GPT

def enviar_pdf_para_gpt(page, caminho_pdf):
    if upload_esta_desabilitado(page):
        return UPLOAD_DESABILITADO

    tentativas = 0
    while page.locator("div[role='listitem']").is_visible() and tentativas < 30:
        time.sleep(1.5)
        tentativas += 1

    if page.locator("div[role='listitem']").is_visible():
        return ARQUIVO_JA_ANEXADO

    try:
        if page.locator("input[type='file']").count() > 0:
            page.set_input_files("input[type='file']", caminho_pdf)
        else:
            return ERRO_ENVIO
    except Exception as e:
        print(f"[ERRO] Falha ao anexar o arquivo: {e}")
        return ERRO_ENVIO

    time.sleep(4)

    if upload_esta_desabilitado(page):
        return UPLOAD_DESABILITADO

    page.keyboard.type("T2")
    page.keyboard.press("Enter")
    return esperar_resposta_gpt(page)

# --- Streamlit Interface ---
st.title("Automatizador de PDFs para ChatGPT")

if st.button("Abrir Chrome com Debug"):
    abrir_chrome_com_debug()
    st.success("Chrome iniciado com depuração remota.")

uploaded_files = st.file_uploader("Envie os arquivos PDF", accept_multiple_files=True, type=["pdf"])

if st.button("Iniciar Processamento"):
    if not uploaded_files:
        st.warning("Envie pelo menos um arquivo PDF.")
    else:
        respostas = {}
        progress_bar = st.progress(0)
        log_area = st.empty()
        log_text = ""

        with sync_playwright() as p:
            try:
                browser = p.chromium.connect_over_cdp(CHROME_DEBUG_URL)
                context = browser.contexts[0]
                page = context.pages[0] if context.pages else context.new_page()
            except Exception as e:
                st.error(f"Erro ao conectar ao navegador Chrome com debug remoto: {e}")
                st.stop()

            for idx, uploaded_file in enumerate(uploaded_files):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getbuffer())
                    tmp_path = tmp.name

                log_text += f"Processando: {uploaded_file.name}\n"
                log_area.text(log_text)

                resposta = enviar_pdf_para_gpt(page, tmp_path)

                if resposta.startswith("__ERRO"):
                    log_text += f"Erro ao processar {uploaded_file.name}: {resposta}\n"
                else:
                    log_text += f"{uploaded_file.name} processado com sucesso.\n"
                log_area.text(log_text)

                respostas[uploaded_file.name] = resposta
                progress_bar.progress((idx + 1) / len(uploaded_files))

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            salvar_texto_docx(respostas, tmp_docx.name)
            st.success("Processamento finalizado. Baixe o resultado abaixo.")
            st.download_button(
                "Baixar arquivo .docx",
                data=open(tmp_docx.name, "rb").read(),
                file_name="texto_corrigido_final.docx"
            )