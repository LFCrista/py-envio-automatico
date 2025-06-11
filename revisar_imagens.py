import asyncio
import sys

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

import os
import time
import tempfile
import streamlit as st
from playwright.sync_api import sync_playwright
from docx import Document

# --- Configurações do Chrome Debug ---
CHROME_DEBUG_URL = "http://localhost:9222"

# --- Funções reutilizadas ---
def salvar_texto_docx(respostas_dict, destino):
    doc = Document()
    for nome_arquivo, texto in respostas_dict.items():
        doc.add_heading(f"Arquivo: {nome_arquivo}", level=1)
        doc.add_paragraph(texto)
        doc.add_paragraph("\n")
    doc.save(destino)

def esperar_resposta_gpt(page, tempo_maximo=180, intervalo_check=1.5, tempo_estavel=4):
    conteudo_anterior = ""
    tentativas_estaveis = 0
    tempo_decorrido = 0
    while tempo_decorrido < tempo_maximo:
        if page.url.endswith("/api/auth/error"):
            return "__ERRO_AUTENTICACAO__"
        erro_elements = page.locator(".text-token-text-error")
        if erro_elements.count() > 0:
            return "__ERRO_GPT__"
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
            return conteudo_atual if conteudo_atual else "__ERRO_GPT__"
        conteudo_anterior = conteudo_atual
        time.sleep(intervalo_check)
        tempo_decorrido += intervalo_check
    return conteudo_anterior if conteudo_anterior else "__ERRO_GPT__"

def enviar_pdf_para_gpt(page, caminho_pdf):
    tentativas = 0
    while page.locator("div[role='listitem']").is_visible() and tentativas < 30:
        time.sleep(1.5)
        tentativas += 1
    if page.locator("div[role='listitem']").is_visible():
        return "__ARQUIVO_JA_ANEXADO__"
    try:
        if page.locator("input[type='file']").count() > 0:
            page.set_input_files("input[type='file']", caminho_pdf)
        else:
            page.click("button:has(svg[aria-label='Upload a file'])", timeout=5000)
            page.wait_for_selector("input[type='file']", timeout=5000)
            page.set_input_files("input[type='file']", caminho_pdf)
    except:
        return "__ERRO_ENVIO__"
    time.sleep(4)
    page.keyboard.type("T2")
    page.keyboard.press("Enter")
    return esperar_resposta_gpt(page)

# --- Streamlit Interface ---
st.title("Automatizador de PDFs para ChatGPT")

uploaded_files = st.file_uploader("Faça upload dos arquivos PDF", accept_multiple_files=True, type=["pdf"])

if st.button("Iniciar Processamento"):
    if not uploaded_files:
        st.warning("Faça upload de pelo menos um arquivo PDF.")
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
            st.success("Processamento finalizado. Você pode baixar o arquivo abaixo.")
            st.download_button(
                "Baixar arquivo .docx",
                data=open(tmp_docx.name, "rb").read(),
                file_name="texto_corrigido_final.docx"
            )
