import os
import time
import re
from playwright.sync_api import sync_playwright
from tkinter import Tk, filedialog
from docx import Document
from tqdm import tqdm

# --- Funções Utilitárias ---
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
    erros_visuais = [
        "Tente novamente",
        "Too many requests",
        "aguarde alguns minutos",
        "erro ao carregar",
        "algo deu errado"
    ]
    for texto in erros_visuais:
        if page.locator(f"text={texto}").count() > 0:
            return True
    return False

def resposta_contem_erro(conteudo):
    erros = ["aguarde", "tente novamente", "espera", "erro", "carregar mais tarde"]
    return any(p in conteudo.lower() for p in erros)

def chat_esta_gerando(page):
    return page.locator("button:has(svg[aria-label='Stop generating'])").is_visible()

def ha_arquivo_anexado(page):
    return page.locator("div[role='listitem']").is_visible()

def esperar_resposta_gpt(page, tempo_maximo=180, intervalo_check=1.5, tempo_estavel=4):
    conteudo_anterior = ""
    tentativas_estaveis = 0
    tempo_decorrido = 0

    print("⏳ Aguardando resposta do GPT...")

    while tempo_decorrido < tempo_maximo:
        if page.url.endswith("/api/auth/error"):
            print("⚠️ Erro de autenticação detectado.")
            return "__ERRO_AUTENTICACAO__"

        if houve_erro_visual(page):
            print("❌ Erro detectado visualmente na interface.")
            return "__ERRO_GPT__"

        if chat_esta_gerando(page):
            print("⌛ ChatGPT ainda está gerando resposta...")
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
            print("✅ Resposta estabilizada.")
            if resposta_contem_erro(conteudo_atual):
                print("⚠️ Erro identificado no conteúdo da resposta do GPT.")
                return "__ERRO_GPT__"
            return conteudo_atual

        conteudo_anterior = conteudo_atual
        time.sleep(intervalo_check)
        tempo_decorrido += intervalo_check

    print("⚠️ Tempo máximo atingido. Retornando última resposta detectada.")
    if resposta_contem_erro(conteudo_anterior):
        print("⚠️ Erro identificado no conteúdo da resposta final.")
        return "__ERRO_GPT__"
    return conteudo_anterior

def enviar_pdf_para_gpt(page, caminho_pdf):
    print(f"\n📎 Enviando PDF: {caminho_pdf}")

    # Aguarda remoção de arquivos previamente anexados
    tentativas = 0
    while ha_arquivo_anexado(page) and tentativas < 30:
        print("⏳ Aguardando remoção de arquivo anterior antes de prosseguir...")
        time.sleep(1.5)
        tentativas += 1

    if ha_arquivo_anexado(page):
        print("⚠️ Arquivo ainda anexado após espera. Abortando envio para evitar duplicidade.")
        return "__ARQUIVO_JA_ANEXADO__"

    try:
        page.click("button:has(svg[aria-label='Upload a file'])", timeout=5000)
        time.sleep(1)
    except:
        print("⚠️ Não foi possível clicar no botão de upload. Prosseguindo mesmo assim.")

    page.set_input_files("input[type='file']", caminho_pdf)

    print("⌛ Aguardando reconhecimento do upload do PDF...")
    time.sleep(4)

    print("📝 Enviando comando 'T2'")
    page.keyboard.type("T2")
    page.keyboard.press("Enter")

    resposta = esperar_resposta_gpt(page)
    return resposta

# --- Execução principal ---
def main():
    pasta = selecionar_pasta()
    if not pasta:
        print("Nenhuma pasta selecionada. Encerrando.")
        return

    arquivos_pdf = get_pdf_files(pasta)
    if not arquivos_pdf:
        print("Nenhum PDF encontrado na pasta.")
        return

    arquivos_pdf = sorted(arquivos_pdf, key=lambda x: extrair_indice_final(os.path.basename(x)))

    respostas = {}
    print("\n📤 Iniciando envio automático dos PDFs...\n")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        page = browser.contexts[0].pages[0]

        for pdf in tqdm(arquivos_pdf, desc="Processando PDFs"):
            resposta = enviar_pdf_para_gpt(page, pdf)
            nome = os.path.basename(pdf)
            if resposta == "__ERRO_GPT__":
                respostas[nome] = "[ERRO] GPT não conseguiu processar corretamente este arquivo."
            elif resposta == "__ERRO_AUTENTICACAO__":
                respostas[nome] = "[ERRO] Erro de autenticação no ChatGPT."
            elif resposta == "__ARQUIVO_JA_ANEXADO__":
                respostas[nome] = "[ERRO] Arquivo anterior ainda anexado. Este foi ignorado para evitar envio duplo."
            else:
                respostas[nome] = resposta

    destino_docx = os.path.join(pasta, "texto_corrigido_final.docx")
    salvar_texto_docx(respostas, destino_docx)

    print(f"\n✅ Texto final salvo em: {destino_docx}")

if __name__ == "__main__":
    main()
