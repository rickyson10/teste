import os, re, json, zipfile, pdfplumber
from difflib import SequenceMatcher
from datetime import datetime, timedelta, date
from time import sleep
from playwright.sync_api import sync_playwright
import win32com.client as win32

# === CONFIGURAÇÕES ===
usuario = os.getlogin()
downloads_folder = rf"C:\Users\{usuario}\Downloads"
pdf_destino = downloads_folder

# === LISTA DE DATAS DE PROCESSAMENTO ===
listprocessamento = [
    "2026-01-01","2026-01-03","2026-01-06","2026-01-08","2026-01-11","2026-01-13","2026-01-16","2026-01-18",
    "2026-01-27","2026-01-29","2026-02-01","2026-02-03","2026-02-06","2026-02-08","2026-02-11","2026-02-13",

]
datas_processamento = [datetime.strptime(data_str, "%Y-%m-%d").date() for data_str in listprocessamento]
hoje = date.today() - timedelta(days=1)

# === VERIFICAÇÃO DE DATA ===
if hoje not in datas_processamento:
    print(f"❌ Data de processamento inválida: {hoje}")
    exit()

# === FUNÇÃO DE ENVIO DE EMAIL COM ANEXO E HTML ===
def enviar_email_html_com_anexo(assunto, corpo_html, caminho_pdf, referencia):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = "rickyson.almeida@midway.com.br; caio.requena@midway.com.br; Bmoreira@midway.com.br; geison.dias@midway.com.br; sergioc@midway.com.br; rodrigo.cano@midway.com.br;paulo.antonio@midway.com.br; amandavs@midway.com.br"
    mail.Subject = assunto
    mail.HTMLBody = corpo_html 
    if os.path.exists(caminho_pdf):
        mail.Attachments.Add(caminho_pdf)
    mail.Send()
    print(f"📧 Email enviado com anexo: {assunto}")

# === LOGIN E ACESSO AO PORTAL ===
with open(r"C:\Users\60004277\OneDrive - Lojas Riachuelo S.A\Área de Trabalho\py\site_digital.json", "r") as f:
    dados = json.load(f)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=dados)
    page = context.new_page()
    page.goto("https://portal.rchlo.digital/")
    sleep(7)
    page.get_by_role("link", name="Acessar").nth(3).click()
    sleep(8)
    page.get_by_role("link", name="Processo").click()
    sleep(8)
    page.get_by_role("button", name="󰉁 Serviços").click()
    sleep(8)
    page.locator("#SubClienteSeletor").select_option("RIACHUELO")
    print("📂 Selecionando cliente RIACHUELO...")
    page.get_by_role("row", name="EXTRATO 󰳟").locator("i").click()
    sleep(3)
    page.locator("#dtFinal").fill(str(hoje))
    sleep(3)
    page.locator("#dtInicio").fill(str(hoje))
    sleep(3)
    page.get_by_role("button", name="󰍉 Pesquisar").click()
    sleep(20)

    containers = page.locator("div[id^='ContainerProcesso_']")
    numero = None
    for i in range(containers.count()):
        bloco = containers.nth(i)
        if bloco.locator("div.col-md-4", has_text="SICC").count() > 0:
            id_attr = bloco.get_attribute("id")
            m = re.search(r"ContainerProcesso_(\d+)", id_attr)
            if m:
                numero = m.group(1)
                print(f"✔ Processo SICC: {numero}")
                break

    if not numero:
        print("❌ Nenhum processo SICC encontrado.")
        exit()

    print("💾 Acessando aba de modelos e iniciando download...")

    aba_modelos = page.locator(f"#TabsProcesso_{numero} >> text=Modelos")
    aba_modelos.wait_for(timeout=300000)
    aba_modelos.click()
    sleep(5)

    nome_arquivo_zip = f"Extrato_Modelos_{numero}.zip"
    botao_download = page.locator(f'button[data-nomearquivo="{nome_arquivo_zip}"][onclick="DownloadRelatorio(this)"]')
    botao_download.wait_for(timeout=300000)
    print("⬇️ Preparando para iniciar o download do relatório...")

    try:
        with page.expect_download(timeout=150000) as download_info:
            botao_download.click(force=True)
        download = download_info.value
        nome = download.suggested_filename
        caminho_down = os.path.join(downloads_folder, nome)
        download.save_as(caminho_down)
        print(f"✅ Download salvo com sucesso em: {caminho_down}")
    except Exception as e:
        print(f"❌ Erro ao baixar ou salvar o arquivo: {e}")
    browser.close()
    sleep(10)

    caminho_zip = os.path.join(downloads_folder, nome_arquivo_zip)
    if os.path.exists(caminho_zip):
        with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
            zip_ref.extractall(downloads_folder)
        print(f"📂 Arquivos extraídos com sucesso de {nome_arquivo_zip}")
    else:
        print(f"❌ Arquivo ZIP não encontrado: {caminho_zip}")

sleep(5)

# === TEXTOS DE REFERÊNCIA ===
texto_referencia_BD = """Parcelamento de Fatura Até 17,89 % a.m. 620,65 % a.a. 18,58 % a.m. 673,09 % a.a.
Parcelamento de Fatura Max. Próximo Mês 18,99 % a.m. 705,61 % a.a. 19,68 % a.m. 763,67 % a.a.
Multa Contratual Por Atraso 2,00 % a.m. 0,00 % a.a. 0,00 % a.m. 0,00 % a.a.
Juros Remuneratórios de Atraso (Pró-Rata) 19,99 % a.m. 790,72 % a.a. 20,71 % a.m. 856,58 % a.a.
Juros de Mora (Pró-Rata) 1,00 % a.m. 0,00 % a.a. 0,00 % a.m. 0,00 % a.a.
Ressarcimento de Cobrança R$ 6,50 R$ 0,00 0,00 % a.m. 0,00 % a.a.
Saque (Pró-Rata) 17,49 % a.m. 591,85 % a.a. 18,21 % a.m. 644,13 % a.a.
Juros Remuneratórios de Atraso Max. Próximo Mês (Pró-Rat 20,99 % a.m. 884,00 % a.a. 21,71 % a.m. 956,13 % a.a.
IOF: 0,0082% ao dia + 0,38% 0,00 % a.m. 0,00 % a.a. 0,00 % a.m. 0,00 % a.a.
Saque Max. Próximo Mês (Pró-Rata) 17,49 % a.m. 591,85 % a.a. 18,21 % a.m. 644,13 % a.a.
Tarifa Saque R$ 18,90 R$ 0,00 0,00 % a.m. 0,00 % a.a.
Tarifa Aval. Emerg. de Crédito (Limite Emergencial) R$ 17,90 R$ 0,00 0,00 % a.m. 0,00 % a.a.""".splitlines()

texto_referencia_PL = """Parcelamento de Fatura Até 17,89 % a.m. 620,65 % a.a. 18,58 % a.m. 673,09 % a.a.
Parcelamento de Fatura Max. Próximo Mês 18,99 % a.m. 705,61 % a.a. 19,68 % a.m. 763,67 % a.a.
Multa Contratual Por Atraso 2,00 % a.m. 0,00 % a.a. 0,00 % a.m. 0,00 % a.a.
Juros Remuneratórios de Atraso (Pró-Rata) 19,99 % a.m. 790,72 % a.a. 20,71 % a.m. 856,58 % a.a.
Juros de Mora (Pró-Rata) 1,00 % a.m. 0,00 % a.a. 0,00 % a.m. 0,00 % a.a.
Ressarcimento de Cobrança R$ 6,50 R$ 0,00 0,00 % a.m. 0,00 % a.a.
Juros Remuneratórios de Atraso Max. Próximo Mês (Pró-Rat 20,99 % a.m. 884,00 % a.a. 21,71 % a.m. 956,13 % a.a.
IOF: 0,0082% ao dia + 0,38% 0,00 % a.m. 0,00 % a.a. 0,00 % a.m. 0,00 % a.a.""".splitlines()

# === VALIDAÇÃO DOS PDFs ===
arquivos = [
    {"path": os.path.join(downloads_folder, "MAS_MIDWAY_PIX.pdf"), "tipo": "BD", "index_offset": 13, "referencia": texto_referencia_BD},
    {"path": os.path.join(downloads_folder, "RCH_MIDWAY_PIX.pdf"), "tipo": "PL", "index_offset": 9, "referencia": texto_referencia_PL}
]

def destacar_diferencas_html(linha1, linha2):
    matcher = SequenceMatcher(None, linha1, linha2)
    resultado1, resultado2 = "", ""
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            resultado1 += linha1[i1:i2]
            resultado2 += linha2[j1:j2]
        elif tag == "replace":
            resultado1 += f"<span style='color:red;font-weight:bold;'>{linha1[i1:i2]}</span>"
            resultado2 += f"<span style='color:blue;font-weight:bold;'>{linha2[j1:j2]}</span>"
        elif tag == "delete":
            resultado1 += f"<span style='color:red;font-weight:bold;'>{linha1[i1:i2]}</span>"
        elif tag == "insert":
            resultado2 += f"<span style='color:blue;font-weight:bold;'>{linha2[j1:j2]}</span>"
    return resultado1, resultado2

for arquivo in arquivos:
    if not os.path.exists(arquivo["path"]):
        print(f"⚠️ Arquivo não encontrado: {arquivo['path']}")
        continue

    with pdfplumber.open(arquivo["path"]) as pdf:
        for page in pdf.pages[:2]:
            text = page.extract_text()
            if "Encargo" in text or "Encargos" in text:
                linhas = text.splitlines()
                index = next((i for i, linha in enumerate(linhas) if "Descrição" in linha and "Taxa de Juros" in linha), None)
                if index is not None:
                    tabela = linhas[index + 1:index + arquivo["index_offset"]]
                    todas_corretas = True
                    corpo_html = f"""
<html><body style="font-family: Arial; font-size: 14px; color: #333;">
<p>Prezados,</p>
<p>Durante o processo automatizado de <strong>validação de Taxas</strong> dos clientes <strong>{arquivo['tipo']}</strong>, foram identificadas divergências em relação às referências oficiais.</p>
<ul>
<li>❌ <strong>Status da execução:</strong> Finalizado com inconsistências</li>
<li>📁 <strong>Data de processamento:</strong> {hoje.strftime('%d/%m/%Y')}</li>
</ul>
<p><strong>Referência utilizada:</strong></p>
<pre>{chr(10).join(arquivo["referencia"])}</pre>
<p>Abaixo seguem os detalhes das divergências encontradas:</p>
"""
                    for i, (linha_extraida, linha_referencia) in enumerate(zip(tabela, arquivo["referencia"])):
                        if linha_extraida.strip() != linha_referencia.strip():
                            todas_corretas = False
                            destaque1, destaque2 = destacar_diferencas_html(linha_extraida.strip(), linha_referencia.strip())
                            corpo_html += f"""
<p><strong>❌ Divergência na linha {i + 1}:</strong><br>
• Extraído da Fatura: {destaque1}<br>
• Referência esperada: {destaque2}</p>
"""
                    if todas_corretas:
                        corpo_html = f"""
<html><body style="font-family: Arial; font-size: 14px; color: #333;">
<p>Prezados,</p>
<p>Informo que o processo automatizado de <strong>validação de Taxas</strong> dos clientes <strong>{arquivo['tipo']}</strong> foi concluído com êxito.</p>
<ul>
<li>✅ <strong>Status da execução:</strong> Finalizado com sucesso</li>
<li>📁 <strong>Data de processamento:</strong> {hoje.strftime('%d/%m/%Y')}</li>
</ul>
<p><strong>Referência utilizada:</strong></p>
<pre>{chr(10).join(arquivo["referencia"])}</pre>
<p>Atenciosamente,<br><strong>Time Analytics</strong><br>Operações - Midway</p>
</body></html>
"""
                        enviar_email_html_com_anexo(f"Validação de Taxas - {arquivo['tipo']} Conforme", corpo_html, arquivo["path"], arquivo["referencia"])
                    else:
                        corpo_html += """
<p>Solicitamos atenção para os pontos destacados. Caso necessário, estamos à disposição para reprocessamento ou esclarecimentos.</p>
<p>Atenciosamente,<br><strong>Time Analytics</strong><br>Operações - Midway</p>
</body></html>
"""
                        enviar_email_html_com_anexo(f"[ATENÇÃO] Validação de Taxas - {arquivo['tipo']} com Divergências", corpo_html, arquivo["path"], arquivo["referencia"])
