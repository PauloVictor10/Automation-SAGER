import io
import os
import json
import time
import smtplib
import requests
from datetime import datetime, timedelta
from PyPDF2 import PdfReader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from urllib.parse import quote_plus, urljoin
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ================== CONFIGURAÇÕES ==================
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USUARIO = 'aplan.notificacoes@gmail.com'
EMAIL_SENHA = 'dwbu nxga jnjt riwj'
DESTINATARIOS = ['paulovictormcarneiro@gmail.com']
NOME_REMETENTE = "A-Plan Notificação"

URL_BASE = "https://sintegre.ons.org.br"
URL = "https://sintegre.ons.org.br/sites/2/14/Paginas/servicos/historico-de-produtos.aspx?produto=RO%20(Relatório%20de%20Análise%20de%20Ocorrência)"
USERNAME = "paulo.victor@casadosventos.com.br"
PASSWORD = "windpower07#"

USUARIO_XPATH = '/html/body/div/div[2]/div/div/div[1]/div/form/div[1]/input'
SENHA_XPATH   = '/html/body/div/div[2]/div/div/div[1]/div/form/div[2]/input'
ENTRAR_XPATH  = '/html/body/div/div[2]/div/div/div[1]/div/form/div[4]/input[2]'

XPATH_BASE        = '//div[contains(@class,"item_produto_")]'
XPATH_SUFIXO_DATA = './/small[contains(text(),"Publicado:")]'

HISTORICO_ARQUIVO  = 'historico.json'
SHEETS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxq9PvbyeOgbfSDLQd9ECblGvlPjMhiQAF8wXPC-7rNAQaVxsT2oxbS92qWFKXFzK425g/exec"

# ================== FUNÇÕES UTILITÁRIAS ==================
def carregar_historico():
    if not os.path.exists(HISTORICO_ARQUIVO):
        return []
    try:
        with open(HISTORICO_ARQUIVO, 'r', encoding='utf-8') as f:
            return json.load(f) or []
    except:
        return []

def salvar_historico(h):
    with open(HISTORICO_ARQUIVO, 'w', encoding='utf-8') as f:
        json.dump(h, f, indent=2, ensure_ascii=False)

def extrair_distribuicao(pdf_bytes):
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        texto = reader.pages[0].extract_text() or ""
    except:
        return "Não informado"
    for linha in texto.split("\n"):
        if linha.strip().startswith("Distribuição"):
            return linha.split(":", 1)[-1].strip()
    return "Não informado"

def extrair_texto(pdf_bytes, palavra_chave="", num_caracteres=20, modo="simples"):
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        texto = reader.pages[0].extract_text() or ""
    except:
        return "Não encontrado"

    if modo == "simples":
        for linha in texto.split("\n"):
            if palavra_chave in linha:
                i = linha.find(palavra_chave)
                return linha[i + len(palavra_chave): i + len(palavra_chave) + num_caracteres].strip()
        return "Não encontrado"

    if modo == "descricao":
        linhas = texto.split("\n")
        for i, l in enumerate(linhas):
            if "Descrição" in l:
                trecho = " ".join(linhas[max(0, i - 3):i]).strip()
                return trecho[:500] if trecho else "Não encontrado"

    return "Não encontrado"

def gerar_link_assinatura(numero_ro):
    FORM_BASE_URL = "https://docs.google.com/forms/d/e/1FAIpQLSctOQ-WW_GW-2YdIz4d0O3FbxHP5eV4adET2hOZ_Rm_qTKCsg/viewform"
    FORM_RO_FIELD = "entry.1445388733"
    return f"{FORM_BASE_URL}?usp=pp_url&{FORM_RO_FIELD}={quote_plus(numero_ro)}"

def montar_url_absoluta(link):
    """Garante que o link do PDF seja sempre uma URL absoluta."""
    if not link:
        return None
    if link.startswith("http://") or link.startswith("https://"):
        return link
    return urljoin(URL_BASE, link)

# ================== EMAIL ==================
def enviar_email(pdf_bytes, numero_ro, descricao, distribuicao, data_pub):
    link_assinatura = gerar_link_assinatura(numero_ro)
    ano_atual = datetime.today().year

    msg = MIMEMultipart('alternative')
    msg['From']    = f"{NOME_REMETENTE} <{EMAIL_USUARIO}>"
    msg['To']      = ", ".join(DESTINATARIOS)
    msg['Subject'] = f"[RO ONS] {numero_ro} — Publicado em {data_pub}"

    corpo_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Relatório de Ocorrência ONS</title>
</head>
<body style="margin:0;padding:0;background-color:#ECEEF1;
             font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">

  <!-- Wrapper externo -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0"
         style="background-color:#ECEEF1;padding:40px 16px;">
    <tr>
      <td align="center">

        <!-- Card principal -->
        <table width="620" cellpadding="0" cellspacing="0" border="0"
               style="max-width:620px;width:100%;background-color:#ffffff;
                      border-radius:12px;overflow:hidden;
                      box-shadow:0 4px 24px rgba(0,0,0,0.10);">

          <!-- ══ HEADER ══ -->
          <tr>
            <td style="background:linear-gradient(135deg,#0A2540 0%,#1A3F6F 100%);
                       padding:36px 40px 28px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td>
                    <!-- Badge -->
                    <table cellpadding="0" cellspacing="0" border="0"
                           style="margin-bottom:14px;">
                      <tr>
                        <td style="background:rgba(255,255,255,0.12);
                                   border:1px solid rgba(255,255,255,0.25);
                                   border-radius:20px;padding:4px 14px;">
                          <span style="color:#93C5FD;font-size:11px;font-weight:700;
                                       letter-spacing:1.2px;text-transform:uppercase;">
                            ONS &middot; Relatório de Ocorrência
                          </span>
                        </td>
                      </tr>
                    </table>
                    <!-- Número RO -->
                    <div style="color:#FFFFFF;font-size:22px;font-weight:700;
                                line-height:1.3;margin-bottom:6px;">
                      {numero_ro}
                    </div>
                    <!-- Data -->
                    <div style="color:#93C5FD;font-size:13px;">
                      Publicado em {data_pub}
                    </div>
                  </td>
                  <!-- Ícone -->
                  <td width="60" valign="top" align="right">
                    <table cellpadding="0" cellspacing="0" border="0">
                      <tr>
                        <td align="center" width="52" height="52"
                            style="background:rgba(255,255,255,0.10);
                                   border-radius:50%;font-size:24px;
                                   line-height:52px;">
                          &#x1F4CB;
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ══ SAUDAÇÃO ══ -->
          <tr>
            <td style="padding:32px 40px 0;">
              <p style="margin:0 0 12px;color:#374151;font-size:15px;line-height:1.7;">
                Prezados(as),
              </p>
              <p style="margin:0;color:#374151;font-size:15px;line-height:1.7;">
                Encaminhamos abaixo e em anexo o
                <strong>Relatório de Análise de Ocorrência (RO)</strong>
                recém-publicado pelo <strong>ONS</strong>. Verifique os detalhes
                e confirme a leitura por meio do botão de assinatura eletrônica.
              </p>
            </td>
          </tr>

          <!-- ══ DIVISOR GRADIENTE ══ -->
          <tr>
            <td style="padding:24px 40px 0;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td height="1"
                      style="background:linear-gradient(90deg,#E5E7EB 0%,#93C5FD 50%,#E5E7EB 100%);
                             font-size:0;line-height:0;">&nbsp;</td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ══ DETALHES ══ -->
          <tr>
            <td style="padding:24px 40px 0;">

              <!-- Rótulo seção -->
              <p style="margin:0 0 14px;color:#0A2540;font-size:12px;font-weight:700;
                         letter-spacing:1px;text-transform:uppercase;">
                Detalhes do documento
              </p>

              <!-- Número do RO (destaque azul) -->
              <table width="100%" cellpadding="0" cellspacing="0" border="0"
                     style="background:#EFF6FF;border-radius:8px;
                            margin-bottom:10px;">
                <tr>
                  <td style="padding:14px 18px;">
                    <div style="color:#6B7280;font-size:11px;font-weight:600;
                                text-transform:uppercase;letter-spacing:0.7px;
                                margin-bottom:5px;">
                      Número do RO
                    </div>
                    <div style="color:#0A2540;font-size:17px;font-weight:700;">
                      {numero_ro}
                    </div>
                  </td>
                </tr>
              </table>

              <!-- Distribuição -->
              <table width="100%" cellpadding="0" cellspacing="0" border="0"
                     style="background:#F9FAFB;border:1px solid #E5E7EB;
                            border-radius:8px;margin-bottom:10px;">
                <tr>
                  <td style="padding:14px 18px;">
                    <div style="color:#6B7280;font-size:11px;font-weight:600;
                                text-transform:uppercase;letter-spacing:0.7px;
                                margin-bottom:5px;">
                      Distribuição
                    </div>
                    <div style="color:#1F2937;font-size:14px;line-height:1.5;">
                      {distribuicao}
                    </div>
                  </td>
                </tr>
              </table>

              <!-- Descrição -->
              <table width="100%" cellpadding="0" cellspacing="0" border="0"
                     style="background:#F9FAFB;border:1px solid #E5E7EB;
                            border-radius:8px;">
                <tr>
                  <td style="padding:14px 18px;">
                    <div style="color:#6B7280;font-size:11px;font-weight:600;
                                text-transform:uppercase;letter-spacing:0.7px;
                                margin-bottom:6px;">
                      Descrição
                    </div>
                    <div style="color:#1F2937;font-size:14px;line-height:1.6;">
                      {descricao}
                    </div>
                  </td>
                </tr>
              </table>

            </td>
          </tr>

          <!-- ══ AVISO ══ -->
          <tr>
            <td style="padding:18px 40px 0;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0"
                     style="background:#FFFBEB;border-left:4px solid #F59E0B;
                            border-radius:0 8px 8px 0;">
                <tr>
                  <td style="padding:13px 16px;color:#92400E;
                             font-size:13px;line-height:1.5;">
                    &#x26A0;&#xFE0F; &nbsp;
                    <strong>Ação necessária:</strong>
                    Este documento requer confirmação de leitura.
                    Utilize o botão abaixo para registrar sua assinatura eletrônica.
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ══ BOTÃO CTA ══ -->
          <tr>
            <td style="padding:32px 40px;" align="center">
              <table cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td align="center"
                      style="background:linear-gradient(135deg,#1D4ED8 0%,#2563EB 100%);
                             border-radius:8px;">
                    <a href="{link_assinatura}" target="_blank"
                       style="display:inline-block;color:#ffffff;font-size:15px;
                              font-weight:700;text-decoration:none;
                              padding:16px 44px;letter-spacing:0.3px;">
                      &#x270D;&#xFE0F; &nbsp; Confirmar leitura e assinar
                    </a>
                  </td>
                </tr>
              </table>
              <p style="margin:12px 0 0;color:#9CA3AF;font-size:12px;">
                O link abre um formulário seguro do Google Forms
              </p>
            </td>
          </tr>

          <!-- ══ DIVISOR SIMPLES ══ -->
          <tr>
            <td style="padding:0 40px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td height="1"
                      style="background:#E5E7EB;font-size:0;line-height:0;">&nbsp;</td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ══ ASSINATURA ══ -->
          <tr>
            <td style="padding:24px 40px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td>
                    <p style="margin:0;color:#374151;font-size:14px;line-height:1.7;">
                      Atenciosamente,<br/>
                      <strong style="color:#0A2540;">
                        COG — Centro de Operação e Gestão
                      </strong><br/>
                      <span style="color:#6B7280;font-size:12px;">
                        A-Plan Notificação Automática
                      </span>
                    </p>
                  </td>
                  <td align="right" valign="bottom">
                    <span style="color:#D1D5DB;font-size:11px;">
                      {data_pub} &middot; {ano_atual}
                    </span>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ══ RODAPÉ LEGAL ══ -->
          <tr>
            <td style="background:#F3F4F6;padding:14px 40px;
                       border-top:1px solid #E5E7EB;">
              <p style="margin:0;color:#9CA3AF;font-size:11px;
                         line-height:1.6;text-align:center;">
                Este é um e-mail automático gerado pelo sistema A-Plan.
                Não responda a esta mensagem.<br/>
                O PDF completo do relatório está disponível em anexo.
              </p>
            </td>
          </tr>

        </table>
        <!-- fim card -->

      </td>
    </tr>
  </table>

</body>
</html>"""

    msg.attach(MIMEText(corpo_html, 'html', 'utf-8'))

    # Anexo PDF
    parte = MIMEBase('application', 'pdf')
    parte.set_payload(pdf_bytes)
    encoders.encode_base64(parte)
    nome_arquivo = f"RO_{numero_ro.replace('/', '-').replace(' ', '_')}.pdf"
    parte.add_header('Content-Disposition', f'attachment; filename="{nome_arquivo}"')
    msg.attach(parte)

    with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_USUARIO, EMAIL_SENHA)
        smtp.sendmail(EMAIL_USUARIO, DESTINATARIOS, msg.as_string())

# ================== SHEETS / DRIVE ==================
def registrar_no_sheets(pdf_bytes, numero_ro, descricao, distribuicao, data_pub):
    import base64
    payload = {
        "numero_ro":       numero_ro,
        "descricao":       descricao,
        "distribuicao":    distribuicao,
        "data_publicacao": data_pub,
        "pdf_base64":      base64.b64encode(pdf_bytes).decode("utf-8"),
    }
    try:
        r = requests.post(SHEETS_WEBHOOK_URL, json=payload, timeout=30)
        if r.status_code == 200:
            print("📁 PDF salvo no Drive e registrado no Sheets.")
        else:
            print(f"⚠️ Sheets respondeu com status {r.status_code}")
    except Exception as e:
        print(f"⚠️ Falha ao enviar ao Sheets/Drive: {e}")

# ================== LOGIN ==================
def realizar_login(page, context):
    print("🔐 Realizando login...")
    try:
        page.goto(URL_BASE, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)

        page.wait_for_selector(f'xpath={USUARIO_XPATH}', timeout=30000)

        page.fill(f'xpath={USUARIO_XPATH}', USERNAME)
        page.wait_for_timeout(2000)
        page.fill(f'xpath={SENHA_XPATH}', PASSWORD)
        page.wait_for_timeout(2000)
        page.click(f'xpath={ENTRAR_XPATH}')

        page.wait_for_timeout(15000)

        url_atual = page.url
        print(f"  URL pós-login: {url_atual}")

        if "auth" not in url_atual.lower() and "login" not in url_atual.lower():
            context.storage_state(path="storage_state.json")
            print("✅ Login realizado e sessão salva")
            return True
        else:
            print("❌ Login pode ter falhado")
            return False

    except Exception as e:
        print(f"❌ Erro no login: {e}")
        return False

# ================== ACESSO À PÁGINA ==================
def acessar_pagina_ros(page, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            print(f"🌐 Tentativa {tentativa + 1}/{max_tentativas} - Acessando ROs...")

            page.goto(URL, wait_until="networkidle", timeout=120000)

            print("  ⏳ Aguardando carregamento inicial...")
            page.wait_for_timeout(10000)

            print("  🔍 Procurando lista de ROs...")

            if page.locator(f"xpath={XPATH_BASE}").count() > 0:
                print("  ✅ Lista encontrada (método XPath)")
                return True

            if page.locator('div[class*="item_produto"]').count() > 0:
                print("  ✅ Lista encontrada (método CSS)")
                return True

            print("  ⏳ Aguardando carregamento adicional...")
            page.wait_for_timeout(20000)

            if page.locator(f"xpath={XPATH_BASE}").count() > 0:
                print("  ✅ Lista encontrada após espera adicional")
                return True

            print(f"  ⚠️ Lista não encontrada na tentativa {tentativa + 1}")

            if tentativa < max_tentativas - 1:
                espera = 45 * (tentativa + 1)
                print(f"  ⏱️  Aguardando {espera}s antes de tentar novamente...")
                time.sleep(espera)

        except PlaywrightTimeoutError:
            print(f"  ⏱️ Timeout na tentativa {tentativa + 1}")
            if tentativa < max_tentativas - 1:
                time.sleep(60)
        except Exception as e:
            print(f"  ❌ Erro: {e}")
            if tentativa < max_tentativas - 1:
                time.sleep(60)

    return False

# ================== MAIN ==================
def main():
    ros_ja_enviados = carregar_historico()

    print("\n" + "=" * 60)
    print("🤖 INICIANDO COLETA DE ROs - ONS")
    print("=" * 60 + "\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
                "--disable-gpu",
                "--disable-software-rasterizer",
                "--disable-extensions",
                "--no-first-run",
                "--no-default-browser-check",
            ],
        )

        storage_file  = "storage_state.json"
        storage_state = storage_file if os.path.exists(storage_file) else None

        if storage_state:
            print("📂 Sessão encontrada, tentando usar...")
        else:
            print("🆕 Primeira execução, será necessário fazer login")

        context = browser.new_context(
            storage_state=storage_state,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1920, "height": 1080},
            locale="pt-BR",
            ignore_https_errors=True,
        )

        page = context.new_page()

        try:
            # Tenta acessar diretamente
            sucesso = acessar_pagina_ros(page)

            # Se falhar, faz login e tenta de novo
            if not sucesso:
                print("\n⚠️ Acesso direto falhou, tentando fazer login...")
                if realizar_login(page, context):
                    print("✅ Login realizado, tentando acessar novamente...")
                    sucesso = acessar_pagina_ros(page)
                else:
                    print("❌ Não foi possível fazer login")
                    return

            if not sucesso:
                print("\n❌ FALHA: Não foi possível acessar a lista de ROs")
                print("💡 Sugestões:")
                print("   1. Execute o script localmente")
                print("   2. Verifique se o IP do VPS não está bloqueado")
                print("   3. Considere usar um proxy/VPN no VPS")
                return

            print("\n✅ Página carregada com sucesso!\n")

            # Aceita cookies se aparecer
            try:
                btn_cookie = page.locator('//button[contains(text(),"Concordo")]')
                if btn_cookie.count() > 0:
                    btn_cookie.first.click()
                    page.wait_for_timeout(2000)
                    print("🍪 Cookies aceitos")
            except:
                pass

            # Datas válidas para busca (hoje e ontem)
            DIAS_PARA_BUSCAR = 1
            datas_validas = [
                (datetime.today() - timedelta(days=i)).strftime("%d/%m/%Y")
                for i in range(DIAS_PARA_BUSCAR + 1)
            ]
            print(f"📅 Buscando ROs publicados em: {', '.join(datas_validas)}\n")

            # Aguarda e coleta itens
            page.wait_for_selector(f"xpath={XPATH_BASE}", timeout=30000)
            itens       = page.locator(f"xpath={XPATH_BASE}")
            total_itens = itens.count()

            print(f"📋 {total_itens} ROs encontrados na página\n")

            enviados    = 0
            processados = 0

            for idx in range(total_itens):
                item = itens.nth(idx)
                processados += 1

                try:
                    # ── Extrai data ───────────────────────────────────────
                    data_pub = None

                    data_locator = item.locator(f"xpath={XPATH_SUFIXO_DATA}")
                    if data_locator.count() > 0:
                        data_txt = data_locator.first.inner_text()
                        data_pub = data_txt.replace("Publicado:", "").strip().split()[0]

                    if not data_pub:
                        small_tags = item.locator("small")
                        for i in range(small_tags.count()):
                            texto = small_tags.nth(i).inner_text()
                            if "Publicado" in texto or "/" in texto:
                                data_pub = texto.replace("Publicado:", "").strip().split()[0]
                                break

                    if not data_pub:
                        print(f"[{processados}/{total_itens}] ⚠️ Sem data, pulando...")
                        continue

                    if data_pub not in datas_validas:
                        print(f"[{processados}/{total_itens}] ⛔ Data {data_pub} fora do período. Encerrando busca.")
                        break

                    # ── Extrai link PDF ───────────────────────────────────
                    link_pdf = None

                    link_locator = item.get_by_role("link", name="Baixar")
                    if link_locator.count() > 0:
                        link_pdf = link_locator.first.get_attribute("href")

                    if not link_pdf:
                        all_links = item.locator("a[href*='/Produtos/']")
                        if all_links.count() > 0:
                            link_pdf = all_links.first.get_attribute("href")

                    if not link_pdf:
                        print(f"[{processados}/{total_itens}] ⚠️ Sem link de download, pulando...")
                        continue

                    # ── CORREÇÃO: garante URL absoluta ────────────────────
                    link_pdf = montar_url_absoluta(link_pdf)
                    if not link_pdf:
                        print(f"[{processados}/{total_itens}] ⚠️ Link inválido após normalização, pulando...")
                        continue

                    # ── Download do PDF com cookies da sessão ─────────────
                    session = requests.Session()
                    for cookie in context.cookies():
                        session.cookies.set(cookie["name"], cookie["value"])

                    print(f"[{processados}/{total_itens}] 📥 Baixando PDF ({data_pub})...")
                    r = session.get(link_pdf, timeout=60)

                    if r.status_code != 200 or b"%PDF" not in r.content[:10]:
                        print(f"[{processados}/{total_itens}] ❌ PDF inválido (status {r.status_code})")
                        continue

                    # ── Extrai informações do PDF ─────────────────────────
                    numero_ro    = extrair_texto(r.content, "N. º:", 20, "simples")
                    descricao    = extrair_texto(r.content, "", 0, "descricao")
                    distribuicao = extrair_distribuicao(r.content)

                    identificador = f"{numero_ro}|{data_pub}"

                    if identificador in ros_ja_enviados:
                        print(f"[{processados}/{total_itens}] 🔁 {numero_ro} já enviado anteriormente")
                        continue

                    # ── Envia e-mail e registra no Sheets ─────────────────
                    print(f"[{processados}/{total_itens}] 📧 Enviando {numero_ro}...")
                    enviar_email(r.content, numero_ro, descricao, distribuicao, data_pub)
                    registrar_no_sheets(r.content, numero_ro, descricao, distribuicao, data_pub)

                    ros_ja_enviados.append(identificador)
                    enviados += 1
                    print(f"[{processados}/{total_itens}] ✅ {numero_ro} enviado com sucesso\n")

                except Exception as e:
                    print(f"[{processados}/{total_itens}] ❌ Erro: {e}\n")
                    continue

            print("\n" + "=" * 60)
            print("✅ PROCESSO CONCLUÍDO")
            print(f"   📊 ROs processados : {processados}")
            print(f"   📨 ROs enviados    : {enviados}")
            print("=" * 60 + "\n")

        except Exception as e:
            print(f"\n❌ ERRO CRÍTICO: {e}\n")

        finally:
            salvar_historico(ros_ja_enviados)
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
