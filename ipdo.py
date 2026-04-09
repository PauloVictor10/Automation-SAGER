import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
from pypdf import PdfReader
import re
from io import BytesIO
import base64

DRIVE_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxuLH_wgaEojh1fqCWETqACH4rUjkStW1GVLhX_J1eJ9t1hTJo-YNVv3y9_IjqgyP1v/exec"

# ------------------------------
# NORMALIZA TEXTO DO PDF
# ------------------------------
def normalizar_texto_pdf(txt):
    txt = txt.replace("\xad", "").replace("\u200b", "").replace("\ufeff", "")
    txt = re.sub(r"[ \t]+", " ", txt)
    return txt.strip()

# ------------------------------
# HELPER: extrai bloco de texto entre dois padrões
# ------------------------------
def extrair_bloco(texto, inicio_pattern, fim_patterns):
    m = re.search(inicio_pattern, texto, flags=re.IGNORECASE)
    if not m:
        return None
    trecho = texto[m.end():]
    fim_re = "|".join(fim_patterns)
    m_fim = re.search(fim_re, trecho, flags=re.IGNORECASE | re.DOTALL)
    if m_fim:
        trecho = trecho[:m_fim.start()]
    trecho = re.sub(r"\n+", " ", trecho)
    trecho = re.sub(r"[ \t]+", " ", trecho).strip()
    return trecho if len(trecho) > 3 else None

# ------------------------------
# HELPER: converte texto em itens de lista HTML
# ------------------------------
def texto_para_itens_html(texto):
    raw_itens = re.split(r"(?<=\.)\s+(?=[A-ZÁÉÍÓÚÂÊÔÃÕÇ])", texto)
    itens = [i.strip(" .\n") for i in raw_itens if len(i.strip()) > 5]
    if not itens:
        return "<p style='margin:0; font-size:13px; color:#666; font-style:italic;'>Nada a relatar.</p>"
    html = "<ul style='margin:0; padding-left:16px; font-size:13px; line-height:1.8;'>\n"
    for item in itens:
        frase = item[0].upper() + item[1:]
        if not frase.endswith("."):
            frase += "."
        html += f"  <li style='margin-bottom:4px;'>{frase}</li>\n"
    html += "</ul>"
    return html

# ------------------------------
# HELPER: bloco de seção com novo layout
# ------------------------------
def bloco_secao(titulo, conteudo_html, cor_fundo, cor_barra, cor_titulo):
    return f"""
    <div style="margin-bottom:16px; border-radius:8px; border:1px solid #e0e0e0; overflow:hidden;">
      <div style="background:{cor_fundo}; padding:10px 16px; display:flex; align-items:center; gap:8px;">
        <div style="width:4px; height:18px; background:{cor_barra}; border-radius:2px; flex-shrink:0;"></div>
        <span style="font-size:13px; font-weight:600; color:{cor_titulo};">{titulo}</span>
      </div>
      <div style="padding:12px 16px 12px 20px; background:#ffffff;">
        {conteudo_html}
      </div>
    </div>
    """

# ------------------------------
# EXTRAI TODAS AS SEÇÕES DO PDF
# ------------------------------
def extrair_secoes_formatado(pdf_bytes):
    reader = PdfReader(BytesIO(pdf_bytes))

    texto_completo = ""
    for page in reader.pages:
        txt = page.extract_text() or ""
        texto_completo += "\n" + txt

    texto_completo = normalizar_texto_pdf(texto_completo)

    PADROES_FIM = [
        r"OPERADOR\s+NACIONAL",
        r"ONS\s*/",
        r"Página\s+\d+",
        r"\Z"
    ]

    secoes_html = ""

    # ── 1. Destaques Submercado Nordeste ──────────────────────────────
    m4 = re.search(r"4\s*[-–]\s*Destaques da Operação", texto_completo, flags=re.IGNORECASE)
    if m4:
        bloco4 = texto_completo[m4.start():]
        texto_nord = extrair_bloco(
            bloco4,
            r"Submercado\s+Nordeste\s*:",
            [r"Submercado\s+(?!Nordeste)\w+"] + PADROES_FIM
        )
        conteudo = texto_para_itens_html(texto_nord) if texto_nord else "<p style='margin:0; font-size:13px; color:#666; font-style:italic;'>Nada a relatar.</p>"
        secoes_html += bloco_secao(
            "Destaques – Submercado Nordeste",
            conteudo,
            cor_fundo="#E6F1FB",
            cor_barra="#185FA5",
            cor_titulo="#0C447C"
        )

    # ── 2. Intercâmbio de Energia do Submercado Nordeste ─────────────
    texto_intercambio = extrair_bloco(
        texto_completo,
        r"Intercâmbio de Energia do Submercado Nordeste",
        [r"Intercâmbio de Energia do Submercado Norte",
         r"Intercâmbio Internacional",
         r"OCORRÊNCIAS"] + PADROES_FIM
    )
    conteudo = texto_para_itens_html(texto_intercambio) if texto_intercambio else "<p style='margin:0; font-size:13px; color:#666; font-style:italic;'>Nada a relatar.</p>"
    secoes_html += bloco_secao(
        "Intercâmbio de Energia – Submercado Nordeste",
        conteudo,
        cor_fundo="#FAEEDA",
        cor_barra="#BA7517",
        cor_titulo="#633806"
    )

    # ── 3. Ocorrências na Rede de Operação ───────────────────────────
    texto_op = extrair_bloco(
        texto_completo,
        r"OCORRÊNCIAS NA REDE DE OPERAÇÃO",
        [r"OCORRÊNCIAS NA REDE DE DISTRIBUIÇÃO",
         r"INTEGRAÇÃO DE NOVAS"] + PADROES_FIM
    )
    conteudo = texto_para_itens_html(texto_op) if texto_op else "<p style='margin:0; font-size:13px; color:#666; font-style:italic;'>Nada a relatar.</p>"
    secoes_html += bloco_secao(
        "Ocorrências na Rede de Operação",
        conteudo,
        cor_fundo="#FCEBEB",
        cor_barra="#A32D2D",
        cor_titulo="#791F1F"
    )

    # ── 4. Ocorrências na Rede de Distribuição ───────────────────────
    texto_dist = extrair_bloco(
        texto_completo,
        r"OCORRÊNCIAS NA REDE DE DISTRIBUIÇÃO",
        [r"INTEGRAÇÃO DE NOVAS",
         r"INFORMAÇÕES ADICIONAIS"] + PADROES_FIM
    )
    conteudo = texto_para_itens_html(texto_dist) if texto_dist else "<p style='margin:0; font-size:13px; color:#666; font-style:italic;'>Nada a relatar.</p>"
    secoes_html += bloco_secao(
        "Ocorrências na Rede de Distribuição",
        conteudo,
        cor_fundo="#FCEBEB",
        cor_barra="#A32D2D",
        cor_titulo="#791F1F"
    )

    # ── 5. Integração de Novas Instalações ───────────────────────────
    texto_integ = extrair_bloco(
        texto_completo,
        r"INTEGRAÇÃO DE NOVAS INSTALAÇÕES",
        [r"INFORMAÇÕES ADICIONAIS"] + PADROES_FIM
    )
    conteudo = texto_para_itens_html(texto_integ) if texto_integ else "<p style='margin:0; font-size:13px; color:#666; font-style:italic;'>Nada a relatar.</p>"
    secoes_html += bloco_secao(
        "Integração de Novas Instalações",
        conteudo,
        cor_fundo="#EAF3DE",
        cor_barra="#3B6D11",
        cor_titulo="#27500A"
    )

    return secoes_html if secoes_html else None

# ------------------------------
# SALVA NO GOOGLE DRIVE
# ------------------------------
def salvar_ipdo_no_drive(pdf_bytes, data_ref):
    ano = data_ref.strftime("%Y")
    mes = data_ref.strftime("%m")
    nome_arquivo = f"IPDO-{data_ref.strftime('%d-%m-%Y')}.pdf"
    payload = {
        "tipo": "IPDO",
        "ano": ano,
        "mes": mes,
        "nome_arquivo": nome_arquivo,
        "pdf_base64": base64.b64encode(pdf_bytes).decode("utf-8")
    }
    try:
        r = requests.post(DRIVE_WEBHOOK_URL, json=payload, timeout=30)
        if r.status_code == 200:
            print("📁 IPDO salvo no Drive com sucesso.")
        else:
            print(f"⚠️ Erro ao salvar IPDO no Drive: {r.status_code}")
    except Exception as e:
        print(f"⚠️ Falha ao enviar IPDO ao Drive: {e}")

# ------------------------------
# ENVIA O E-MAIL
# ------------------------------
def enviar_email():
    EMAIL_HOST = 'smtp.gmail.com'
    EMAIL_PORT = 587
    EMAIL_USUARIO = 'aplan.notificacoes@gmail.com'
    EMAIL_SENHA = 'dwbu nxga jnjt riwj'
    DESTINATARIOS = ['paulovictormcarneiro@gmail.com']
    NOME_REMETENTE = "A-Plan Notificação"

    hoje = datetime.now()
    ontem = hoje - timedelta(1)
    data_ontem = ontem.strftime('%d-%m-%Y')
    data_formatada = ontem.strftime('%d/%m/%Y')

    MENSAGEM = MIMEMultipart()
    MENSAGEM['From'] = NOME_REMETENTE
    MENSAGEM['To'] = EMAIL_USUARIO
    MENSAGEM['Subject'] = f'IPDO - {data_formatada} ⚠️'

    pdf_url = f"https://www.ons.org.br/AcervoDigitalDocumentosEPublicacoes/IPDO-{data_ontem}.pdf"

    try:
        pdf_response = requests.get(pdf_url, timeout=30)
        if pdf_response.status_code == 200:
            salvar_ipdo_no_drive(pdf_response.content, ontem)

            secoes_html = extrair_secoes_formatado(pdf_response.content)

            if not secoes_html:
                print("⚠️ Nenhuma seção extraída — verifique o PDF manualmente.")
                secoes_html = "<p style='color:red;font-weight:bold;'>Erro: Não foi possível extrair as seções do PDF.</p>"
            else:
                print("✅ Seções extraídas com sucesso.")

            corpo_html = f"""
<!DOCTYPE html>
<html>
<body style="margin:0; padding:0; background:#f0f4f8; font-family: Arial, sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f8; padding: 24px 0;">
    <tr>
      <td align="center">
        <table width="620" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:10px; overflow:hidden; border:1px solid #dde3ea;">

          <!-- HEADER -->
          <tr>
            <td style="background:#0C3B6E; padding:20px 28px;">
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="48">
                    <div style="width:42px; height:42px; border-radius:50%; background:#1E90FF; text-align:center; line-height:42px; font-size:14px; font-weight:600; color:#ffffff;">AP</div>
                  </td>
                  <td style="padding-left:12px;">
                    <p style="margin:0; font-size:15px; font-weight:600; color:#ffffff;">A-Plan Notificação</p>
                    <p style="margin:0; font-size:12px; color:#9EC8F5;">Informativo Preliminar Diário da Operação</p>
                  </td>
                  <td align="right">
                    <p style="margin:0; font-size:14px; font-weight:600; color:#ffffff;">IPDO</p>
                    <p style="margin:0; font-size:12px; color:#9EC8F5;">{data_formatada}</p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- BODY -->
          <tr>
            <td style="padding:24px 28px;">
              <p style="margin:0 0 20px; font-size:14px; color:#555;">Bom dia, segue abaixo os destaques do IPDO para o submercado Nordeste.</p>

              {secoes_html}

              <!-- ASSINATURA -->
              <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #e0e0e0; padding-top:16px; margin-top:8px;">
                <tr>
                  <td>
                    <p style="margin:0; font-size:13px; color:#888;">Atenciosamente,</p>
                    <p style="margin:0; font-size:13px; font-weight:600; color:#185FA5;">A-plan Notificação</p>
                  </td>
                  <td align="right">
                    <span style="font-size:11px; color:#888; background:#f5f5f5; padding:4px 10px; border-radius:4px; border:1px solid #e0e0e0;">ONS · IPDO</span>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
            """

            MENSAGEM.attach(MIMEText(corpo_html, 'html'))

            pdf_attachment = MIMEApplication(pdf_response.content, _subtype='pdf')
            pdf_attachment.add_header(
                'Content-Disposition',
                'attachment',
                filename=f'IPDO-{data_ontem}.pdf'
            )
            MENSAGEM.attach(pdf_attachment)

            destinatarios_reais = [EMAIL_USUARIO] + DESTINATARIOS
            with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as servidor:
                servidor.starttls()
                servidor.login(EMAIL_USUARIO, EMAIL_SENHA)
                servidor.sendmail(
                    EMAIL_USUARIO,
                    destinatarios_reais,
                    MENSAGEM.as_string()
                )
            print("✅ E-mail enviado com sucesso!")

        else:
            print(f"❌ Erro ao acessar o PDF: status {pdf_response.status_code}")

    except requests.exceptions.RequestException as e:
        print(f"❌ Erro ao tentar baixar o PDF: {e}")

# ------------------------------
# EXECUÇÃO
# ------------------------------
enviar_email()
