import os
import json
import zipfile
import time
import traceback
import smtplib
import shutil
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta

import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
from playwright.sync_api import sync_playwright
from google.oauth2 import service_account

# ==============================
# CONFIGURAÇÕES
# ==============================

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

URL_LOGIN    = "https://sintegre.ons.org.br/"
URL_SAGER    = "https://pops.ons.org.br/pop/#17269"
USUARIO      = "paulo.victor@casadosventos.com.br"
SENHA        = "windpower07#"

# Linha onde os dados começam no Excel (pula cabeçalho do relatório ONS)
LINHA_INICIO_EXCEL = 10

# ==============================
# MAPEAMENTO: código → nome da usina
# ==============================
MAPA_USINAS = {
    "140": "Folha Larga Sul",
    "166": "Rio do Vento",
    "210": "Babilônia Sul",
    "220": "Rio do Vento Expansão",
    "256": "Umari",
    "275": "Babilônia Centro",
    "291": "Serra do Tigre",
    "303": "Babilônia Sul Solar",
}


# ==============================
# EMAIL
# ==============================
EMAIL_SISTEMA = "aplan.notificacoes@gmail.com"
EMAIL_SENHA   = "dwbu nxga jnjt riwj"  # senha de app Gmail
EMAIL_CC      = "paulo.victor.carneiro@outlook.com"
EMAIL_TO      = "paulovictormcarneiro@gmail.com"


def saudacao():
    hora = datetime.now().hour
    if 5 <= hora < 12:
        return "Bom dia"
    elif 12 <= hora < 18:
        return "Boa tarde"
    else:
        return "Boa noite"


def enviar_email_sucesso(processados, ignorados, data_inicio, data_fim):
    try:
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        saud  = saudacao()

        usinas_rows = "".join([f"""
            <tr>
              <td style="padding:10px 16px;border-bottom:1px solid #f0f4f8;font-size:13px;color:#333;">
                🏭 {MAPA_USINAS[k]}
              </td>
              <td style="padding:10px 16px;border-bottom:1px solid #f0f4f8;text-align:center;">
                <span style="background:#d4edda;color:#155724;padding:3px 12px;border-radius:20px;font-size:11px;font-weight:bold;">✔ OK</span>
              </td>
            </tr>""" for k in MAPA_USINAS])

        ignorados_rows = "".join([f"""
            <tr>
              <td style="padding:10px 16px;border-bottom:1px solid #f0f4f8;font-size:13px;color:#333;">⚠️ {a}</td>
              <td style="padding:10px 16px;border-bottom:1px solid #f0f4f8;text-align:center;">
                <span style="background:#fff3cd;color:#856404;padding:3px 12px;border-radius:20px;font-size:11px;font-weight:bold;">Ignorado</span>
              </td>
            </tr>""" for a in ignorados]) if ignorados else """
            <tr>
              <td colspan="2" style="padding:10px 16px;font-size:13px;color:#888;font-style:italic;">
                Nenhum arquivo ignorado
              </td>
            </tr>"""

        corpo = f"""<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:40px 0;">
<tr><td align="center">
<table width="580" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.10);">

  <!-- HEADER -->
  <tr>
    <td style="background:linear-gradient(135deg,#1a73e8 0%,#0d47a1 100%);padding:36px 40px;text-align:center;">
      <p style="margin:0;font-size:28px;">⚡</p>
      <h1 style="margin:8px 0 4px;color:#ffffff;font-size:20px;font-weight:700;letter-spacing:1px;">SAGER Automação</h1>
      <p style="margin:0;color:#90caf9;font-size:12px;letter-spacing:1px;">SISTEMA DE EXTRAÇÃO AUTOMÁTICA · ONS</p>
    </td>
  </tr>

  <!-- STATUS BADGE -->
  <tr>
    <td style="padding:28px 40px 0;text-align:center;">
      <span style="background:#d4edda;color:#155724;padding:8px 24px;border-radius:30px;font-size:13px;font-weight:bold;">
        ✅ Extração concluída com sucesso
      </span>
    </td>
  </tr>

  <!-- SAUDAÇÃO -->
  <tr>
    <td style="padding:20px 40px 0;">
      <p style="margin:0;font-size:15px;color:#333;font-weight:600;">{saud}! 👋</p>
      <p style="margin:8px 0 0;font-size:13px;color:#666;line-height:1.6;">
        A extração automática dos relatórios do SAGER foi concluída. Os dados foram salvos no banco de dados Supabase com sucesso.
      </p>
    </td>
  </tr>

  <!-- CARDS -->
  <tr>
    <td style="padding:24px 40px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td width="32%" style="padding-right:8px;">
            <div style="background:#e8f5e9;border-radius:12px;padding:18px 12px;text-align:center;">
              <p style="margin:0;font-size:30px;font-weight:bold;color:#2e7d32;">{processados}</p>
              <p style="margin:4px 0 0;font-size:11px;color:#555;text-transform:uppercase;letter-spacing:1px;">Processados</p>
            </div>
          </td>
          <td width="32%" style="padding:0 4px;">
            <div style="background:#e3f2fd;border-radius:12px;padding:18px 12px;text-align:center;">
              <p style="margin:0;font-size:30px;font-weight:bold;color:#1565c0;">{len(MAPA_USINAS)}</p>
              <p style="margin:4px 0 0;font-size:11px;color:#555;text-transform:uppercase;letter-spacing:1px;">Usinas</p>
            </div>
          </td>
          <td width="32%" style="padding-left:8px;">
            <div style="background:{'#fff3cd' if ignorados else '#f3e5f5'};border-radius:12px;padding:18px 12px;text-align:center;">
              <p style="margin:0;font-size:30px;font-weight:bold;color:{'#856404' if ignorados else '#6a1b9a'};">{len(ignorados)}</p>
              <p style="margin:4px 0 0;font-size:11px;color:#555;text-transform:uppercase;letter-spacing:1px;">Ignorados</p>
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- DETALHES -->
  <tr>
    <td style="padding:0 40px 20px;">
      <table width="100%" cellpadding="0" cellspacing="0" style="border-radius:12px;overflow:hidden;border:1px solid #e9ecef;">
        <tr style="background:#f8f9fa;">
          <td colspan="2" style="padding:10px 16px;font-size:11px;font-weight:bold;color:#6c757d;text-transform:uppercase;letter-spacing:1px;">
            📋 Detalhes da Execução
          </td>
        </tr>
        <tr>
          <td style="padding:10px 16px;font-size:13px;color:#666;width:40%;">🕐 Data/hora</td>
          <td style="padding:10px 16px;font-size:13px;color:#333;font-weight:600;">{agora}</td>
        </tr>
        <tr style="background:#f8f9fa;">
          <td style="padding:10px 16px;font-size:13px;color:#666;">📅 Período</td>
          <td style="padding:10px 16px;font-size:13px;color:#333;font-weight:600;">{data_inicio} → {data_fim}</td>
        </tr>
        <tr>
          <td style="padding:10px 16px;font-size:13px;color:#666;">🗄️ Banco de dados</td>
          <td style="padding:10px 16px;">
            <span style="background:#d4edda;color:#155724;padding:3px 12px;border-radius:20px;font-size:11px;font-weight:bold;">Supabase ✔</span>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- USINAS -->
  <tr>
    <td style="padding:0 40px 32px;">
      <p style="margin:0 0 8px;font-size:11px;font-weight:bold;color:#6c757d;text-transform:uppercase;letter-spacing:1px;">🏭 Status das Usinas</p>
      <table width="100%" cellpadding="0" cellspacing="0" style="border-radius:12px;overflow:hidden;border:1px solid #e9ecef;">
        <tr style="background:#f8f9fa;">
          <td style="padding:10px 16px;font-size:11px;font-weight:bold;color:#6c757d;text-transform:uppercase;letter-spacing:1px;">Usina</td>
          <td style="padding:10px 16px;font-size:11px;font-weight:bold;color:#6c757d;text-transform:uppercase;letter-spacing:1px;text-align:center;">Status</td>
        </tr>
        {usinas_rows}
        {ignorados_rows}
      </table>
    </td>
  </tr>

  <!-- FOOTER -->
  <tr>
    <td style="background:#f8f9fa;padding:20px 40px;text-align:center;border-top:1px solid #e9ecef;">
      <p style="margin:0;font-size:12px;color:#aaa;">Aplan Notificações · Sistema SAGER Automação</p>
      <p style="margin:4px 0 0;font-size:11px;color:#ccc;">{agora}</p>
    </td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>"""

        msg = MIMEMultipart("alternative")
        msg["From"]    = EMAIL_SISTEMA
        msg["To"]      = EMAIL_TO
        msg["Cc"]      = EMAIL_CC
        msg["Subject"] = f"✅ SAGER | Extração concluída — {agora}"
        msg.attach(MIMEText(corpo, "html", "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_SISTEMA, EMAIL_SENHA)
            smtp.send_message(msg)

        print(f"  📧 E-mail enviado para {EMAIL_TO}")

    except Exception as e:
        print(f"  ⚠️  Falha ao enviar e-mail: {e}")


def enviar_email_erro(erro, tentativa):
    try:
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        saud  = saudacao()

        corpo = f"""
{saud}!

A extração automática do SAGER falhou após {tentativa} tentativa(s). ❌

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 DETALHES DO ERRO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 Data/hora  : {agora}
 Tentativas : {tentativa}
 Erro       : {str(erro)}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Por favor, verifique o sistema.

— Aplan Notificações 
        """.strip()

        msg = MIMEMultipart()
        msg["From"]    = EMAIL_SISTEMA
        msg["To"]      = EMAIL_TO
        msg["Cc"]      = EMAIL_CC
        msg["Subject"] = f"❌ SAGER | Falha na extração — {agora}"
        msg.attach(MIMEText(corpo, "plain", "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_SISTEMA, EMAIL_SENHA)
            smtp.send_message(msg)

        print(f"  📧 E-mail de erro enviado para {EMAIL_TO}")

    except Exception as e:
        print(f"  ⚠️  Falha ao enviar e-mail de erro: {e}")



# ==============================
# CREDENCIAIS GOOGLE
# ==============================

def get_credentials():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        try:
            info = json.loads(creds_json)
            print("  🔑 Credenciais via variável de ambiente.")
            return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        except json.JSONDecodeError as e:
            raise ValueError(f"❌ GOOGLE_CREDENTIALS com JSON inválido: {e}")

    pasta_script  = os.path.dirname(os.path.abspath(__file__))
    arquivo_creds = os.path.join(pasta_script, "credencial.json")
    if os.path.exists(arquivo_creds):
        print(f"  🔑 Credenciais via arquivo: {arquivo_creds}")
        return service_account.Credentials.from_service_account_file(arquivo_creds, scopes=SCOPES)

    raise EnvironmentError(
        "❌ Credenciais não encontradas!\n"
        "   1. Configure a variável GOOGLE_CREDENTIALS\n"
        f"  2. Coloque credencial.json em: {pasta_script}"
    )


# ==============================
# CONEXÃO POSTGRESQL — Supabase
# ==============================

# Fallback: connection string direta (caso variável de ambiente não esteja disponível)
DATABASE_URL_FALLBACK = "postgresql://postgres.evfmquiajtvzvhfzwvfs:owvyNbmiYetZ61JA@aws-1-sa-east-1.pooler.supabase.com:5432/postgres"

def get_db_connection():
    db_url = os.environ.get("DATABASE_URL") or DATABASE_URL_FALLBACK
    origem = "variável de ambiente" if os.environ.get("DATABASE_URL") else "fallback direto"
    print(f"  🗄️  Conectando via: {origem}")
    return psycopg2.connect(db_url)


def criar_tabela_se_nao_existir(cursor, nome_tabela):
    """Cria a tabela da usina se ainda não existir."""
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS "{nome_tabela}" (
            id               SERIAL PRIMARY KEY,
            data             DATE,
            hora_inicial     TIME,
            hora_final       TIME,
            razao            TEXT,
            origem           TEXT,
            valor_limitacao  NUMERIC,
            descricao        TEXT,
            data_extracao    TIMESTAMP DEFAULT NOW(),
            UNIQUE (data, hora_inicial, hora_final)
        );
    """)


# ==============================
# UTILITÁRIOS
# ==============================

def extrair_codigo(nome_arquivo):
    try:
        partes = os.path.splitext(nome_arquivo)[0].split("_")
        for parte in partes:
            if parte.isdigit() and len(parte) == 3:
                return parte
    except Exception:
        pass
    return None


# ==============================
# UPSERT NO POSTGRESQL
# ==============================

def upsert_no_postgres(caminho_arquivo, nome_usina):
    """
    Lê aba 'Restrições' do Excel a partir da linha 10.
    Faz upsert no PostgreSQL — chave: (data, hora_inicial, hora_final).
    """
    print(f"  🗄️  Salvando no banco: {nome_usina}")

    ext = os.path.splitext(caminho_arquivo)[1].lower()
    if ext == ".xlsx":
        try:
            df_raw = pd.read_excel(caminho_arquivo, sheet_name="Restrições", header=None)
        except Exception as e:
            print(f"    ⚠️  Erro ao ler aba 'Restrições': {e}")
            return
    elif ext == ".csv":
        df_raw = pd.read_csv(caminho_arquivo, header=None, encoding="utf-8",
                             sep=None, engine="python")
    else:
        print(f"    ⚠️  Formato não suportado: {ext}")
        return

    # Pega dados a partir da linha 10 (índice 9), colunas A:G (0:7)
    df = df_raw.iloc[LINHA_INICIO_EXCEL - 1:, :7].copy()
    df.columns = ["data", "hora_inicial", "hora_final", "razao",
                  "origem", "valor_limitacao", "descricao"]
    df = df.dropna(subset=["data"])
    df = df.fillna("")

    if df.empty:
        print("    ⚠️  Nenhuma linha com dados.")
        return

    print(f"    📄 {len(df)} linha(s) encontradas")

    # Nome da tabela = código da usina sem espaços
    nome_tabela = nome_usina.lower().replace(" ", "_").replace("ô", "o").replace("â", "a").replace("ã", "a").replace("é", "e").replace("í", "i")

    conn   = get_db_connection()
    cursor = conn.cursor()

    try:
        # Cria tabela se não existir
        criar_tabela_se_nao_existir(cursor, nome_tabela)

        # Prepara linhas para upsert
        linhas = []
        for _, row in df.iterrows():
            try:
                data            = pd.to_datetime(row["data"], dayfirst=True).date()
                hora_inicial    = str(row["hora_inicial"]).strip() or None
                hora_final      = str(row["hora_final"]).strip() or None
                razao           = str(row["razao"]).strip()
                origem          = str(row["origem"]).strip()
                valor_limitacao = row["valor_limitacao"]
                descricao       = str(row["descricao"]).strip()

                # Converte valor para float
                if isinstance(valor_limitacao, str):
                    valor_limitacao = valor_limitacao.replace(",", ".").strip()
                    valor_limitacao = float(valor_limitacao) if valor_limitacao else None
                elif pd.isna(valor_limitacao):
                    valor_limitacao = None

                linhas.append((data, hora_inicial, hora_final, razao,
                               origem, valor_limitacao, descricao))
            except Exception as e:
                print(f"    ⚠️  Linha ignorada: {e}")
                continue

        # Upsert — se (data, hora_inicial, hora_final) já existe, atualiza
        sql = f"""
            INSERT INTO "{nome_tabela}"
                (data, hora_inicial, hora_final, razao, origem, valor_limitacao, descricao, data_extracao)
            VALUES %s
            ON CONFLICT (data, hora_inicial, hora_final)
            DO UPDATE SET
                razao           = EXCLUDED.razao,
                origem          = EXCLUDED.origem,
                valor_limitacao = EXCLUDED.valor_limitacao,
                descricao       = EXCLUDED.descricao,
                data_extracao   = NOW();
        """

        # Adiciona timestamp de extração
        linhas_com_ts = [l + (datetime.now(),) for l in linhas]

        execute_values(cursor, sql, linhas_com_ts)
        conn.commit()

        print(f"    ✅ {len(linhas)} linha(s) salvas no banco!")

    except Exception as e:
        conn.rollback()
        raise e
    finally:
        cursor.close()
        conn.close()


# ==============================
# SELEÇÃO DE DATA NO CALENDÁRIO
# ==============================

def selecionar_data(frame, indice_calendario, data_dt):
    meses = {
        1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR",
        5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO",
        9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
    }

    mes_desejado = meses[data_dt.month]
    ano_desejado = str(data_dt.year)
    dia_desejado = str(data_dt.day)

    frame.locator("button[aria-label='Open calendar']").nth(indice_calendario).click()
    frame.wait_for_selector(".mat-calendar-body")

    while True:
        header = frame.locator(".mat-calendar-period-button").inner_text()
        print(f"    Calendário: {header}")
        if mes_desejado in header and ano_desejado in header:
            break
        frame.locator(".mat-calendar-previous-button").click()
        frame.wait_for_timeout(300)

    frame.locator(".mat-calendar-body-cell-content", has_text=dia_desejado).first.click()


# ==============================
# FLUXO PRINCIPAL
# ==============================

def run():
    print("🔑 Validando credenciais Google...")
    get_credentials()
    print("  ✅ OK!\n")

    print("🗄️  Validando conexão com banco de dados...")
    conn = get_db_connection()
    conn.close()
    print("  ✅ Supabase conectado!\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--start-maximized"])
        context = browser.new_context()
        page    = context.new_page()

        # --- LOGIN ---
        print("🔐 Fazendo login...")
        page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_selector('input[name="username"]', timeout=15000)
        page.fill('input[name="username"]', USUARIO)
        page.fill('input[name="password"]', SENHA)
        page.click("#kc-login")
        page.wait_for_load_state("domcontentloaded", timeout=60000)
        page.wait_for_timeout(3000)
        print(f"  ✅ Login OK! URL: {page.url}")

        try:
            page.wait_for_selector("text=Concordo", timeout=5000)
            page.get_by_text("Concordo", exact=True).click()
            page.wait_for_timeout(2000)
        except:
            pass

        # --- NAVEGAR PARA O SAGER ---
        print("🌐 Acessando SAGER...")
        try:
            page.goto(URL_SAGER, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            print("  ℹ️  Navegação abortada (normal em SPA)...")

        page.wait_for_timeout(5000)
        page.wait_for_load_state("networkidle", timeout=60000)
        page.wait_for_selector("iframe", timeout=60000)

        frame = None
        for _ in range(30):
            for f in page.frames:
                if "apps18.ons.org.br" in f.url:
                    frame = f
                    break
            if frame:
                break
            page.wait_for_timeout(1000)

        if not frame:
            raise Exception("Frame não encontrado após espera")

        # --- PERÍODO ---
        hoje           = datetime.now()
        data_fim_dt    = hoje - timedelta(days=1)
        data_inicio_dt = data_fim_dt - timedelta(days=5)
        print(f"📅 Período: {data_inicio_dt.date()} → {data_fim_dt.date()}")

        selecionar_data(frame, 0, data_inicio_dt)
        print("  ✅ Data início selecionada!")
        selecionar_data(frame, 1, data_fim_dt)
        print("  ✅ Data fim selecionada!")
        page.wait_for_timeout(2000)

        # --- TIPO DE RELATÓRIO ---
        frame.locator("mat-select").first.click()
        frame.get_by_text("Relatório Geral", exact=True).click()
        print("📋 Tipo: Relatório Geral")
        page.wait_for_timeout(1000)

        # --- FONTES ---
        frame.locator("mat-select").nth(1).click()
        page.wait_for_timeout(1000)
        page.keyboard.press("Enter")
        page.keyboard.press("ArrowDown")
        page.wait_for_timeout(300)
        page.keyboard.press("Enter")
        page.keyboard.press("Escape")
        print("⚡ Fontes: Eolielétrica + Fotovoltaica")
        page.wait_for_timeout(1000)

        # --- GERAR RELATÓRIO ---
        print("\n⏳ Gerando relatório...")
        botao_gerar = frame.get_by_role("button", name="Gerar Relatório", exact=True)

        with page.expect_download(timeout=120000) as download_info:
            botao_gerar.click()

        download      = download_info.value
        nome_original = download.suggested_filename
        pasta_temp    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_temp")
        os.makedirs(pasta_temp, exist_ok=True)
        caminho_zip   = os.path.join(pasta_temp, nome_original)
        download.save_as(caminho_zip)
        print(f"  📦 ZIP salvo: {caminho_zip}")

        # --- EXTRAIR ZIP ---
        with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
            zip_ref.extractall(pasta_temp)
        os.remove(caminho_zip)
        print("  ✅ ZIP extraído e removido")

        # --- PROCESSAR CADA ARQUIVO ---
        print("\n🚀 Processando arquivos...\n")
        processados = 0
        ignorados   = []

        for arquivo in sorted(os.listdir(pasta_temp)):
            ext = os.path.splitext(arquivo)[1].lower()
            if ext not in (".xlsx", ".csv"):
                continue

            caminho_completo = os.path.join(pasta_temp, arquivo)
            codigo           = extrair_codigo(arquivo)

            print(f"📁 {arquivo}  →  código: {codigo}")

            if codigo and codigo in MAPA_USINAS:
                nome = MAPA_USINAS[codigo]
                upsert_no_postgres(caminho_completo, nome)
                processados += 1
            else:
                print(f"  ⚠️  Código '{codigo}' não mapeado — ignorado.")
                ignorados.append(arquivo)

        # --- RESUMO ---
        print(f"\n{'='*50}")
        print(f"✅ Processados : {processados} arquivo(s)")
        if ignorados:
            print(f"⚠️  Ignorados   : {len(ignorados)} arquivo(s)")
            for a in ignorados:
                print(f"   - {a}")
        # Limpa pasta temporária
        shutil.rmtree(pasta_temp, ignore_errors=True)
        print("🎉 Processo finalizado!")

        # --- E-MAIL DE SUCESSO ---
        print("\n📧 Enviando e-mail de notificação...")
        enviar_email_sucesso(processados, ignorados, data_inicio_dt.date(), data_fim_dt.date())


# ==============================
# RETRY AUTOMÁTICO
# ==============================

def executar_com_retry(max_tentativas=3):
    for tentativa in range(1, max_tentativas + 1):
        try:
            print(f"\n===== TENTATIVA {tentativa} de {max_tentativas} =====\n")
            run()
            return
        except Exception as e:
            print(f"\n❌ Erro na tentativa {tentativa}: {e}")
            traceback.print_exc()
            if tentativa == max_tentativas:
                print("\n🚨 Todas as tentativas falharam.")
                enviar_email_erro(e, tentativa)
                raise
            print("🔄 Aguardando 10 segundos...\n")
            time.sleep(10)


executar_com_retry()