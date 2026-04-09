import os
import json
import zipfile
import time
import traceback
from datetime import datetime, timedelta

import pandas as pd
from playwright.sync_api import sync_playwright
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account

# ==============================
# CONFIGURAÇÕES
# ==============================

PASTA_DRIVE_ID = "1MZyMtpxoNCMgYbFXSGwbRT-2fABA8ArR"

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

URL_LOGIN    = "https://sintegre.ons.org.br/"
URL_SAGER    = "https://pops.ons.org.br/pop/#17269"
USUARIO      = "paulo.victor@casadosventos.com.br"
SENHA        = "windpower07#"
PASTA_PRINTS = r"C:\Users\paulo\Desktop\prints_ons"

# Range de dados no Sheets
LINHA_INICIO_EXCEL = 10  # linha onde os dados começam no Excel (pula cabeçalho do relatório)
LINHA_INICIO_SHEETS = 2  # linha onde inserir no Sheets (após cabeçalho formatado)
COL_INICIO   = "A"
COL_FIM      = "I"
NUM_COLUNAS  = 9   # A até I
ABA_DADOS    = "Dados"

# ==============================
# MAPEAMENTO: código → (nome da usina, ID do Google Sheets)
#
# Para cada usina:
#   1. Crie um Google Sheets no Drive
#   2. Compartilhe com o e-mail da service account (campo "client_email"
#      no credencial.json) como Editor
#   3. Copie o ID da URL (parte entre /d/ e /edit) e cole abaixo
# ==============================
MAPA_USINAS = {
    "140": {"nome": "Folha Larga Sul",       "sheets_id": "1ydvkFfgHKW5Tamy8fjIu_xN1RNz7ei3fwEbsvSn641I"},
    "166": {"nome": "Rio do Vento",           "sheets_id": "1TYTvz0fRLXXR69eQQKQ_DW8H2CqA7IaoIT8ziLD7EXY"},
    "210": {"nome": "Babilônia Sul",          "sheets_id": "1qtFIaoOV2AKePSLrnBx01UYHTAKApwPvs1xj1euIz9k"},
    "220": {"nome": "Rio do Vento Expansão",  "sheets_id": "1I7brsGhefmNcHKm3OFlPo18lvQSjcPY-fZtt5QP5IIc"},
    "256": {"nome": "Umari",                  "sheets_id": "1oy5SmrzBpylsV1LPF91_Ltco3WxRTx7fsS46rbtvvjQ"},
    "275": {"nome": "Babilônia Centro",       "sheets_id": "1q8A-aG-ydx6yjIkR8kxJ_UjHj-vYXPB8-wC7QQXZVHo"},
    "291": {"nome": "Serra do Tigre",         "sheets_id": "1IBrZz-HT0RJI9XUWN73Mk-G6gLXfQhc_HPJRcPk5RBU"},
    "303": {"nome": "Babilônia Sul Solar",    "sheets_id": "1P0lhzBsaJl5UqToy8x7lY7hWzmJ7afrWpIu5KR4YWMA"},
}


# ==============================
# CREDENCIAIS — variável de ambiente
# ==============================

def get_credentials():
    """
    Carrega credenciais em duas etapas:
      1. Variável de ambiente GOOGLE_CREDENTIALS  (produção)
      2. Arquivo credencial.json na pasta do script (desenvolvimento)
    """
    # 1️⃣ Tenta variável de ambiente
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        try:
            info = json.loads(creds_json)
            print("  🔑 Credenciais via variável de ambiente.")
            return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        except json.JSONDecodeError as e:
            raise ValueError(f"❌ GOOGLE_CREDENTIALS com JSON inválido: {e}")

    # 2️⃣ Fallback: credencial.json na mesma pasta do script
    pasta_script  = os.path.dirname(os.path.abspath(__file__))
    arquivo_creds = os.path.join(pasta_script, "credencial.json")

    if os.path.exists(arquivo_creds):
        print(f"  🔑 Credenciais via arquivo: {arquivo_creds}")
        return service_account.Credentials.from_service_account_file(arquivo_creds, scopes=SCOPES)

    raise EnvironmentError(
        "❌ Credenciais não encontradas! Faça uma das opções:\n"
        "   1. Configure a variável de ambiente GOOGLE_CREDENTIALS\n"
        f"  2. Coloque o credencial.json em: {pasta_script}"
    )


# ==============================
# UTILITÁRIOS
# ==============================

def extrair_codigo(nome_arquivo):
    """
    'RelatorioGeral_140_01012026_28022026.xlsx' → '140'
    """
    try:
        partes = os.path.splitext(nome_arquivo)[0].split("_")
        for parte in partes:
            if parte.isdigit() and len(parte) == 3:
                return parte
    except Exception:
        pass
    return None


# ==============================
# GOOGLE SHEETS — upsert A9:I
# ==============================

def upsert_no_sheets(caminho_arquivo, sheets_id, nome_usina):
    """
    Lê colunas A:I a partir da linha 9 do arquivo.
    - Chave = coluna A
    - Se a chave já existe no Sheets → substitui a linha
    - Se não existe → insere ao final
    - Linhas com coluna A vazia são ignoradas
    """
    print(f"  📊 Processando Sheets: {nome_usina}")

    ext = os.path.splitext(caminho_arquivo)[1].lower()
    if ext == ".xlsx":
        df_raw = pd.read_excel(caminho_arquivo, sheet_name='Restrições', header=None)
    elif ext == ".csv":
        df_raw = pd.read_csv(caminho_arquivo, header=None,
                             encoding="utf-8", sep=None, engine="python")
    else:
        print(f"    ⚠️  Formato não suportado: {ext}")
        return

    # Colunas A:I a partir da linha 9
    df = df_raw.iloc[LINHA_INICIO_EXCEL - 1:, :NUM_COLUNAS].copy()
    df = df.fillna("").astype(str)
    df = df[df.iloc[:, 0].str.strip() != ""]  # ignora linhas sem chave

    if df.empty:
        print("    ⚠️  Nenhuma linha com dado na coluna A.")
        return

    novas_linhas = df.values.tolist()
    print(f"    📄 {len(novas_linhas)} linha(s) no arquivo")

    # Lê estado atual do Sheets
    creds          = get_credentials()
    sheets_service = build("sheets", "v4", credentials=creds)
    sheet          = sheets_service.spreadsheets()

    resultado = sheet.values().get(
        spreadsheetId=sheets_id,
        range=f"{COL_INICIO}{LINHA_INICIO_SHEETS}:{COL_FIM}"
    ).execute()

    linhas_sheets = resultado.get("values", [])

    # Mapa: chave (col A) → número real da linha no Sheets
    mapa_chave_linha = {}
    for i, linha in enumerate(linhas_sheets):
        chave = linha[0].strip() if linha else ""
        if chave:
            mapa_chave_linha[chave] = LINHA_INICIO_SHEETS + i

    updates         = []   # substituições (batchUpdate)
    linhas_inserir  = []   # novas linhas (append)

    for nova_linha in novas_linhas:
        chave = nova_linha[0].strip()
        if chave in mapa_chave_linha:
            num = mapa_chave_linha[chave]
            updates.append({
                "range":          f"{COL_INICIO}{num}:{COL_FIM}{num}",
                "values":         [nova_linha],
                "majorDimension": "ROWS"
            })
        else:
            linhas_inserir.append(nova_linha)

    # Executa substituições em batch
    if updates:
        sheet.values().batchUpdate(
            spreadsheetId=sheets_id,
            body={"valueInputOption": "USER_ENTERED", "data": updates}
        ).execute()
        print(f"    🔄 {len(updates)} linha(s) substituída(s)")

    # Insere linhas novas
    if linhas_inserir:
        sheet.values().append(
            spreadsheetId=sheets_id,
            range=f"{COL_INICIO}{LINHA_INICIO_SHEETS}:{COL_FIM}",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": linhas_inserir}
        ).execute()
        print(f"    ➕ {len(linhas_inserir)} linha(s) inserida(s)")

    print(f"    ✅ Sheets atualizado!")


# ==============================
# GOOGLE DRIVE — upload arquivo original
# ==============================

def upload_para_drive(caminho_arquivo):
    creds   = get_credentials()
    service = build("drive", "v3", credentials=creds)

    file_metadata = {
        "name":    os.path.basename(caminho_arquivo),
        "parents": [PASTA_DRIVE_ID]
    }
    media = MediaFileUpload(caminho_arquivo, resumable=True)
    file  = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True
    ).execute()

    print(f"    ✅ Arquivo enviado para o Drive! ID: {file.get('id')}")


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
    # Valida credenciais antes de abrir o browser
    print("🔑 Validando credenciais...")
    get_credentials()
    print("  ✅ Credenciais OK!\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--start-maximized"])
        context = browser.new_context()
        page    = context.new_page()

        os.makedirs(PASTA_PRINTS, exist_ok=True)

        # --- LOGIN ---
        print("🔐 Fazendo login...")
        page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_selector('input[name="username"]', timeout=15000)
        page.fill('input[name="username"]', USUARIO)
        page.fill('input[name="password"]', SENHA)
        page.click("#kc-login")
        page.wait_for_load_state("domcontentloaded", timeout=60000)
        page.wait_for_timeout(3000)
        print(f"  ✅ Login OK! URL atual: {page.url}")

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
            print("  ℹ️  Navegação abortada (normal em SPA), aguardando carregamento...")

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
        caminho_zip   = os.path.join(PASTA_PRINTS, nome_original)
        download.save_as(caminho_zip)
        print(f"  📦 ZIP salvo: {caminho_zip}")

        # --- EXTRAIR ZIP ---
        with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
            zip_ref.extractall(PASTA_PRINTS)
        os.remove(caminho_zip)
        print("  ✅ ZIP extraído e removido")

        # --- PROCESSAR CADA ARQUIVO ---
        print("\n🚀 Processando arquivos extraídos...\n")
        processados = 0
        ignorados   = []

        for arquivo in sorted(os.listdir(PASTA_PRINTS)):
            ext = os.path.splitext(arquivo)[1].lower()
            if ext not in (".xlsx", ".csv"):
                continue

            caminho_completo = os.path.join(PASTA_PRINTS, arquivo)
            codigo           = extrair_codigo(arquivo)

            print(f"📁 {arquivo}  →  código: {codigo}")

            if codigo and codigo in MAPA_USINAS:
                usina     = MAPA_USINAS[codigo]
                sheets_id = usina["sheets_id"]
                nome      = usina["nome"]

                # Salva dados no Sheets (A9:I, chave = col A)
                upsert_no_sheets(caminho_completo, sheets_id, nome)

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
        print("🎉 Processo finalizado!")


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
                raise
            print("🔄 Aguardando 10 segundos...\n")
            time.sleep(10)


executar_com_retry()