from playwright.sync_api import sync_playwright

URL_LOGIN = "https://sintegre.ons.org.br/sites/2/14/Paginas/servicos/historico-de-produtos.aspx?produto=RO%20(Relatório%20de%20Análise%20de%20Ocorrência)"

USUARIO_XPATH = '/html/body/div/div[2]/div/div/div[1]/div/form/div[1]/input'
SENHA_XPATH   = '/html/body/div/div[2]/div/div/div[1]/div/form/div[2]/input'
ENTRAR_XPATH  = '/html/body/div/div[2]/div/div/div[1]/div/form/div[4]/input[2]'

USERNAME = "paulo.victor@casadosventos.com.br"
PASSWORD = "windpower07#"

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-dev-shm-usage"]
    )

    context = browser.new_context()
    page = context.new_page()

    # Abre página
    page.goto(URL_LOGIN, timeout=60000)

    # Aguarda campos aparecerem
    page.wait_for_selector(f"xpath={USUARIO_XPATH}", timeout=60000)

    # Preenche login
    page.locator(f"xpath={USUARIO_XPATH}").fill(USERNAME)
    page.locator(f"xpath={SENHA_XPATH}").fill(PASSWORD)
    page.locator(f"xpath={ENTRAR_XPATH}").click()

    # Aguarda login concluir
    page.wait_for_timeout(15000)

    # Salva sessão
    context.storage_state(path="storage_state.json")
    print("✅ Sessão salva com sucesso na VPS")

    browser.close()

