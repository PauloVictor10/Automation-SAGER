 SAGER Automação — ONS + Supabase

Automação completa de extração de relatórios do sistema SAGER/ONS, com armazenamento no Supabase e notificação por e-mail.

📋 O que o sistema faz

🔐 Login automático no portal ONS (sintegre.ons.org.br) via sessão salva
📅 Calcula o período automaticamente (últimos 5 dias até ontem)
📊 Seleciona fontes Eólica e Fotovoltaica
📥 Gera e baixa o Relatório Geral em .xlsx
📦 Extrai o arquivo ZIP automaticamente
🗄️ Salva os dados no banco Supabase
📧 Envia e-mail com resumo da execução (usinas processadas, ignoradas, período)
🔄 Retry automático — 3 tentativas com intervalo de 10 segundos


🖥️ Exemplo de Notificação
✅ SAGER | Extração concluída — 08/04/2026 17:08

8 Processados  |  8 Usinas  |  1 Ignorado

Período: 2026-04-02 → 2026-04-07
Banco de dados: Supabase ✅

STATUS DAS USINAS:
✅ Folha Larga Sul
✅ Rio do Vento
✅ Babilônia Sul
✅ Rio do Vento Expansão
✅ Umari
✅ Babilônia Centro
✅ Serra do Tigre
✅ Babilônia Sul Solar
⚠️ RelatorioGeral_304 — Ignorado

🗂️ Estrutura do Projeto
/opt/automation/
├── gerar_sessao.py          # Gera e salva sessão autenticada no ONS
├── SAGER com Supabase.py    # Script principal de automação
├── requirements.txt         # Dependências do projeto
├── .gitignore               # Arquivos ignorados pelo Git
├── credencial.json          # 🔒 Chave Google (NÃO versionar)
├── storage_state.json       # 🔒 Sessão salva (NÃO versionar)
└── logs/
    ├── gerar_sessao.log
    └── sager.log

⚙️ Configuração
Variáveis principais
python# Google Drive
PASTA_DRIVE_ID       = "ID_DA_PASTA_NO_DRIVE"
SERVICE_ACCOUNT_FILE = "/opt/automation/credencial.json"
SCOPES               = ['https://www.googleapis.com/auth/drive']

# ONS
URL_LOGIN = "https://sintegre.ons.org.br/"
URL_SAGER = "https://pops.ons.org.br/pop/#17269"
USUARIO   = "seu.email@empresa.com.br"
SENHA     = "sua_senha"

# Supabase
SUPABASE_URL = "https://xxxx.supabase.co"
SUPABASE_KEY = "sua_chave_supabase"

# E-mail
EMAIL_REMETENTE    = "aplan.notificacoes@gmail.com"
EMAIL_DESTINATARIO = "paulo.victor.carneiro@empresa.com.br"

🚀 Instalação na VPS
bash# Conectar na VPS
ssh root@IP_DA_VPS

# Entrar na pasta
cd /opt/automation
source venv/bin/activate

# Instalar dependências
pip install -r requirements.txt
playwright install chromium
playwright install-deps chromium

🕐 Agendamento Cron
Roda todo dia automaticamente:
cron# Gera sessão às 08:00
0 8 * * * cd /opt/automation && /opt/automation/venv/bin/python gerar_sessao.py >> logs/gerar_sessao.log 2>&1

# Extrai relatório às 08:05
5 8 * * * cd /opt/automation && /opt/automation/venv/bin/python "SAGER com Supabase.py" >> logs/sager.log 2>&1

📦 Dependências
BibliotecaUsoplaywrightAutomação do navegador headlessgoogle-authAutenticação Google Service Accountgoogle-api-python-clientUpload para o Google DrivesupabaseArmazenamento dos dados extraídospandas / openpyxlLeitura e manipulação do .xlsxsmtplibEnvio de e-mail com resumo
Instalar tudo:
bashpip install -r requirements.txt

🔄 Fluxo Completo
gerar_sessao.py          SAGER com Supabase.py
      │                          │
      ▼                          ▼
 Login ONS              Carrega sessão salva
      │                          │
 Salva sessão           Seleciona datas + fontes
      │                          │
storage_state.json      Gera relatório .xlsx
                                 │
                         Salva no Supabase
                                 │
                         Envia e-mail ✅
