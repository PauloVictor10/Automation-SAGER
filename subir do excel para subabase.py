import os
import glob
from datetime import datetime

import pandas as pd
import psycopg2
from psycopg2.extras import execute_values

# ==============================
# CONFIGURAÇÕES
# ==============================

PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")

DATABASE_URL_FALLBACK = "postgresql://postgres.evfmquiajtvzvhfzwvfs:owvyNbmiYetZ61JA@aws-1-sa-east-1.pooler.supabase.com:5432/postgres"

# Linha onde os dados começam no Excel (pula cabeçalho do relatório ONS)
LINHA_INICIO_EXCEL = 10

# Mapeamento: trecho no nome do arquivo → nome da usina (tabela)
MAPA_SIGLAS = {
    "BBS Solar":         "Babilônia Sul Solar",
    "BBS":               "Babilônia Sul",
    "BBC":               "Babilônia Centro",
    "FLS":               "Folha Larga Sul",
    "RDV":               "Rio do Vento",
    "RVE":               "Rio do Vento Expansão",
    "UMR":               "Umari",
    "TGR":               "Serra do Tigre",
}

# ==============================
# CONEXÃO POSTGRESQL
# ==============================

def get_db_connection():
    db_url = os.environ.get("DATABASE_URL") or DATABASE_URL_FALLBACK
    return psycopg2.connect(db_url)


def criar_tabela_se_nao_existir(cursor, nome_tabela):
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


def nome_para_tabela(nome_usina):
    return (nome_usina.lower()
        .replace(" ", "_")
        .replace("ô", "o").replace("â", "a").replace("ã", "a")
        .replace("é", "e").replace("í", "i").replace("õ", "o")
        .replace("ú", "u").replace("ç", "c").replace("à", "a"))


# ==============================
# IDENTIFICAR USINA PELO NOME DO ARQUIVO
# ==============================

def identificar_usina(nome_arquivo):
    nome_upper = nome_arquivo.upper()
    # Verifica "BBS Solar" antes de "BBS" para evitar match errado
    for sigla, nome in MAPA_SIGLAS.items():
        if sigla.upper() in nome_upper:
            return nome
    return None


# ==============================
# LER EXCEL
# ==============================

def ler_excel(caminho):
    try:
        df_raw = pd.read_excel(caminho, sheet_name="Restrições", header=None)
    except Exception:
        try:
            # Tenta primeira aba se não encontrar "Restrições"
            df_raw = pd.read_excel(caminho, sheet_name=0, header=None)
        except Exception as e:
            print(f"    ❌ Erro ao ler arquivo: {e}")
            return pd.DataFrame()

    # Dados a partir da linha 10 (índice 9), colunas A:G
    df = df_raw.iloc[LINHA_INICIO_EXCEL - 1:, :7].copy()
    df.columns = ["data", "hora_inicial", "hora_final", "razao",
                  "origem", "valor_limitacao", "descricao"]

    df = df.fillna("").astype(str)

    # Filtros de limpeza
    df = df[df["data"].str.strip() != ""]
    df = df[df["data"].str.lower() != "nan"]
    df = df[df["hora_inicial"].str.strip() != ""]
    df = df[df["hora_inicial"].str.lower() != "nan"]
    df = df[df["hora_final"].str.strip() != ""]
    df = df[df["hora_final"].str.lower() != "nan"]

    # Remove linhas onde razao E origem estão vazios
    df = df[~((df["razao"].str.strip() == "") & (df["origem"].str.strip() == ""))]

    print(f"    📄 {len(df)} linha(s) válidas")
    return df


# ==============================
# UPSERT NO POSTGRESQL
# ==============================

def upsert_no_postgres(df, nome_usina):
    nome_tabela = nome_para_tabela(nome_usina)
    conn        = get_db_connection()
    cursor      = conn.cursor()

    try:
        criar_tabela_se_nao_existir(cursor, nome_tabela)

        linhas    = []
        ignoradas = 0

        for _, row in df.iterrows():
            try:
                data         = pd.to_datetime(row["data"], dayfirst=True).date()
                hora_inicial = str(row["hora_inicial"]).strip() or None
                hora_final   = str(row["hora_final"]).strip() or None
                razao        = str(row["razao"]).strip()
                origem       = str(row["origem"]).strip()
                descricao    = str(row["descricao"]).strip()

                val             = str(row["valor_limitacao"]).replace(",", ".").strip()
                valor_limitacao = float(val) if val and val.lower() not in ("nan", "") else None

                linhas.append((data, hora_inicial, hora_final, razao,
                               origem, valor_limitacao, descricao, datetime.now()))
            except Exception:
                ignoradas += 1
                continue

        if not linhas:
            print(f"    ⚠️  Nenhuma linha válida para inserir.")
            return

        sql = f"""
            INSERT INTO "{nome_tabela}"
                (data, hora_inicial, hora_final, razao, origem,
                 valor_limitacao, descricao, data_extracao)
            VALUES %s
            ON CONFLICT (data, hora_inicial, hora_final)
            DO UPDATE SET
                razao           = EXCLUDED.razao,
                origem          = EXCLUDED.origem,
                valor_limitacao = EXCLUDED.valor_limitacao,
                descricao       = EXCLUDED.descricao,
                data_extracao   = NOW();
        """

        execute_values(cursor, sql, linhas)
        conn.commit()

        print(f"    ✅ {len(linhas)} linha(s) salvas!")
        if ignoradas:
            print(f"    ⚠️  {ignoradas} linha(s) ignoradas")

    except Exception as e:
        conn.rollback()
        raise e
    finally:
        cursor.close()
        conn.close()


# ==============================
# FLUXO PRINCIPAL
# ==============================

def main():
    print("=" * 55)
    print("   CARGA HISTÓRICA — Excel Downloads → Supabase")
    print("=" * 55)

    print(f"\n📂 Pasta: {PASTA_DOWNLOADS}")

    # Valida conexão
    conn = get_db_connection()
    conn.close()
    print("✅ Supabase conectado!\n")

    # Busca todos os .xlsx dentro das subpastas de ano (ex: Downloads/2021/*.xlsx)
    arquivos = sorted(glob.glob(os.path.join(PASTA_DOWNLOADS, "*", "*.xlsx")))

    # Também inclui arquivos soltos direto na pasta Downloads
    arquivos += sorted(glob.glob(os.path.join(PASTA_DOWNLOADS, "*.xlsx")))

    if not arquivos:
        print("⚠️  Nenhum arquivo .xlsx encontrado!")
        print(f"   Estrutura esperada: Downloads/2021/FLS_2021.xlsx")
        return

    print(f"📋 {len(arquivos)} arquivo(s) encontrado(s)\n")

    total_ok     = 0
    total_erro   = 0
    nao_mapeados = []
    ano_atual    = None

    for caminho in arquivos:
        nome_arquivo = os.path.basename(caminho)
        pasta_pai    = os.path.basename(os.path.dirname(caminho))
        usina        = identificar_usina(nome_arquivo)

        # Imprime separador quando muda de ano
        if pasta_pai != ano_atual:
            ano_atual = pasta_pai
            print(f"\n{'─'*55}")
            print(f"  📅 {ano_atual}")
            print(f"{'─'*55}")

        if not usina:
            print(f"  ⏭️  {nome_arquivo}  →  sigla não reconhecida, ignorado.")
            nao_mapeados.append(nome_arquivo)
            continue

        print(f"  📊 {nome_arquivo}  →  {usina}")

        try:
            df = ler_excel(caminho)
            if df.empty:
                continue
            upsert_no_postgres(df, usina)
            total_ok += 1
        except Exception as e:
            print(f"    ❌ Erro: {e}")
            total_erro += 1

    print(f"\n{'='*55}")
    print(f"✅ Processados com sucesso : {total_ok}")
    if total_erro:
        print(f"❌ Com erro               : {total_erro}")
    if nao_mapeados:
        print(f"⏭️  Não mapeados           : {len(nao_mapeados)}")
        for n in nao_mapeados:
            print(f"   - {n}")
    print("🎉 Carga histórica finalizada!")


main()