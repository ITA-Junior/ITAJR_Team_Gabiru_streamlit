"""
Transferência de dados do Excel para o Supabase
Arquivo: seus_dados.xls
Sheets: Orders, Returns, People

Dependências:
    pip install pandas openpyxl supabase

Configuração:
    Defina as variáveis SUPABASE_URL e SUPABASE_KEY abaixo,
    ou exporte-as como variáveis de ambiente antes de rodar o script.
"""

import os
import math
try:
    import pandas as pd
except ImportError:
    print("pandas is not installed. Please install it with 'pip install pandas'")
    exit(1)
from supabase import create_client, Client

# ──────────────────────────────────────────────
# CONFIGURAÇÃO — edite aqui ou use variáveis de ambiente
# ──────────────────────────────────────────────
SUPABASE_URL = os.getenv("SUPABASE_URL", "https://yzcesukcjnrpyywthlxn.supabase.co")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl6Y2VzdWtjam5ycHl5d3RobHhuIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3ODQ2OTU4NSwiZXhwIjoyMDk0MDQ1NTg1fQ.NA1jYmS1aLihrNdtimAX_EW7rPOECzTlWwlZLINmU_Y")

# Caminho do arquivo Excel
EXCEL_PATH = "seus_dados.xls"

# Tamanho do lote para inserção (Supabase aceita até ~1000 linhas por vez)
BATCH_SIZE = 500

# Mapeamento: nome da aba → nome da tabela no Supabase
SHEET_TABLE_MAP = {
    "Orders":  "orders",
    "Returns": "returns",
    "People":  "people",
}


# ──────────────────────────────────────────────
# FUNÇÕES AUXILIARES
# ──────────────────────────────────────────────

def sanitize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converte tipos problemáticos para formatos aceitos pelo Supabase (JSON):
      - datetime → string ISO 8601
      - NaN / NaT → None
      - int64 numpy → int nativo Python
    """
    df = df.copy()

    # Substitui NaN/NaT por None (serializa como null em JSON)
    df = df.where(pd.notnull(df), other=None)

    for col in df.columns:
        # Datas → string ISO
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].apply(
                lambda v: v.isoformat() if v is not None and not _is_nan(v) else None
            )
        # Floats que sejam NaN explícito
        elif pd.api.types.is_float_dtype(df[col]):
            df[col] = df[col].apply(lambda v: None if v is None or _is_nan(v) else float(v))

    return df


def _is_nan(value) -> bool:
    try:
        return math.isnan(value)
    except (TypeError, ValueError):
        return False


def dataframe_to_records(df: pd.DataFrame) -> list[dict]:
    """Converte DataFrame em lista de dicts com nomes de colunas em snake_case."""
    df = sanitize_dataframe(df)
    df.columns = [col.strip().lower().replace(" ", "_").replace("-", "_") for col in df.columns]
    return df.to_dict(orient="records")


def insert_in_batches(supabase: Client, table: str, records: list[dict]) -> None:
    """Insere registros em lotes para respeitar limites da API do Supabase."""
    total = len(records)
    n_batches = math.ceil(total / BATCH_SIZE)

    print(f"  → {total} registros | {n_batches} lote(s) de até {BATCH_SIZE}")

    for i in range(n_batches):
        batch = records[i * BATCH_SIZE : (i + 1) * BATCH_SIZE]
        response = supabase.table(table).insert(batch).execute()

        # A lib python-supabase levanta exceção automaticamente em caso de erro,
        # mas checamos o retorno mesmo assim.
        if hasattr(response, "error") and response.error:
            raise RuntimeError(f"Erro ao inserir lote {i+1} em '{table}': {response.error}")

        print(f"    Lote {i+1}/{n_batches} inserido ({len(batch)} linhas)")


# ──────────────────────────────────────────────
# TABELAS SQL — crie estas tabelas no Supabase antes de rodar
# ──────────────────────────────────────────────
#
# -- ORDERS
# create table orders (
#   row_id        int,
#   order_id      text,
#   order_date    text,
#   ship_date     text,
#   ship_mode     text,
#   customer_id   text,
#   customer_name text,
#   segment       text,
#   country       text,
#   city          text,
#   state         text,
#   postal_code   int,
#   region        text,
#   product_id    text,
#   category      text,
#   sub_category  text,
#   product_name  text,
#   sales         numeric,
#   quantity      int,
#   discount      numeric,
#   profit        numeric
# );
#
# -- RETURNS
# create table returns (
#   returned  text,
#   order_id  text
# );
#
# -- PEOPLE
# create table people (
#   person  text,
#   region  text
# );


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    # Conecta ao Supabase
    print("Conectando ao Supabase...")
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    print("Conexão estabelecida.\n")

    # Lê o arquivo Excel
    print(f"Lendo arquivo: {EXCEL_PATH}")
    # .xls (formato legado) exige xlrd; .xlsx usa openpyxl
    engine = "xlrd" if EXCEL_PATH.lower().endswith(".xls") else "openpyxl"
    excel_file = pd.ExcelFile(EXCEL_PATH, engine=engine)
    print(f"Abas encontradas: {excel_file.sheet_names}\n")

    # Processa cada aba
    for sheet_name, table_name in SHEET_TABLE_MAP.items():
        if sheet_name not in excel_file.sheet_names:
            print(f"[AVISO] Aba '{sheet_name}' não encontrada, pulando...\n")
            continue

        print(f"Processando aba '{sheet_name}' → tabela '{table_name}'")
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        records = dataframe_to_records(df)

        insert_in_batches(supabase, table_name, records)
        print(f"  ✓ Aba '{sheet_name}' transferida com sucesso!\n")

    print("✅ Transferência concluída!")


if __name__ == "__main__":
    main()