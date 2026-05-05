import pandas as pd
from sqlalchemy import create_engine, text

# Conexões
engine_sqlite = create_engine(r"sqlite:///CAMINHO/DO/BANCO.db")
engine_pg = create_engine("postgresql://USUARIO:SENHA@localhost:1234/BANCO")

# Busca todas as tabelas do SQLite
query_tabelas = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';"

with engine_sqlite.connect() as conn:
    tabelas_df = pd.read_sql(text(query_tabelas), conn)

# Migração tabelas
for nome_tabela in tabelas_df['name']:
    print(f"Lendo: {nome_tabela}")

    df = pd.read_sql(f'SELECT * FROM "{nome_tabela}"', engine_sqlite)

    nome_pg = nome_tabela.lower()

    df.to_sql(nome_pg, engine_pg, if_exists='replace', index=False)
    print(f"  -> {len(df)} linhas enviadas para '{nome_pg}'")

print("Migração finalizada.")