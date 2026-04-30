from rapidfuzz import fuzz 
import sqlite3
import pandas as pd

caminho_banco = r"caminho_banco_de_dados.db"
tabela_base = "nome_da_tabela"

def correcao_chaves_dif():
    conexao = sqlite3.connect(caminho_banco)
    print("Iniciando limpeza...")

    cursor = conexao.cursor()

    print("\nLendo tabelas...")

    df_baixas = pd.read_sql_query("SELECT * FROM baixas", conexao)
    df_base = pd.read_sql_query(f"SELECT * FROM [{tabela_base}]", conexao)

    print(f"gusigos: {len(df_baixas)} linhas")
    print(f"Adesão Tratada: {len(df_base)} linhas")

    df_baixas["Matrícula"] = df_baixas["Matrícula"].astype(str).str.split('.').str[0]
    df_base["MATRÍCULA"] = df_base["MATRÍCULA"].astype(str).str.split('.').str[0]

    df_merge = df_baixas.merge(
        df_base,
        left_on="Matrícula",
        right_on="MATRÍCULA",
        suffixes=('_certa', '_errada')
    )

    df_erros = df_merge[df_merge['chave_certa'] != df_merge['chave_errada']]

    if df_erros.empty:
        print("Chaves condizentes, nada a se corrigir.")
        conexao.close()
        return

    print(f"{len(df_erros)} divergências encontradas. Analisando e corrigindo...")

    correcoes = 0

    for _, linha in df_erros.iterrows():
        ratio = fuzz.ratio(linha['chave_certa'], linha['chave_errada'])

        if ratio >= 70:
            cursor.execute(
                f'UPDATE [{tabela_base}] SET chave = ? WHERE "MATRÍCULA" = ?',
                (linha['chave_certa'], linha['MATRÍCULA'])
            )
            correcoes += 1

    conexao.commit()
    conexao.close()
    print(f"{correcoes} correções executadas com sucesso.")

if __name__ == "__main__":
    correcao_chaves_dif()