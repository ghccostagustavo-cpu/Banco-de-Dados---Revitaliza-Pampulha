# =============================================================================
# IMPORTAÇÃO DE BIBLIOTECAS
# =============================================================================

from rapidfuzz import fuzz 
import sqlite3
import pandas as pd

# =============================================================================
# CONFIGURAÇÕES DE ACESSO
# =============================================================================
caminho_banco = r"caminho_banco_de_dados.db"
tabela_base = "nome_da_tabela" # Tabela que será atualizada (Alvo da correção)

def correcao_chaves_dif():
    """
    Compara a chave de endereço da tabela de baixas com a tabela base.
    Se as matrículas forem iguais mas as chaves forem diferentes, 
    o script avalia a similaridade e corrige a tabela base.
    """
    conexao = sqlite3.connect(caminho_banco)
    print("Iniciando limpeza...")

    cursor = conexao.cursor()

    print("\nLendo tabelas...")
    # Carrega os dados do SQLite para DataFrames Pandas
    df_baixas = pd.read_sql_query("SELECT * FROM baixas", conexao)
    df_base = pd.read_sql_query(f"SELECT * FROM [{tabela_base}]", conexao)

    print(f"Baixas: {len(df_baixas)} linhas")
    print(f"Base Alvo: {len(df_base)} linhas")

    # =============================================================================
    # TRATAMENTO DE MATRÍCULAS
    # =============================================================================
    # Garante que a matrícula seja string e remove decimais (ex: '123.0' vira '123')
    # Isso é essencial para que o merge entre as tabelas não falhe por tipo de dado
    df_baixas["Matrícula"] = df_baixas["Matrícula"].astype(str).str.split('.').str[0]
    df_base["MATRÍCULA"] = df_base["MATRÍCULA"].astype(str).str.split('.').str[0]

    # =============================================================================
    # CRUZAMENTO DE DADOS (MERGE)
    # =============================================================================
    # Faz um INNER JOIN usando a matrícula como chave de ligação
    df_merge = df_baixas.merge(
        df_base,
        left_on="Matrícula",
        right_on="MATRÍCULA",
        suffixes=('_certa', '_errada') # Identifica de onde vem cada 'chave'
    )

    # Filtra apenas as linhas onde as matrículas batem, mas as chaves de endereço são diferentes
    df_erros = df_merge[df_merge['chave_certa'] != df_merge['chave_errada']]

    if df_erros.empty:
        print("Chaves condizentes, nada a se corrigir.")
        conexao.close()
        return

    print(f"{len(df_erros)} divergências encontradas. Analisando e corrigindo...")

    correcoes = 0

    # =============================================================================
    # ANÁLISE DE SIMILARIDADE E ATUALIZAÇÃO
    # =============================================================================
    for _, linha in df_erros.iterrows():
        # O fuzz.ratio calcula a distância de Levenshtein (similaridade entre 0 e 100)
        # Ex: "RUA A" vs "RUA B" terá um ratio alto, mas "RUA A" vs "AVENIDA X" terá baixo
        ratio = fuzz.ratio(linha['chave_certa'], linha['chave_errada'])

        # Se a similaridade for maior ou igual a 70%, assume-se que é o mesmo lugar
        # com algum erro de digitação ou formatação, e aplica a correção.
        if ratio >= 70:
            cursor.execute(
                f'UPDATE [{tabela_base}] SET chave = ? WHERE "MATRÍCULA" = ?',
                (linha['chave_certa'], linha['MATRÍCULA'])
            )
            correcoes += 1

    # Salva as alterações no banco de dados e encerra a conexão
    conexao.commit()
    conexao.close()
    print(f"{correcoes} correções executadas com sucesso.")

if __name__ == "__main__":
    correcao_chaves_dif()