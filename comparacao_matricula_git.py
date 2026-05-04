# =============================================================================
# IMPORTAÇÃO DE BIBLIOTECAS 
# =============================================================================

import pandas as pd
import sqlite3
import unicodedata

# =============================================================================
# CONFIGURAÇÕES DE CAMINHOS E CONEXÃO
# =============================================================================
pasta_excel = r'caminho_pasta_excel_destino.xlsx'
cnc = sqlite3.connect(r'caminho_banco_dados.db')

# Carrega as duas tabelas do banco para comparação em memória (DataFrames)
df_caça = pd.read_sql_query('SELECT * FROM EsgotoCampo', cnc)
df_adesao = pd.read_sql_query('SELECT * FROM BaseAdesao', cnc)

# =============================================================================
# FUNÇÕES DE TRATAMENTO (PADRONIZAÇÃO)
# =============================================================================

def compl(texto):
    """Padroniza o complemento: trata nulos e abreviações de 'CASA'."""
    if pd.isna(texto) or str(texto).strip() == "" or str(texto).upper() == "NAN": 
        return "CASA"
    c = str(texto).upper().strip()
    if c == "CA":
        return "CASA"
    c = c.replace("CA ", "CASA ")
    return "".join(filter(str.isalnum, c))

def limpar(texto):
    """Limpeza geral: remove acentos, padroniza 'RUA' e remove pontuação."""
    if pd.isna(texto) or str(texto).strip() == "" or str(texto).upper() == "NAN": 
        return ""
    t = str(texto).upper().replace("R ", "RUA")
    # Normalização para remover acentos (Ex: 'São' -> 'Sao')
    t = unicodedata.normalize('NFKD', t).encode('ASCII', 'ignore').decode('utf-8')
    return "".join(filter(str.isalnum, t))

def matricula_vazia_novA(texto):
    """Padroniza o termo 'NOVO' em matrículas ainda não geradas."""
    if not texto: return ""
    m = str(texto).upper().replace('NOVA', 'NOVO')
    return "".join(filter(str.isalnum, m))

# =============================================================================
# APLICAÇÃO DO TRATAMENTO E CRIAÇÃO DE CHAVES
# =============================================================================

# Aplica a limpeza nas colunas de matrícula e complemento em ambos os DataFrames
df_caça['Matricula'] = df_caça['Matricula'].apply(matricula_vazia_novA)
df_adesao['MATRÍCULA'] = df_adesao['MATRÍCULA'].apply(matricula_vazia_novA)
df_caça['Complemento'] = df_caça['Complemento'].apply(compl)
df_adesao['COMPLEMENTO:'] = df_adesao['COMPLEMENTO:'].apply(compl)

# Gera a 'chave' única na tabela de campo
df_caça['chave'] = (
    df_caça['Logradouro'].apply(limpar) + "|" +
    df_caça['Num'].apply(limpar) + "|" +
    df_caça['Complemento'].apply(compl)+ "|" +
    df_caça['Bairro'].apply(limpar)
)

# Gera a 'chave' única na tabela de adesão (base oficial)
df_adesao['chave'] = (
    df_adesao['LOGRADORO:'].apply(limpar)+ "|" + 
    df_adesao['NUMERO:'].apply(limpar) + "|" + 
    df_adesao['COMPLEMENTO:'].apply(compl) + "|" + 
    df_adesao['BAIRRO'].apply(limpar)
)

# =============================================================================
# CRUZAMENTO (JOIN) E FINALIZAÇÃO
# =============================================================================

# Realiza um LEFT JOIN: mantém tudo da tabela de campo (caça) e traz a matrícula da adesão
# O parâmetro 'indicator=True' cria uma coluna '_merge' que diz se a chave existe nos dois ou só na esquerda
resultado = pd.merge(
    df_caça, 
    df_adesao[['chave', 'MATRÍCULA']], on='chave', how='left', indicator=True)

# Trata valores nulos após o cruzamento (onde não houve correspondência)
resultado['Matricula'] = resultado['Matricula'].fillna('NOVO').replace("NOVA", 'NOVO')
resultado['MATRÍCULA'] = resultado['MATRÍCULA'].fillna('NOVO')

# Remove possíveis duplicatas geradas por erros na base de origem
resultado = resultado.drop_duplicates()

# Salva o resultado final em dois destinos: Excel para a equipe e SQLite para o sistema
resultado.to_excel(pasta_excel, index=False)
resultado.to_sql('Tabela_destino_banco', cnc, if_exists='replace', index=False)

print("Comparação concluída. Resultado salvo em Excel e banco de dados.")