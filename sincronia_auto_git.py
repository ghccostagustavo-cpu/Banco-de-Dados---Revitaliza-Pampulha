# =============================================================================
# IMPORTAÇÃO DE BIBLIOTECAS
# =============================================================================

import pandas as pd
import sqlite3
from datetime import datetime
import unicodedata
import re

# =============================================================================
# CONFIGURAÇÕES DE CAMINHOS E TABELAS
# =============================================================================

# Caminhos das Baixas Gerais
caminho_baixas = r"caminho_planilha_baixas.xlsm"
tabela_baixas = "tabela_baixas"
aba_baixas = "BAIXAS"

# Caminhos da Base de Adesão
caminho_planilha_adesao = r"caminho_planilha_base_adesao.xlsx"
tabela_adesao = "Tabela_adesao_banco"
aba_adesao = "Aba_adesao"

# Caminhos de Serviços Esgoto
caminho_caca = r"caminho_planilha_caca.xlsx"
tabela_caca = "Caca_Esgoto_banco"
aba_caca = "Esgoto_Serviços"

# Caminhos do Banco de Dados SQLite e Arquivo de Log
caminho_banco = r"caminho_banco_dados.db"
caminho_log = r"caminho_logs.txt"

# Caminhos das Baixas Secundárias
caminho_baixas_secundario = r"caminho_planilha_baixas_secundario.xlsx"
tabela_baixas_secundario = "tabela_baixas_secundario"
aba_baixas_secundario = "BAIXAS_SECUNDARIO"


# =============================================================================
# FUNÇÕES AUXILIARES DE LOG E LIMPEZA DE DADOS
# =============================================================================

def registrar_log(mensagem):
    """
    Gera um log com data e hora da execução e salva em um arquivo de texto.
    Também imprime a mensagem no console para acompanhamento em tempo real.
    """
    agora = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    linha = f"[{agora}] {mensagem}\n"
    
    # Abre o arquivo em modo 'a' (append) para adicionar no final sem apagar o histórico.
    # O encoding ISO-8859-1 garante que caracteres especiais em português não quebrem o log.
    with open(caminho_log, "a", encoding="ISO-8859-1") as f:
        f.write(linha)
    print(linha.strip())

def limpar(texto):
    """
    Padroniza strings: remove valores nulos, substitui abreviações comuns, 
    remove acentuação e mantém apenas caracteres alfanuméricos.
    """
    # Tratamento de valores vazios ou nulos trazidos pelo Pandas
    if pd.isna(texto) or str(texto).strip() == "" or str(texto).upper() == "NAN":
        return ""
    
    # Substitui "R " por "RUA" para padronizar logradouros
    t = str(texto).upper().replace("R ", "RUA")
    
    # Remove acentuação usando normalização Unicode
    t = unicodedata.normalize('NFKD', t).encode('ASCII', 'ignore').decode('utf-8')
    
    # Filtra e retorna apenas o que for letra ou número (remove espaços, hifens, vírgulas)
    return "".join(filter(str.isalnum, t))

def compl(texto):
    """
    Trata especificamente o campo de complemento do endereço.
    Se estiver vazio, assume o padrão "CASA".
    """
    if pd.isna(texto) or str(texto).strip() == "" or str(texto).upper() == "NAN":
        return "CASA"
    
    c = str(texto).upper().strip()
    
    # Expande abreviações comuns de "CASA"
    if c == "CA":
        return "CASA"
    c = c.replace("CA ", "CASA ")
    
    # Retorna apenas caracteres alfanuméricos
    return "".join(filter(str.isalnum, c))

def matricula_vazia_nova(texto):
    """
    Garante a padronização do campo de matrícula, alterando "NOVA" para "NOVO".
    """
    if not texto:
        return ""
    m = str(texto).upper().replace('NOVA', 'NOVO')
    return "".join(filter(str.isalnum, m))

def apagar_vazias(df):
    """
    Remove colunas que o Pandas cria automaticamente (ex: 'Unnamed: 0') 
    quando o Excel possui colunas vazias formatadas.
    """
    colunas_apagar = [col for col in df.columns if 'Unnamed:' in str(col)]
    return df.drop(columns=colunas_apagar)

def criar_chave_composta(texto):
    """
    Quebra uma string de endereço completo e extrai seus componentes 
    (logradouro, número, complemento e bairro) para criar uma chave única de cruzamento.
    
    Formato esperado de saída: "LOGRADOURO|NUMERO|COMPLEMENTO|BAIRRO"
    """
    a = str(texto).strip().upper()

    # 1. Remove CEP (qualquer sequência de 5 dígitos, com ou sem hífen, seguidos de 3 dígitos)
    a = re.sub(r'\d{5}-?\d{3}', '', a)                          
    
    # 2. Extrai o complemento que geralmente vem entre parênteses
    compl_match = re.search(r'\(([^)]+)\)', a)
    compl_val = compl_match.group(1).strip() if compl_match else "CASA"
    
    # Remove os parênteses e seu conteúdo da string original para não sujar o logradouro
    a = re.sub(r'\([^)]*\)', '', a)                              

    # 3. Remove informações de cidade e estado genéricas
    a = a.replace("NOME_DA_CIDADE", "").replace("- UF", "").replace(",UF", "")

    # 4. Separa o Bairro do Endereço (geralmente divididos por " - ")
    partes = a.split(" - ", 1)
    endereco = partes[0].strip()
    bairro   = partes[1].strip() if len(partes) > 1 else ""

    # 5. Separa o Logradouro do Número (geralmente divididos por vírgula)
    partes_end = endereco.split(",", 1)
    logradouro = partes_end[0].strip()
    numero     = partes_end[1].strip() if len(partes_end) > 1 else ""

    # 6. Padroniza prefixos de logradouro
    logradouro = logradouro.replace("R. ", "RUA").replace("R.", "RUA")
    logradouro = logradouro.replace("BC. ", "BC").replace("BC.", "BC")
    logradouro = logradouro.replace("AV. ", "AV").replace("AV.", "AV")

    compl_val = compl_val.replace("CA ", "CASA")

    # 7. Limpa todos os componentes deixando apenas letras e números
    logradouro = "".join(filter(str.isalnum, logradouro))
    numero     = "".join(filter(str.isalnum, numero))
    compl_val  = "".join(filter(str.isalnum, compl_val))
    bairro     = "".join(filter(str.isalnum, bairro))

    # 8. Monta a chave final separada por pipes (|)
    return f"{logradouro}|{numero}|{compl_val}|{bairro}"


# =============================================================================
# FUNÇÕES PRINCIPAIS DE SINCRONIZAÇÃO (ETL)
# =============================================================================

def sincronizar_dados_baixas_geral():
    """
    Lê a planilha de baixas gerais, processa o campo de endereço para criar 
    uma chave composta e salva os dados no banco SQLite.
    """
    try:
        registrar_log("Lendo baixas gerais...")
        # openpyxl é especificado como engine para garantir compatibilidade com .xlsm/.xlsx
        df = pd.read_excel(caminho_baixas, sheet_name=aba_baixas, engine='openpyxl')

        # Aplica a função complexa de regex para derivar a chave a partir do endereço completo
        df['chave'] = df['Endereço'].apply(criar_chave_composta)

        registrar_log("Conectando ao banco SQLite...")
        conexao = sqlite3.connect(caminho_banco)
        # if_exists='replace' fará o 'drop' da tabela se ela já existir e a recriará
        df.to_sql(tabela_baixas, conexao, if_exists='replace', index=False)
        conexao.close()

        registrar_log(f"SUCESSO: {len(df)} linhas sincronizadas para '{tabela_baixas}'.")

    except Exception as e:
        registrar_log(f"ERRO CRÍTICO: {str(e)}")

def sincronizar_dados_adesao():
    """
    Lê a planilha de adesão, realiza limpeza profunda de colunas e textos 
    (matrícula, logradouro, número, complemento, bairro), remove duplicadas e salva no banco.
    """
    try:
        registrar_log("Lendo adesão...")
        df = pd.read_excel(caminho_planilha_adesao, sheet_name=aba_adesao, engine='openpyxl')

        # Limpeza estrutural e de campos específicos
        df = apagar_vazias(df)
        df['MATRÍCULA']    = df['MATRÍCULA'].apply(matricula_vazia_nova)
        df['COMPLEMENTO:'] = df['COMPLEMENTO:'].apply(compl)

        # Concatena os campos isolados da planilha para formar a chave padrão
        df['chave'] = (
            df['LOGRADORO:'].apply(limpar) + "|" +
            df['NUMERO:'].apply(limpar)    + "|" +
            df['COMPLEMENTO:'].apply(compl) + "|" +
            df['BAIRRO'].apply(limpar)
        )

        # Remove linhas inteiras que sejam cópias exatas para evitar redundância no banco
        df = df.drop_duplicates()

        registrar_log("Conectando ao banco SQLite...")
        conexao = sqlite3.connect(caminho_banco)
        df.to_sql(tabela_adesao, conexao, if_exists='replace', index=False)
        conexao.close()

        registrar_log(f"SUCESSO: {len(df)} linhas sincronizadas para '{tabela_adesao}'.")

    except Exception as e:
        registrar_log(f"ERRO: {str(e)}")

def sincronizar_dados_baixas_secundario():
    """
    Lê a planilha de baixas secundárias e realiza uma carga direta para o banco SQLite
    (sem tratamento de dados no DataFrame).
    """
    try:
        registrar_log("Lendo baixas secundárias...")
        df = pd.read_excel(caminho_baixas_secundario, sheet_name=aba_baixas_secundario, engine='openpyxl')

        registrar_log("Conectando ao banco SQLite...")
        conexao = sqlite3.connect(caminho_banco)
        df.to_sql(tabela_baixas_secundario, conexao, if_exists='replace', index=False)
        conexao.close()

        registrar_log(f"SUCESSO: {len(df)} linhas sincronizadas para '{tabela_baixas_secundario}'.")

    except Exception as e:
        registrar_log(f"ERRO: {str(e)}")

def sincronizar_dados_caca_esgoto():
    """
    Lê a base de serviços de esgoto e carrega diretamente para o banco SQLite.
    """
    try:
        registrar_log("Lendo caça esgoto...")
        df = pd.read_excel(caminho_caca, sheet_name=aba_caca, engine='openpyxl')

        registrar_log("Conectando ao banco SQLite...")
        conexao = sqlite3.connect(caminho_banco)
        df.to_sql(tabela_caca, conexao, if_exists='replace', index=False)
        conexao.close()

        registrar_log(f"SUCESSO: {len(df)} linhas sincronizadas para '{tabela_caca}'.")

    except Exception as e:
        registrar_log(f"ERRO: {str(e)}")


# =============================================================================
# FLUXO DE EXECUÇÃO PRINCIPAL
# =============================================================================
if __name__ == "__main__":
    sincronizar_dados_baixas_geral()
    sincronizar_dados_adesao()
    sincronizar_dados_baixas_secundario()
    sincronizar_dados_caca_esgoto()