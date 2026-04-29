import pandas as pd
import sqlite3
from datetime import datetime
import unicodedata
import re

#caminhos das baixas 
caminho_baixas = r"caminho_planilha_baixas.xlsm"
tabela_baixas = "tabela_baixas"
aba_baixas = "BAIXAS"

#caminhos da base de adesao
caminho_planilha_adesao = r"caminho_planilha_base_adesao.xlsx"
tabela_adesao = "Tabela_adesao_banco"
aba_adesao = "Aba_adesao"

#caminho Serviços Esgoto
caminho_caca = r"caminho_planilha_caca.xlsx"
tabela_caca = "Caca_Esgoto_banco"
aba_caca = "Esgoto_Serviços"

#caminho Cadastro
caminho_cadastro = r"caminho_planilha_cadastro.xlsx"
tabela_cadastro = "cadastro"
aba_cadastro = "aba_cadastro"

#caminho banco e log
caminho_banco = r"caminho_banco_dados.db"
caminho_log = r"caminho_logs.txt"

#caminho baixas secundário
caminho_baixas_secundario = r"caminho_planilha_baixas_secundario.xlsx"
tabela_baixas_secundario = "tabela_baixas_secundario"
aba_baixas_secundario = "BAIXAS_SECUNDARIO"




def registrar_log(mensagem):
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{agora}] {mensagem}\n"
    with open(caminho_log, "a", encoding="ISO-8859-1") as f:
        f.write(linha)
    print(linha.strip())

def limpar(texto):
    if pd.isna(texto) or str(texto).strip() == "" or str(texto).upper() == "NAN":
        return ""
    t = str(texto).upper().replace("R ", "RUA")
    t = unicodedata.normalize('NFKD', t).encode('ASCII', 'ignore').decode('utf-8')
    return "".join(filter(str.isalnum, t))

def compl(texto):
    if pd.isna(texto) or str(texto).strip() == "" or str(texto).upper() == "NAN":
        return "CASA"
    c = str(texto).upper().strip()
    if c == "CA":
        return "CASA"
    c = c.replace("CA ", "CASA ")
    return "".join(filter(str.isalnum, c))

def matricula_vazia_nova(texto):
    if not texto:
        return ""
    m = str(texto).upper().replace('NOVA', 'NOVO')
    return "".join(filter(str.isalnum, m))

def apagar_vazias(df):
    colunas_apagar = [col for col in df.columns if 'Unnamed:' in str(col)]
    return df.drop(columns=colunas_apagar)

def criar_chave_composta(texto):
    a = str(texto).strip().upper()

    a = re.sub(r'\d{5}-?\d{3}', '', a)                          
    compl_match = re.search(r'\(([^)]+)\)', a)
    compl_val = compl_match.group(1).strip() if compl_match else "CASA"
    a = re.sub(r'\([^)]*\)', '', a)                              

    a = a.replace("NOME_DA_CIDADE", "").replace("- UF", "").replace(",UF", "")

    partes = a.split(" - ", 1)
    endereco = partes[0].strip()
    bairro   = partes[1].strip() if len(partes) > 1 else ""

    partes_end = endereco.split(",", 1)
    logradouro = partes_end[0].strip()
    numero     = partes_end[1].strip() if len(partes_end) > 1 else ""

    logradouro = logradouro.replace("R. ", "RUA").replace("R.", "RUA")
    logradouro = logradouro.replace("BC. ", "BC").replace("BC.", "BC")
    logradouro = logradouro.replace("AV. ", "AV").replace("AV.", "AV")

    compl_val = compl_val.replace("CA ", "CASA")

    logradouro = "".join(filter(str.isalnum, logradouro))
    numero     = "".join(filter(str.isalnum, numero))
    compl_val  = "".join(filter(str.isalnum, compl_val))
    bairro     = "".join(filter(str.isalnum, bairro))

    return f"{logradouro}|{numero}|{compl_val}|{bairro}"

def sincronizar_dados_baixas_geral():
    try:
        registrar_log("Lendo baixas gerais...")
        df = pd.read_excel(caminho_baixas, sheet_name=aba_baixas, engine='openpyxl')

        df['chave'] = df['Endereço'].apply(criar_chave_composta)

        registrar_log("Conectando ao banco SQLite...")
        conexao = sqlite3.connect(caminho_banco)
        df.to_sql(tabela_baixas, conexao, if_exists='replace', index=False)
        conexao.close()

        registrar_log(f"SUCESSO: {len(df)} linhas sincronizadas para '{tabela_baixas}'.")

    except Exception as e:
        registrar_log(f"ERRO CRÍTICO: {str(e)}")

def sincronizar_dados_adesao():
    try:
        registrar_log("Lendo adesão...")
        df = pd.read_excel(caminho_planilha_adesao, sheet_name=aba_adesao, engine='openpyxl')

        df = apagar_vazias(df)
        df['MATRÍCULA']    = df['MATRÍCULA'].apply(matricula_vazia_nova)
        df['COMPLEMENTO:'] = df['COMPLEMENTO:'].apply(compl)

        df['chave'] = (
            df['LOGRADORO:'].apply(limpar) + "|" +
            df['NUMERO:'].apply(limpar)    + "|" +
            df['COMPLEMENTO:'].apply(compl) + "|" +
            df['BAIRRO'].apply(limpar)
        )

        df = df.drop_duplicates()

        registrar_log("Conectando ao banco SQLite...")
        conexao = sqlite3.connect(caminho_banco)
        df.to_sql(tabela_adesao, conexao, if_exists='replace', index=False)
        conexao.close()

        registrar_log(f"SUCESSO: {len(df)} linhas sincronizadas para '{tabela_adesao}'.")

    except Exception as e:
        registrar_log(f"ERRO: {str(e)}")

def sincronizar_dados_baixas_secundario():
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


if __name__ == "__main__":
    sincronizar_dados_baixas_geral()
    sincronizar_dados_adesao()
    sincronizar_dados_baixas_secundario()
    sincronizar_dados_caca_esgoto()

    conexao = sqlite3.connect(caminho_banco)
    print(pd.read_sql_query(f'SELECT * FROM {tabela_adesao};', conexao))
    conexao.close()
