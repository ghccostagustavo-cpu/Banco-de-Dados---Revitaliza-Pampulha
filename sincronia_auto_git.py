import pandas as pd
import sqlite3
from datetime import datetime
import os

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




def registrar_log(mensagem):


    agora = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    linha = f"[{agora}] {mensagem}\n"

    with open(caminho_log, "a", encoding="ISO-8859-1") as arquivo_log:

        arquivo_log.write(linha)

    print(linha.strip())


def sincronizar_dados_tbbaixas(): 
    try:
        registrar_log("Lendo baixas G...")

        df_excel = pd.read_excel(caminho_baixas, sheet_name=aba_baixas, engine='openpyxl')

        registrar_log("Conectando ao banco SQLite...")

        conexao = sqlite3.connect(caminho_banco)

        df_excel.to_sql(tabela_baixas, conexao, if_exists='replace', index=False)

        conexao.close()

        registrar_log(f"SUCESSO: {len(df_excel)} linhas sincronizadas para a tabela '{tabela_baixas}'.")


    except Exception as e:

        registrar_log(f"ERRO CRÍTICO: {str(e)}")




if __name__ == "__main__":

    sincronizar_dados_tbbaixas()



def sincronizar_dados_baseadesao():
    try:
        registrar_log("Lendo adesao SU...") 
        df_excel = pd.read_excel(caminho_planilha_adesao, sheet_name=aba_adesao, engine='openpyxl')

        registrar_log("Conectando ao banco SQLite...")
        conexao = sqlite3.connect(caminho_banco)

        df_excel.to_sql(tabela_adesao, conexao, if_exists='replace', index=False)

        conexao.close()

        registrar_log(f"SUCESSO: {len(df_excel)} linhas sincronizadas para a tabela '{tabela_adesao}'.")

    except Exception as e:
        registrar_log(f"ERRO: {str(e)}")

if __name__ == "__main__":
    sincronizar_dados_baseadesao()


def sincronizar_dados_esgoto():

    try:
        registrar_log("Lendo caça esgoto...") 

        df_excel = pd.read_excel(caminho_caca, sheet_name=aba_caca, engine='openpyxl')

        registrar_log("Conectando ao banco SQLite...")

        conexao = sqlite3.connect(caminho_banco)

        df_excel.to_sql(tabela_caca, conexao, if_exists='replace', index=False)

        conexao.close()

        registrar_log(f"SUCESSO: {len(df_excel)} linhas sincronizadas para a tabela '{tabela_caca}'.")

    except Exception as e:
        registrar_log(f"ERRO: {str(e)}")

if __name__ == "__main__":
    sincronizar_dados_esgoto()

