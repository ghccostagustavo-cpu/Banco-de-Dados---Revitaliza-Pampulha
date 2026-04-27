import pandas as pd
import sqlite3
from datetime import datetime

# Caminho Cadastro
caminho_cadastro = r"caminho\para\cadastro.xlsx"
tabela_cadastro = "cadastro"
aba_cadastro = "aba_cadastro"

# Caminho banco e log
caminho_banco = r"caminho_banco.db"
caminho_log = r"caminho_log.txt"


def registrar_log(msg):
    agora = datetime.now().strftime("%d-%m-%Y, %H:%M:%S")
    linha = f"[{agora}] {msg}\n"
    with open(caminho_log, "a", encoding="ISO-8859-1") as arquivo_log:
        arquivo_log.write(linha)
    print(linha.strip())


def limpar_converter(df):
    from openpyxl.utils import get_column_letter, column_index_from_string
    import openpyxl
    import re

    def letras_range(inicio, fim):
        return [get_column_letter(i) for i in range(
            column_index_from_string(inicio),
            column_index_from_string(fim) + 1
        )]

    excluir_colunas = set(
        letras_range('A', 'C') +
        letras_range('H', 'H') +
        letras_range('N', 'N') +
        letras_range('P', 'X') +
        letras_range('Z', 'AD') +
        letras_range('AF', 'FK') +
        letras_range('FN', 'FO')
    )

    wb = openpyxl.load_workbook(caminho_cadastro, read_only=True)
    ws = wb[aba_cadastro]
    cabecalho = {cell.column_letter: cell.value for cell in list(ws.rows)[5]}
    wb.close()

    nomes_excluir = [
        cabecalho[l] for l in excluir_colunas
        if l in cabecalho and cabecalho[l]
    ]

    registrar_log(f"Excluindo {len(nomes_excluir)} colunas desnecessárias...")

    df = df.drop(columns=[col for col in nomes_excluir if col in df.columns])

    def dms_para_decimal(valor):
        try:
            nums = re.findall(r'[\d,.]+', str(valor))
            graus = float(nums[0])
            minutos = float(nums[1])
            segundos = float(nums[2].replace(',', '.'))
            decimal = graus + minutos / 60 + segundos / 3600
            if 'S' in str(valor) or 'O' in str(valor) or 'W' in str(valor):
                decimal = -decimal
            return round(decimal, 7)
        except:
            return None

    df['Latitude_decimal'] = df['Latitude'].apply(dms_para_decimal)
    df['Longitude_decimal'] = df['Longitude'].apply(dms_para_decimal)
    df['Coordenadas'] = 'POINT(' + df['Longitude_decimal'].astype(str) + ' ' + df['Latitude_decimal'].astype(str) + ')'

    return df


def sincronizar_dados_cadastro():
    try:
        registrar_log("Lendo cadastro...")
        df = pd.read_excel(caminho_cadastro, sheet_name=aba_cadastro, header=5, engine='openpyxl')

        registrar_log("Limpando e convertendo dados...")
        df = limpar_converter(df)

        registrar_log("Conectando ao banco SQLite...")
        conexao = sqlite3.connect(caminho_banco)
        df.to_sql(tabela_cadastro, conexao, if_exists='replace', index=False)
        conexao.close()

        registrar_log(f"SUCESSO: {len(df)} linhas sincronizadas para a tabela '{tabela_cadastro}'.")

    except Exception as e:
        registrar_log(f"ERRO: {str(e)}")


if __name__ == "__main__":
    sincronizar_dados_cadastro()
    print("Sincronização concluída com sucesso!")