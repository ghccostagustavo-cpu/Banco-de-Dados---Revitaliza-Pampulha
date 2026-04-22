import pandas as pd
import sqlite3
import unicodedata

cnc = sqlite3.connect(r'c:\Users\Revitaliza03\OneDrive - Enops Engenharia S.A\Área de Trabalho\BANCO DE DADOS\BD\BancoDeDados.db')
df_caca = pd.read_sql_query('SELECT * FROM Copia_Caca_Esgoto_Serviços', cnc)
df_adesao = pd.read_sql_query('SELECT * FROM suadesao', cnc)

def compl(texto):
    if pd.isna(texto) or str(texto).strip() == "" or str(texto).upper() == "NAN": 
        return "CASA"
    c = str(texto).upper().strip()
    if c == "CA":
        return "CASA"
    c = c.replace("CA ", "CASA ")
    return "".join(filter(str.isalnum, c))


def limpar(texto):
    if pd.isna(texto) or str(texto).strip() == "" or str(texto).upper() == "NAN": 
        return ""
        
    t = str(texto).upper().replace("R ", "RUA")
    

    t = unicodedata.normalize('NFKD', t).encode('ASCII', 'ignore').decode('utf-8')
    
    return "".join(filter(str.isalnum, t))

def matricula_vazia_novA(texto):
    if not texto: return ""
    m = str(texto).upper().replace('NOVA', 'NOVO')
    return "".join(filter(str.isalnum, m))

df_caca['MATRÍCULA'] = df_caca['MATRÍCULA'].apply(matricula_vazia_novA)
df_adesao['MATRÍCULA'] = df_adesao['MATRÍCULA'].apply(matricula_vazia_novA)
df_caca['Complemento'] = df_caca['Complemento'].apply(compl)
df_adesao['COMPLEMENTO:'] = df_adesao['COMPLEMENTO:'].apply(compl)


df_caca['chave'] = (
    df_caca['Logradouro '].apply(limpar) + "|" +
    df_caca['Num'].apply(limpar) + "|" +
    df_caca['Complemento'].apply(compl)+ "|" +
    df_caca['Bairro'].apply(limpar)
)

df_adesao['chave'] = (
    df_adesao['LOGRADORO:'].apply(limpar)+ "|" + 
    df_adesao['NUMERO:'].apply(limpar) + "|" + 
    df_adesao['COMPLEMENTO:'].apply(compl) + "|" + 
    df_adesao['BAIRRO'].apply(limpar)
)

resultado = pd.merge(
    df_caca, 
    df_adesao[['chave', 'MATRÍCULA'] ], on='chave', how='left', indicator=True)
    
resultado['MATRÍCULA_y'] = resultado['MATRÍCULA_y'].fillna('NOVO')
resultado['MATRÍCULA_x'] = resultado['MATRÍCULA_x'].replace("NOVA", 'NOVO')

resultado = resultado.drop_duplicates()

resultado.to_excel("Resultado_Comparacao_Matricula.xlsx", index=False)
resultado.to_sql('Resultado_Comparacao_Matricula', cnc, if_exists='replace', index=False)
print("Comparação concluída. Resultado salvo em Excel e banco de dados.")
print(df_caca.head(10))