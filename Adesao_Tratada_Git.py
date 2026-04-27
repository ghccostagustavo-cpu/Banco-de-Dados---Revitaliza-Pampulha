import pandas as pd
import sqlite3
import unicodedata

cnc = sqlite3.connect(r'caminho_banco_dados.db')
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

df_adesao['MATRÍCULA'] = df_adesao['MATRÍCULA'].apply(matricula_vazia_novA)
df_adesao['COMPLEMENTO:'] = df_adesao['COMPLEMENTO:'].apply(compl)
df_adesao

df_adesao['chave'] = (
    df_adesao['LOGRADORO:'].apply(limpar)+ "|" + 
    df_adesao['NUMERO:'].apply(limpar) + "|" + 
    df_adesao['COMPLEMENTO:'].apply(compl) + "|" + 
    df_adesao['BAIRRO'].apply(limpar)
)

resultado = pd.DataFrame(df_adesao)

resultado = resultado.drop_duplicates()

resultado.to_excel("Adesão Tratada.xlsx", index=False)
resultado.to_sql('Adesão Tratada', cnc, if_exists='replace', index=False)
print("Tratamento concluída[o]. Resultado salvo em Excel e banco de dados.")
print(resultado.head())
