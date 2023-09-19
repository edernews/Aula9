# Aula9


## importando Pandas

import pandas as pd

## declarar as variaveis para acessar o arquivo

path = 'C:\\Users\\FIC\\Documents\\Eder do Nascimento Bernardo\\arquivos\\003 - Arquivos\\'

file = 'dados-vendas.xlsx'

dados = pd.read_excel(path + file, sheet_name='Vendas')

# Seleciona apenas a coluna Vendedor e remove as duplicadas

dados.head()

# Seleciona apenas a coluna Vendedor e remove as duplicadas
df_Estado = dados['Estado'].drop_duplicates()

# Converter o Pandas SERIES em uma lista
lista_Estado = df_Estado.tolist()

lista_Estado

## estrutura para gerar os arquivos de forma automatica


for Estados in lista_Estado:

    df_final_Estado = dados[dados['Estado'] == Estados]

    #####  gerar arquivo 

    path_new = path + 'relatorios-vendedores\\'

    df_final_Estado.to_excel(path_new + Estados +'.xlsx',index = False) ### index tira a coluna inicial 
