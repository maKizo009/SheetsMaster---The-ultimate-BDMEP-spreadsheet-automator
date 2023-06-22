import pandas as pd

# Criar DataFrame
df = pd.DataFrame({'NOME': ['ABC', 'ABC', 'DEF', 'DEF', 'DEF', 'DEF', 'ABC'],
                   'DATA': ['05/01/2023', '07/01/2023', '01/02/2023', '09/02/2023', '25/01/2023', '03/01/2023', '02/02/2023']})

# Converter coluna de data para tipo datetime
df['DATA'] = pd.to_datetime(df['DATA'], format='%d/%m/%Y')

# Criar coluna de mês
df['MÊS'] = df['DATA'].dt.month

# Agrupar o DataFrame pelas colunas NOME e MÊS
grouped = df.groupby(['NOME', 'MÊS'])

# Selecionar a data mais recente em cada grupo
result = grouped['DATA'].max().reset_index()

# Exibir resultado
print(result)