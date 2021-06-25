import os
import pandas as pd
import matplotlib.pyplot as plt
plt.style.use('seaborn')

df = pd.read_excel('AdventureWorks.xlsx')


#Qual a receita total?
ReceitaTotal = df['Valor Venda'].sum()

#Qual o custo total?
df['Custo'] = df['Custo Unitário'].mul(df['Quantidade']) #Criando a coluna de custo
CustoTotal = round(df['Custo'].sum(),2)

#Agora, com receita e custo, calcula-se lucro:
df['Lucro'] = df['Valor Venda'] - df['Custo']

#Lucro total:
LucroTotal = round(df['Lucro'].sum(),2)

#Total de dias para enviar o produto:
#queremos saber a média do tempo de envio para cada marca e para isso é necessário
#transformar a coluna Tempo Envio em um valor numério:
df['Tempo Envio'] = (df['Data Envio'] - df['Data Venda']).dt.days
MediaTempoPorMarca = df.groupby('Marca')['Tempo Envio'].mean()

#Analisar a existência de valores ausentes:
Nnan = df.isnull().sum()

#Lucro por ano:
LucroPorAno = df.groupby(df['Data Venda'].dt.year)['Lucro'].sum()

#Agrupar por ano e por marca:
LucroPorAnoPorMarca = df.groupby([df['Data Venda'].dt.year, 'Marca'])['Lucro'].sum().reset_index()

#Total de produtos vendidos:
TotalProdVendidos = df.groupby('Produto')['Quantidade'].sum().sort_values(ascending=True)

#Vendas do ano de 2009:
df2009 = df[df['Data Venda'].dt.year == 2009]
Lucro2009PorMes = df2009.groupby(df2009['Data Venda'].dt.month)['Lucro'].sum()
Lucro2009PorMarca = df2009.groupby('Marca')['Lucro'].sum()
Lucro2009PorClasse = df2009.groupby('Classe')['Lucro'].sum()


#Gráficos:
plt.figure(1)
TotalProdVendidos.plot.barh(title='Total de produtos vendidos')
plt.xlabel('Total')
plt.ylabel('Produto')

plt.figure(2)
LucroPorAno.plot.bar(title='Lucro por ano')
plt.xlabel('Ano')
plt.ylabel('Receita')

plt.figure(3)
Lucro2009PorMes.plot(title='Lucro por mês')
plt.xlabel('Mês')
plt.ylabel('Lucro')

plt.figure(4)
Lucro2009PorMarca.plot.bar(title='Lucro por marca em 2009')
plt.xlabel('Marca')
plt.ylabel('Lucro')
plt.xticks(rotation='horizontal')

plt.figure(5)
Lucro2009PorClasse.plot.bar(title='Lucro por classe em 2009')
plt.xlabel('Classe')
plt.ylabel('Lucro')
plt.xticks(rotation='horizontal')


#Análises estatísticas:
Descricao = df['Tempo Envio'].describe()

plt.figure(6)
plt.boxplot(df['Tempo Envio'])
#Identificar outlier:
outlier = df['Tempo Envio'].max()
print('Envio com tempo anormal:')
print(df[df['Tempo Envio'] == outlier])

plt.show()

#Por fim, salvar como CSV
df.to_csv('df_vendas_novo.csv', index=False)

