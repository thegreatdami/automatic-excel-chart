import pandas as pd 

# 1-Importando Dados
data = pd.read_excel("data/VendaCarros (1).xlsx")
print(type(data))

# 2- Selecionando colunas específicas do dataframe
df = data[["Fabricante", "ValorVenda", "Ano"]]
print(df)

# 3-Criando tabela pivô
pivot_table = df.pivot_table(
    index="Ano",
    columns="Fabricante",
    values="ValorVenda",
    aggfunc="sum"
)

print(pivot_table)

#4- Exportando tabela pivô em arquivo excel
pivot_table.to_excel("data/pivot_table.xlsx", "Relatório")