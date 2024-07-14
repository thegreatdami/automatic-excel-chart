from openpyxl import load_workbook

# 1- Lê pasta de trabalho e planilha
wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatório"]

# 2- Acessando valor especifico
print(sheet["A3"].value)
print(sheet["B3"].value)

# 3- Iterando valores por meio de loop
for i in range(2, 6):
    ano = sheet["A%s" %i].value
    am = sheet["B%s" %i].value
    bt = sheet["C%s" %i].value
    print("{0} o Aston martin vendeu {1} e o Bentley vendeu {2}".format(ano, am, bt))
