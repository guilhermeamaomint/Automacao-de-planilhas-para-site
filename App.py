import openpyxl
woerkbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = woerkbook["Produtos"]

for linha in sheet_produtos.iter_rows(min_row=2):
    print(linha[0].value)


