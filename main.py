import tabula
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import jpype
produtos = "VendasPorProdutos.pdf"

print("Lendo o arquivo ``VendasPorProdutos.pdf`` \n")
try:
    ler1 = tabula.read_pdf(produtos, pages="all")
except:
    print("não foi possivel ler o arquivo!!!\n")
else:
    print("Arquivo lido com sucesso!!!\n")
    
tabela = ler1[0]

tabela[['Codigo', 0]] = tabela["Código UN"].str.split(" ", expand=True)
tabela = tabela.drop("Código UN", axis=1)
tabela = tabela.drop(0, axis=1)
tabela = tabela.drop("Barras", axis=1)
tabela = tabela.drop("Custo", axis=1)
tabela = tabela.drop("% Média", axis=1)
tabela = tabela.drop("Lucro bruto", axis=1)
tabela['Codigo'] = tabela['Codigo'].astype(str).str[-6:]
tabela = tabela.set_index("Codigo")

tabela.iloc[:, 1] = tabela.iloc[:, 1].str.replace(',000', '')
tabela.iloc[:, 1] = pd.to_numeric(tabela.iloc[:, 1], errors='coerce')

tabela.iloc[:, 2] = tabela.iloc[:, 2].str.replace(',', '.')
tabela.iloc[:, 2] = pd.to_numeric(tabela.iloc[:, 2], errors='coerce')

print(tabela)

relatorio = openpyxl.load_workbook("CAIXA_format.xlsx", data_only=False)
caixa = relatorio['Planilha1']


moeda_style = openpyxl.styles.NamedStyle(name='moeda', number_format='"R$" #,##0.00;[Red]"R$" -#,##0.00')
moeda_style.alignment = openpyxl.styles.Alignment(horizontal='right', vertical='bottom')
moeda_style.font = openpyxl.styles.Font(name='Calibri', size=14)
cor_fundo = PatternFill(fill_type="solid", fgColor="ffff00")
borda_style = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'),
                                     right=openpyxl.styles.Side(style='thin'),
                                     top=openpyxl.styles.Side(style='thin'),
                                     bottom=openpyxl.styles.Side(style='thin'))

for row in caixa.iter_rows(min_row=4, max_row=41, min_col=3, max_col=3):
    for cell in row:
        cell.style = moeda_style
        cell.border = borda_style

for row in caixa.iter_rows(min_row=4, max_row=10, min_col=7, max_col=7):
    for cell in row:
        cell.style = moeda_style
        cell.border = borda_style

for row in caixa.iter_rows(min_row=47, max_row=50, min_col=2, max_col=2):
    for cell in row:
        cell.style = moeda_style
        cell.border = borda_style
for row in caixa.iter_rows(min_row=55, max_row=58, min_col=2, max_col=2):
    for cell in row:
        cell.style = moeda_style
        cell.border = borda_style

caixa['F11'].style = moeda_style
caixa['F11'].fill = cor_fundo
caixa['F11'].border = borda_style

caixa['G11'].style = moeda_style
caixa['G11'].fill = cor_fundo
caixa['G11'].border = borda_style

caixa['C42'].style = moeda_style
caixa['C42'].fill = cor_fundo
caixa['C42'].border = borda_style

caixa['C44'].style = moeda_style
caixa['C44'].fill = cor_fundo
caixa['C44'].border = borda_style

caixa['B51'].style = moeda_style
caixa['B51'].fill = cor_fundo
caixa['B51'].border = borda_style

caixa['B59'].style = moeda_style
caixa['B59'].fill = cor_fundo
caixa['B59'].border = borda_style

caixa['C61'].style = moeda_style
caixa['C61'].fill = cor_fundo
caixa['C61'].border = borda_style

try:
    caixa['B4'] = tabela['Qtd']['000001']
    caixa['C4'] = tabela['Venda']['000001']
except:
    pass

try:
    caixa['B5'] = tabela['Qtd']['000002']
    caixa['C5'] = tabela['Venda']['000002']
except:
    pass

try:
    caixa['B6'] = tabela['Qtd']['000003']
    caixa['C6'] = tabela['Venda']['000003']
except:
    pass

try:
    caixa['B7'] = tabela['Qtd']['000007']
    caixa['C7'] = tabela['Venda']['000007']
except:
    pass

try:
    caixa['B8'] = tabela['Qtd']['000011']
    caixa['C8'] = tabela['Venda']['000011']
except:
    pass

try:
    caixa['B9'] = tabela['Qtd']['000012']
    caixa['C9'] = tabela['Venda']['000012']
except:
    pass

try:
    caixa['B10'] = tabela['Qtd']['000026']
    caixa['C10'] = tabela['Venda']['000026']
except:
    pass

try:
    caixa['B11'] = tabela['Qtd']['000027']
    caixa['C11'] = tabela['Venda']['000027']
except:
    pass

try:
    caixa['B12'] = tabela['Qtd']['000052']
    caixa['C12'] = tabela['Venda']['000052']
except:
    pass

try:
    caixa['B13'] = tabela['Qtd']['000028']
    caixa['C13'] = tabela['Venda']['000028']
except:
    pass

try:
    caixa['B14'] = tabela['Qtd']['000086']
    caixa['C14'] = tabela['Venda']['000086']
except:
    pass

try:
    caixa['B15'] = tabela['Qtd']['000037']
    caixa['C15'] = tabela['Venda']['000037']
except:
    pass

try:
    caixa['B16'] = tabela['Qtd']['000114']
    caixa['C16'] = tabela['Venda']['000114']
except:
    pass

try:
    caixa['B17'] = tabela['Qtd']['000040']
    caixa['C17'] = tabela['Venda']['000040']
except:
    pass

try:
    caixa['B18'] = tabela['Qtd']['000057']
    caixa['C18'] = tabela['Venda']['000057']
except:
    pass

try:
    caixa['B19'] = tabela['Qtd']['000085']
    caixa['C19'] = tabela['Venda']['000085']
except:
    pass

try:
    caixa['B20'] = tabela['Qtd']['000122']
    caixa['C20'] = tabela['Venda']['000122']
except:
    pass

try:
    caixa['B21'] = tabela['Qtd']['000063']
    caixa['C21'] = tabela['Venda']['000063']
except:
    pass

try:
    caixa['B22'] = tabela['Qtd']['000075']
    caixa['C22'] = tabela['Venda']['000075']
except:
    pass

try:
    caixa['B23'] = tabela['Qtd']['000076']
    caixa['C23'] = tabela['Venda']['000076']
except:
    pass

try:
    caixa['B24'] = tabela['Qtd']['000079']
    caixa['C24'] = tabela['Venda']['000079']
except:
    pass

try:
    caixa['B26'] = tabela['Qtd']['000077']
    caixa['C26'] = tabela['Venda']['000077']
except:
    pass

try:
    caixa['B27'] = tabela['Qtd']['000080']
    caixa['C27'] = tabela['Venda']['000080']
except:
    pass

try:
    caixa['B28'] = tabela['Qtd']['000082']
    caixa['C28'] = tabela['Venda']['000082']
except:
    pass

try:
    caixa['B29'] = tabela['Qtd']['000123']
    caixa['C29'] = tabela['Venda']['000123']
except:
    pass

try:
    caixa['B30'] = tabela['Qtd']['000117']
    caixa['C30'] = tabela['Venda']['000117']
except:
    pass

try:
    caixa['B31'] = tabela['Qtd']['000124']
    caixa['C31'] = tabela['Venda']['000124']
except:
    pass



try:
    caixa['B32'] = tabela['Qtd']['000068']
    caixa['C32'] = tabela['Venda']['000068']
except:
    pass

try:
    caixa['B33'] = tabela['Qtd']['000070']
    caixa['C33'] = tabela['Venda']['000070']
except:
    pass

try:
    caixa['B34'] = tabela['Qtd']['000069']
    caixa['C34'] = tabela['Venda']['000069']
except:
    pass

try:
    caixa['B35'] = tabela['Qtd']['000118']
    caixa['C35'] = tabela['Venda']['000118']
except:
    pass

try:
    caixa['B36'] = tabela['Qtd']['000112']
    caixa['C36'] = tabela['Venda']['000112']
except:
    pass

try:
    caixa['B37'] = tabela['Qtd']['000121']
    caixa['C37'] = tabela['Venda']['000121']
except:
    pass

try:
    caixa['B38'] = tabela['Qtd']['000033']
    caixa['C38'] = tabela['Venda']['000033']
except:
    pass

try:
    caixa['B39'] = tabela['Qtd']['000073']
    caixa['C39'] = tabela['Venda']['000073']
except:
    pass

try:
    caixa['B41'] = tabela['Qtd']['000087']
    caixa['C41'] = tabela['Venda']['000087']
except:
    pass

try:
    caixa['B40'] = tabela['Qtd']['000095']
    caixa['C40'] = tabela['Venda']['000095']
except:
    pass

try:
    caixa['F4'] = tabela['Qtd']['000009']
    caixa['G4'] = tabela['Venda']['000009']
except:
    pass

try:
    caixa['F5'] = tabela['Qtd']['000010']
    caixa['G5'] = tabela['Venda']['000010']
except:
    pass

try:
    caixa['F6'] = tabela['Qtd']['000031']
    caixa['G6'] = tabela['Venda']['000031']
except:
    pass

try:
    caixa['F7'] = tabela['Qtd']['000004']
    caixa['G7'] = tabela['Venda']['000004']
except:
    pass

try:
    caixa['F8'] = tabela['Qtd']['000005']
    caixa['G8'] = tabela['Venda']['000005']
except:
    pass

try:
    caixa['F9'] = tabela['Qtd']['000006']
    caixa['G9'] = tabela['Venda']['000006']
except:
    pass

try:
    caixa['F10'] = tabela['Qtd']['000025']
    caixa['G10'] = tabela['Venda']['000025']
except:
    pass

print("Gerando CAIXA fechado!")
try:
    relatorio.save('CAIXA_fechado.xlsx')
except:
    print("não foi possivel gerar o arquivo!\n")
else:
    print("arquivo gerado com sucesso!!!\n")

input("Pressione ENTER para fechar!!!!!!\n")
