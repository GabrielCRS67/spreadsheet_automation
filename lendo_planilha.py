import openpyxl

# Carregando a planilha
book = openpyxl.load_workbook('Importar2.xlsx')

# Lendo a planilha
read_page = book['Import']

read_page.append(['Deu certo'])

book.save('Importar2.xlsx')
