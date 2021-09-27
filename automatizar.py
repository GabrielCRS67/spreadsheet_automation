from openpyxl import Workbook

def CriaLinhaTitulo(page):
    page.append(['ID', 'Marca', 'Produto', 'Perfil', 'Sistema', 'NomeDis', 'DataCriacao', 'Tipo', 'DataGarantia', 'Destino', 'Pedido', 'Status', 'Ícone'])

def CriaLinhas(page, quantidade_de_linhas):
    contador = 0
    identificador = 38471
    while (contador < quantidade_de_linhas):
        identificador += 1
        page.append([identificador,marcacoes, cilindro.cilindro, 235, sistema.gts,'Claro-Sp', '25/09/2021', 'WEB CLIQ MANAGER','25/09/2022', destino.Enel, 3145, 3,3])
        contador += 1
        
# Sistemas da planilha  
class Sistemas:
    def __init__(self):
        self.vivo = 'Vivo Teste'
        self.claro = 'Claro Teste'
        self.nextel = 'Nextel Teste'
        self.enel = 'Enel Teste'
        self.gts = 'GTS Teste'
    pass

sistema = Sistemas()        

# Produtos da planilha
class Produtos:
    def __init__(self, cadeado, chave, cilindro):
        self.cadeados = cadeado
        self.chave = chave
        self.cilindro = cilindro
    pass

cadeadoCliq = Produtos('Cadeado CLIQ G55','','')
chaveTK = Produtos('', 'Chave Temporary Key', '')
cilindro = Produtos('','','Cilindro CLIQ')


# Destino da planilha
class Destinos:
    def __init__(self):
        self.Ericson = 'Ericsson'
        self.Enel = 'Enel'
        self.ClaroSp = 'Claro-SP'
        self.VivoBahia = 'Vivo-Bahia'
    pass

destino = Destinos()

#----------------------------------------------------
marcacoes = str(['A70Z2',
'A70Z8',
'A407V',
'A510V',
'A70VZ',
'A70VV',
'A70W0',
'A70VR',
'A70VT',
'A70VX',
'A70W1',
'A70W2'
])

# Criada a planilha
book = Workbook()
book.create_sheet('Import', 0)

# Selecionando a planilha
page_import = book['Import']
CriaLinhaTitulo(page_import)
CriaLinhas(page_import, 90)


#Salvando a planilha
book.save('Importar.xlsx')

print(book.sheetnames)