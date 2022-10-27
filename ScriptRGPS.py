import PyPDF2
import csv

# Criação de variáveis, listas e dicionários.

classPassivo01 = ['3.31.90.11.01','3.31.90.11.02','3.31.90.11.04','3.31.90.11.06','3.31.90.11.07','3.31.90.11.10','3.31.90.11.34','3.31.90.11.41','3.31.90.11.56','3.31.90.11.57','3.31.90.11.66','3.31.90.11.67','3.31.90.11.95']
classPassivo02 = ['3.31.90.11.22']
classPassivo03 = ['3.31.90.11.31','3.31.90.11.32']
classPassivo04 = ['3.31.90.11.25']

dicionarioClassContabil = {'INSS DE SERVIDORES CELETISTAS': '218830102','PENSAO ALIMENTÍCIA' : '218810110','IRRF SERVIDORES' : '218820104','OUTRAS CONSIGNAÇÕES' : '218810199','VALORES RECEBIDOS P/ OUTROS ÓRGÃOS DO DO GDF' : '218810199'}

elemDespesa = []
totalPNatureza = []
classOcamentario = []
dicionarioclassPassivo = {}

dicionario3Positivo = {}
dicionario3Negativo = {}
dicionario2Positivo = {}
dicionario2Negativo = {}

valoresSubItens = []
nomeSubItens = []
classSubItens = []

totalAbsolutoProventos = 0
totalAbsolutoDesconto = 0
totalDescontosReceita = 0
totalDescontosDespesa = 0

aux = False
aux2 = False
# -----------------------------------------------------------------------------------------------------------------------
# Abertura e leitura do PDF.

#Inserir no primeiro argumento da função "Open" o caminho do arquivo em PDF.
pdf_file = open('C:/Users/caiomartins/Desktop/Arquivos_Projeto_Python/Relatórios de Ativo - OUTUBRO.pdf', 'rb')

read_pdf = PyPDF2.PdfFileReader(pdf_file)

number_of_pages = read_pdf.getNumPages()

# -----------------------------------------------------------------------------------------------------------------------
# Tratamento e alocação dos dados.

for n in range(number_of_pages):
    page = read_pdf.getPage(n)

    page_content = page.extractText()

    listaLinhas = page_content.split('\n')

    #Alimentação das Listas.

    if aux == False:
        for linha in listaLinhas:

            if aux2 == True and linha[-3] == ',' and 'Total por Natureza' not in linha:
                listaAux = linha.split()

                valoresSubItens.append(listaAux[-1].replace('.', '').replace(',', '.'))

                listaAux.pop()
                listaAux.pop()
                listaAux.pop()

                nomeSubItens.append(' '.join(listaAux))
                classSubItens.append(classOcamentario[-1])

            if 'Rubrica Nome do Provento/Desconto Valor Qtd' in linha:
                aux2 = True

            if 'Elem. Despesa' in linha:
                elemDespesa.append(linha[29:])

                if linha[15] == '3':
                    classOcamentario.append(linha[17:28])
                if linha[15] == '2':
                    classOcamentario.append(linha[15:28])

                if linha[15:28] in classPassivo01:
                    dicionarioclassPassivo[linha[29:]] = '211110101'
                elif linha[15:28] in classPassivo02:
                    dicionarioclassPassivo[linha[29:]] = '211110102'
                elif linha[15:28] in classPassivo03:
                    dicionarioclassPassivo[linha[29:]] = '211110103'
                elif linha[15:28] in classPassivo04:
                    dicionarioclassPassivo[linha[29:]] = '211110104'
                else:
                    dicionarioclassPassivo[linha[29:]] = ''

            if 'Total por Natureza' in linha:
                aux2 = False
                totalPNatureza.append(linha[20:].replace('.','').replace(',','.'))

            if 'Fundo Previdência: SEG. SOC. DF. LC 232/99 - FUNDO FINANCEIRO' in linha:
                aux = True
                break

#Alimentação dos Dicionários.
for n in range(len(elemDespesa)):
    if classOcamentario[n][0] == '3' and '-' not in totalPNatureza[n]:
        dicionario3Positivo[elemDespesa[n]] = [totalPNatureza[n],classOcamentario[n]]
    if classOcamentario[n][0] == '3' and '-' in totalPNatureza[n]:
        dicionario3Negativo[elemDespesa[n]] = [totalPNatureza[n],classOcamentario[n]]

    if classOcamentario[n][0] == '2' and '-' in totalPNatureza[n]:
        dicionario2Positivo[elemDespesa[n]] = [totalPNatureza[n][1:], classOcamentario[n]]
    if classOcamentario[n][0] == '2' and '-' not in totalPNatureza[n]:
        dicionario2Negativo[elemDespesa[n]] = [f'-{totalPNatureza[n]}', classOcamentario[n]]

# -----------------------------------------------------------------------------------------------------------------------
# Escrevendo o CSV.

#Inserir o nome do arquivo com o mês desejado no primeiro argumento da função "Open" (como string), logo após o operador "as" e no primeiro argumento da função "csv.writer".
with open('PLANILHA_FOLHA_NORMAL_RGPS_OUTUBRO.csv', mode='w') as PLANILHA_FOLHA_NORMAL_RGPS_OUTUBRO:
    teste_writer = csv.writer(PLANILHA_FOLHA_NORMAL_RGPS_OUTUBRO, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)

    teste_writer.writerow(['', '', 'PROVENTOS', '', ''])
    teste_writer.writerow(['ELEM.DESPESAS', 'CLASS.ORC', 'CLASS.PASS', 'PROVENTOS', 'DESCONTOS', 'SALDO'])

    for elemento in dicionario3Positivo.keys():
        if elemento in dicionario3Negativo.keys():
            teste_writer.writerow([elemento,dicionario3Positivo[elemento][1],dicionarioclassPassivo[elemento], dicionario3Positivo[elemento][0].replace('.',','), dicionario3Negativo[elemento][0].replace('.',','), str(round((float(dicionario3Positivo[elemento][0]) + float(dicionario3Negativo[elemento][0])),2)).replace('.',',')])
            totalAbsolutoProventos += float(dicionario3Positivo[elemento][0])
            totalAbsolutoDesconto += float(dicionario3Negativo[elemento][0][1:])

            for indice in range(len(nomeSubItens)):
                if dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' not in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice],'', valoresSubItens[indice].replace('.',',')])
                elif dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '','', valoresSubItens[indice].replace('.',',')])
        else:
            teste_writer.writerow([elemento, dicionario3Positivo[elemento][1],dicionarioclassPassivo[elemento], dicionario3Positivo[elemento][0].replace('.',','),'',dicionario3Positivo[elemento][0].replace('.',',')])
            totalAbsolutoProventos += float(dicionario3Positivo[elemento][0])

            for indice in range(len(nomeSubItens)):
                if dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' not in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', valoresSubItens[indice].replace('.',',')])
                elif dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', '',valoresSubItens[indice].replace('.',',')])

    teste_writer.writerow(['PROVENTOS', 'DESCONTOS', 'SALDO'])
    teste_writer.writerow([str(round(totalAbsolutoProventos,2)).replace('.', ','), str(round(totalAbsolutoDesconto,2)).replace('.', ','), str(round(totalAbsolutoProventos-totalAbsolutoDesconto,2)).replace('.',',')])

    teste_writer.writerow(['', '', 'DESCONTOS', '', ''])
    teste_writer.writerow(['ELEM.DESPESAS', 'CLASS.ORC', 'CLASS.CONT', 'RECEITA', 'DESPESA', 'SALDO'])

    for elemento2 in dicionario2Positivo.keys():
        if elemento2 in dicionario2Negativo.keys():
            teste_writer.writerow([elemento2,dicionario2Positivo[elemento2][1],dicionarioClassContabil[elemento2], dicionario2Positivo[elemento2][0].replace('.',','), dicionario2Negativo[elemento2][0].replace('.',','), str(round((float(dicionario2Positivo[elemento2][0]) + float(dicionario2Negativo[elemento2][0])),2)).replace('.',',')])
            totalDescontosReceita += float(dicionario2Positivo[elemento2][0])
            totalDescontosDespesa += float(dicionario2Negativo[elemento][0][1:])

            for indice in range(len(nomeSubItens)):
                if dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' not in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', valoresSubItens[indice].replace('.',',')])
                elif dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', '',valoresSubItens[indice].replace('.',',')])

        else:
            teste_writer.writerow([elemento2, dicionario2Positivo[elemento2][1],dicionarioClassContabil[elemento2], dicionario2Positivo[elemento2][0].replace('.',','),'',dicionario2Positivo[elemento2][0].replace('.',',')])
            totalDescontosReceita += float(dicionario2Positivo[elemento2][0])

            for indice in range(len(nomeSubItens)):
                if dicionario2Positivo[elemento2][1] == classSubItens[indice] and '-' not in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', valoresSubItens[indice].replace('.',',')])
                elif dicionario2Positivo[elemento2][1] == classSubItens[indice] and '-' in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', '',valoresSubItens[indice].replace('.',',')])

    teste_writer.writerow(['','RECEITA','DESPESAS', 'SALDO'])
    teste_writer.writerow(['SUBTOTAIS',str(round(totalDescontosReceita,2)).replace('.',','),str(round(totalDescontosDespesa,2)).replace('.',','),str(round(totalDescontosReceita-totalDescontosDespesa,2)).replace('.',',')])
    teste_writer.writerow(['TOTAL DE DESCONTOS', str(round(totalDescontosReceita+totalAbsolutoDesconto,2)).replace('.',',')])
    teste_writer.writerow(['TOTAL DA FOLHA (Líquido)', str(round(totalAbsolutoProventos-(totalDescontosReceita + totalAbsolutoDesconto),2)).replace('.',',')])