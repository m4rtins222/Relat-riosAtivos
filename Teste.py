import PyPDF2
import csv

# -----------------------------------------------------------------------------------------------------------------------
# Criação de variáveis, listas e dicionários.

elemDespesa = []
totalPNatureza = []
classOcamentario = []

dicionario3Positivo = {}
dicionario3Negativo = {}
dicionario2Positivo = {}
dicionario2Negativo = {}

valoresSubItens = []
nomeSubItens = []
classSubItens = []

aux = False
aux2 = False
# -----------------------------------------------------------------------------------------------------------------------
# Abertura e leitura do PDF.

# Inserir no primeiro argumento da função "Open" o caminho do arquivo em PDF.
pdf_file = open('C:/Users/caiomartins/Desktop/Arquivos_Projeto_Python/Relatórios de Ativo - OUTUBRO.pdf', 'rb')

read_pdf = PyPDF2.PdfFileReader(pdf_file)

number_of_pages = read_pdf.getNumPages()

# -----------------------------------------------------------------------------------------------------------------------
# Tratamento e alocação dos dados.

for n in range(number_of_pages):
    page = read_pdf.getPage(n)

    page_content = page.extractText()

    listaLinhas = page_content.split('\n')

    # Alimentação das Listas.

    for linha in listaLinhas:
        if 'Fundo Previdência: SEG. SOC. DF. FUNDO CAPITALIZADO' in linha:
            aux = True
        if aux == True:
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

            if 'Total por Natureza' in linha:
                aux2 = False
                totalPNatureza.append(linha[20:].replace('.', '').replace(',', '.'))

            if 'Total Geral' in linha:
                aux = False
                break

# Alimentação dos Dicionários.

for n in range(len(elemDespesa)):
    if classOcamentario[n][0] == '3' and '-' not in totalPNatureza[n]:
        dicionario3Positivo[elemDespesa[n]] = [totalPNatureza[n], classOcamentario[n]]
    if classOcamentario[n][0] == '3' and '-' in totalPNatureza[n]:
        dicionario3Negativo[elemDespesa[n]] = [totalPNatureza[n], classOcamentario[n]]

    if classOcamentario[n][0] == '2' and '-' in totalPNatureza[n]:
        dicionario2Positivo[elemDespesa[n]] = [totalPNatureza[n][1:], classOcamentario[n]]
    if classOcamentario[n][0] == '2' and '-' not in totalPNatureza[n]:
        dicionario2Negativo[elemDespesa[n]] = [f'-{totalPNatureza[n]}', classOcamentario[n]]
# -----------------------------------------------------------------------------------------------------------------------
# Escrevendo o CSV.

# Inserir o nome do arquivo com o mês desejado no primeiro argumento da função "Open" (como string), logo após o operador "as" e no primeiro argumento da função "csv.writer".
with open('PLANILHA_FOLHA_NORMAL_CAPITALIZADO_OUTUBRO_2022.csv',
          mode='w') as PLANILHA_FOLHA_NORMAL_CAPITALIZADO_OUTUBRO_2022:
    teste_writer = csv.writer(PLANILHA_FOLHA_NORMAL_CAPITALIZADO_OUTUBRO_2022, delimiter=';', quotechar='|',
                              quoting=csv.QUOTE_MINIMAL)

    teste_writer.writerow(['', '', 'PROVENTOS', '', ''])
    teste_writer.writerow(['ELEM.DESPESAS', 'CLASS.ORC', 'PROVENTOS', 'DESCONTOS', 'SALDO'])

    for elemento in dicionario3Positivo.keys():
        if elemento in dicionario3Negativo.keys():
            teste_writer.writerow([elemento, dicionario3Positivo[elemento][1], dicionario3Positivo[elemento][0],
                                   dicionario3Negativo[elemento][0], str(round(
                    (float(dicionario3Positivo[elemento][0]) + float(dicionario3Negativo[elemento][0])), 2))])
        else:
            teste_writer.writerow([elemento, dicionario3Positivo[elemento][1], dicionario3Positivo[elemento][0], '',
                                   dicionario3Positivo[elemento][0]])

    teste_writer.writerow(['', '', 'DESCONTOS', '', ''])
    teste_writer.writerow(['ELEM.DESPESAS', 'CLASS.ORC', 'PROVENTOS', 'DESCONTOS', 'SALDO'])

    for elemento2 in dicionario2Positivo.keys():
        if elemento2 in dicionario2Negativo.keys():
            teste_writer.writerow([elemento2, dicionario2Positivo[elemento2][1], dicionario2Positivo[elemento2][0],
                                   dicionario2Negativo[elemento2][0], str(round(
                    (float(dicionario2Positivo[elemento2][0]) + float(dicionario2Negativo[elemento2][0])), 2))])
        else:
            teste_writer.writerow([elemento2, dicionario2Positivo[elemento2][1], dicionario2Positivo[elemento2][0], '',
                                   dicionario2Positivo[elemento2][0]])

    teste_writer.writerow(['', '', 'SUBELEMENTOS', '', ''])
    teste_writer.writerow(['SUB.ELEM.DESPESAS', 'CLASS.ORC', 'VALOR'])
    for indice in range(len(nomeSubItens)):
        teste_writer.writerow([nomeSubItens[indice], classSubItens[indice], valoresSubItens[indice]])