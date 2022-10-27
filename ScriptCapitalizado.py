import PyPDF2
import csv

# -----------------------------------------------------------------------------------------------------------------------
# Criação de variáveis, listas e dicionários.

classPassivo01 = ['3.31.90.11.01','3.31.90.11.02','3.31.90.11.04','3.31.90.11.06','3.31.90.11.07','3.31.90.11.10','3.31.90.11.34','3.31.90.11.41','3.31.90.11.56','3.31.90.11.57','3.31.90.11.66','3.31.90.11.67','3.31.90.11.95']
classPassivo02 = ['3.31.90.11.22']
classPassivo03 = ['3.31.90.11.31','3.31.90.11.32']
classPassivo04 = ['3.31.90.11.25']

dicionarioClassContabil = {'INSS DE SERVIDORES ESTATUTÁRIOS': '421110201','PENSAO ALIMENTÍCIA' : '218810110','IRRF SERVIDORES' : '218820104','OUTRAS CONSIGNAÇÕES' : '218810199','VALORES RECEBIDOS P/ OUTROS ÓRGÃOS DO DO GDF' : '218810199'}

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
totalReceitasProventos = 0
totalReceitasDescontos = 0

prevcomNormal = 0
prevcom13salario = 0
prevcomDiffOutrosMeses = 0

aux = False
aux2 = False
aux3 = False
aux4 = False
aux5 = False
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
        if 'PREVCOM NORMAL- 42340' in linha:
            aux3 = True
        if 'DF-PREVICOM - CONTRIB. FACULTATIVA- 42344' in linha:
            aux3 = True
        if 'Total de Servidores' in linha and aux3 == True:
            listaaux3 = linha.split()
            prevcomNormal += float(listaaux3[2].replace('.','').replace(',','.'))
            aux3 = False

        if 'DIF. PREVCOM NORMAL - 13º SALÁRIO- 5234' in linha:
            aux4 = True
        if 'PREVCOM NORMAL - 13º SALÁRIO- 42343' in linha:
            aux4 = True
        if 'Total de Servidores' in linha and aux4 == True:
            listaaux4 = linha.split()
            prevcom13salario += float(listaaux4[2].replace('.','').replace(',','.'))
            print(prevcom13salario)
            aux4 = False

        if 'DIF. PREVCOM NORMAL- 52340' in linha and 'Total de Servidores por Ref' not in linha:
            aux5 = True

        if 'Total de Servidores' in linha and aux5 == True:
            listaaux5 = linha.split()
            prevcomDiffOutrosMeses += float(listaaux5[2].replace('.','').replace(',','.'))
            aux5 = False


        if 'SEGURIDADE SOCIAL - FUNDO CAPITALIZADO' in linha and 'INSTITUTO:' not in linha:
            listaux2 = linha.split()

        if 'Fundo Previdência: SEG. SOC. DF. FUNDO CAPITALIZADO' in linha:
            aux = True

        if aux == True:
            if aux2 == True and linha[-3] == ',' and 'Total por Natureza' not in linha and aux == True:
                listaAux = linha.split()

                valoresSubItens.append(listaAux[-1].replace('.', '').replace(',', '.'))

                listaAux.pop()
                listaAux.pop()
                listaAux.pop()

                nomeSubItens.append(' '.join(listaAux))
                classSubItens.append(classOcamentario[-1])

            if 'Rubrica Nome do Provento/Desconto Valor Qtd' in linha:
                aux2 = True

            if 'Elem. Despesa' in linha and aux == True:
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

            if 'Total por Natureza' in linha and aux == True:
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
with open('PLANILHA_FOLHA_NORMAL_CAPITALIZADO_OUTUBRO_2022.csv', mode='w') as PLANILHA_FOLHA_NORMAL_CAPITALIZADO_OUTUBRO_2022:
    teste_writer = csv.writer(PLANILHA_FOLHA_NORMAL_CAPITALIZADO_OUTUBRO_2022, delimiter=';', quotechar='|',quoting=csv.QUOTE_MINIMAL)

    teste_writer.writerow(['', '', 'PROVENTOS', '', ''])
    teste_writer.writerow(['ELEM.DESPESAS', 'CLASS.ORC', 'CLASS.PASS', 'PROVENTOS', 'DESCONTOS', 'SALDO'])

    for elemento in dicionario3Positivo.keys():
        if elemento in dicionario3Negativo.keys():
            teste_writer.writerow([elemento, dicionario3Positivo[elemento][1], dicionarioclassPassivo[elemento],dicionario3Positivo[elemento][0].replace('.', ','),dicionario3Negativo[elemento][0].replace('.', ','), str(round((float(dicionario3Positivo[elemento][0]) + float(dicionario3Negativo[elemento][0])), 2)).replace('.', ',')])
            totalAbsolutoProventos += float(dicionario3Positivo[elemento][0])
            totalAbsolutoDesconto += float(dicionario3Negativo[elemento][0][1:])

            for indice in range(len(nomeSubItens)):
                if dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' not in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', valoresSubItens[indice].replace('.',',')])
                elif dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', '',valoresSubItens[indice].replace('.',',')])

        else:
            teste_writer.writerow([elemento, dicionario3Positivo[elemento][1], dicionarioclassPassivo[elemento],dicionario3Positivo[elemento][0].replace('.', ','), '',dicionario3Positivo[elemento][0].replace('.', ',')])
            totalAbsolutoProventos += float(dicionario3Positivo[elemento][0])

            for indice in range(len(nomeSubItens)):
                if dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' not in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', valoresSubItens[indice].replace('.',',')])
                elif dicionario3Positivo[elemento][1] == classSubItens[indice] and '-' in valoresSubItens[indice]:
                    teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', '',valoresSubItens[indice].replace('.',',')])

    teste_writer.writerow(['PROVENTOS', 'DESCONTOS', 'SALDO'])
    teste_writer.writerow([str(round(totalAbsolutoProventos,2)).replace('.',','), str(round(totalAbsolutoDesconto,2)).replace('.',','), str(round(totalAbsolutoProventos-totalAbsolutoDesconto,2)).replace('.',',')])

    teste_writer.writerow(['', '', 'DESCONTOS', '', ''])
    teste_writer.writerow(['ELEM.DESPESAS', 'CLASS.ORC', 'CLASS.CONT', 'PROVENTOS', 'DESCONTOS', 'SALDO'])
# ------------------------------------------------------------------------------------------------------------------------------------------
    for elemento2 in dicionario2Positivo.keys():
        if elemento2 in dicionario2Negativo.keys():
            teste_writer.writerow([elemento2, dicionario2Positivo[elemento2][1], dicionarioClassContabil[elemento2],dicionario2Positivo[elemento2][0].replace('.', ','),dicionario2Negativo[elemento2][0].replace('.', ','), str(round((float(dicionario2Positivo[elemento2][0]) + float(dicionario2Negativo[elemento2][0])), 2)).replace('.', ',')])
            totalReceitasProventos += float(dicionario2Positivo[elemento2][0])
            totalReceitasDescontos += float(dicionario2Negativo[elemento2][0][1:])

            for indice in range(len(nomeSubItens)):
                if dicionario2Positivo[elemento2][1] == classSubItens[indice] and '-' not in valoresSubItens[indice]:
                    if 'PREVICOM' in nomeSubItens[indice] or 'PREVCOM' in nomeSubItens[indice]:
                        teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '218810111',valoresSubItens[indice].replace('.', ',')])
                    else:
                        teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '',valoresSubItens[indice].replace('.', ',')])

                elif dicionario2Positivo[elemento2][1] == classSubItens[indice] and '-' in valoresSubItens[indice]:
                    if 'PREVICOM' in nomeSubItens[indice] or 'PREVCOM' in nomeSubItens[indice]:
                        teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '218810111', '',valoresSubItens[indice].replace('.', ',')])
                    else:
                        teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', '',valoresSubItens[indice].replace('.', ',')])
        else:
            teste_writer.writerow([elemento2, dicionario2Positivo[elemento2][1], dicionarioClassContabil[elemento2],dicionario2Positivo[elemento2][0].replace('.', ','), '',dicionario2Positivo[elemento2][0].replace('.', ',')])
            totalReceitasProventos += float(dicionario2Positivo[elemento2][0])

            for indice in range(len(nomeSubItens)):
                if dicionario2Positivo[elemento2][1] == classSubItens[indice] and '-' not in valoresSubItens[indice]:
                    if 'PREVICOM' in nomeSubItens[indice] or 'PREVCOM' in nomeSubItens[indice]:
                        teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '218810111',valoresSubItens[indice].replace('.', ',')])
                    else:
                        teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '',valoresSubItens[indice].replace('.', ',')])

                elif dicionario2Positivo[elemento2][1] == classSubItens[indice] and '-' in valoresSubItens[indice]:
                    if 'PREVICOM' in nomeSubItens[indice] or 'PREVCOM' in nomeSubItens[indice]:
                        teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '218810111', '',valoresSubItens[indice].replace('.', ',')])
                    else:
                        teste_writer.writerow([f'       {nomeSubItens[indice].lower()}', classSubItens[indice], '', '',valoresSubItens[indice].replace('.', ',')])

    totalpFundo = (totalAbsolutoProventos + totalReceitasDescontos) - (totalAbsolutoDesconto + totalReceitasProventos)

    teste_writer.writerow(['', 'PROVENTOS', 'DESCONTOS', 'SALDO'])
    teste_writer.writerow(['SUBTOTAIS', str(round(totalReceitasProventos,2)).replace('.', ','), str(round(totalReceitasDescontos,2)).replace('.', ','),str(round(totalReceitasProventos - totalReceitasDescontos,2)).replace('.', ',')])
    teste_writer.writerow(['TOTAL DE PROVENTOS', str(round(totalAbsolutoProventos + totalReceitasDescontos,2)).replace('.', ',')])
    teste_writer.writerow(['TOTAL DE DESCONTOS',str(round(totalAbsolutoDesconto + totalReceitasProventos,2)).replace('.',',')])
    teste_writer.writerow(['TOTAL POR FUNDO DE PREVIDÊNCIA',str(round(totalpFundo,2)).replace('.',',')])

    teste_writer.writerow(['', 'VALOR EMPREGADOR','VALOR EMPREGADOR 13º', 'VALOR DEV.EMPREGADOR', 'VALOR DEV.EMPREGADOR 13º','TOTAL'])
    teste_writer.writerow(['PATRONAL FUNDO CAPITALIZADO',listaux2[1],listaux2[2],listaux2[9],listaux2[10],listaux2[3]])

    teste_writer.writerow(['','VLR EMPREGADOR - MÊS DA FOLHA','VLR 13º','DIF.OUTROS MESES','TOTAL'])
    teste_writer.writerow(['PATRONAL PREVICOM',round(prevcomNormal,2),round(prevcom13salario,2),round(prevcomDiffOutrosMeses,2),round(prevcomNormal+prevcomDiffOutrosMeses+prevcom13salario,2)])