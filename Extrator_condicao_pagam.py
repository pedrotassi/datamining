# Extrair palavras-chave de multiplos arquivos PDFs numa pasta, cria um dataframe, depois exporta num arquivo xlsx.

import os                        # nos permite trabalhar com pastas do windows
import pandas as pd              # ferramenta de analise de dados
import glob                      # usado no tratamento de diversos arquivos
import pdfplumber                # trabalha informações de pdfs
import datetime as dt

# limpando o console 
os.system('cls')


# função para encontrar palavras entre duas palavras
def get_keyword(start, end, text):
    """
    start: palavra antetior a palavra-chave.
    end: palavra logo após a palavra-chave.
    text: representa o texto do pdf em si.
    """
    for i in range(len(start)):
        try:
            field = ((text.split(start[i]))[1].split(end[i])[0])
            return field
        except:
            continue

# função que transforma a data para o primeiro dia do mês
def modificarData_B_PM(data):
    data = data.split('/')
    data[0] = '01'
    nova_data = data[0] + '/' + data[1] + '/' + data[2]
    return nova_data

def main():
    # criando uma dataframe vazio, onde posteriormente as palavras-chave serão adicionadas em linhas
    my_dataframe = pd.DataFrame()

    # loop para ler varios arquivos de pdf em determinada pasta
    for files in glob.glob(r'pdfs\*.pdf'):
        with pdfplumber.open(files) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            text = " ".join(text.split())


            #colunas vazias
            empresa = ''
            empreendimento = ''
            

            #obtendo as palavras-chave
            start = ['Proposta']
            end = [' ']
            keyword1 = get_keyword(start, end, text)

            #condicional que resolve um erro de espaços duplos
            if keyword1 == '':
                start = ['Proposta ']
                end = [' ']
                keyword1 = get_keyword(start, end, text)
            
            #certas palavras necessitaram de outro método de obtenção
            if text.find('Sinal'):
                ato = 'AT'
                pos = text.find('Sinal')
                valor_ato = text[pos + 34:]
                valor_ato = valor_ato.split(' ')
                valor_ato = valor_ato[0]

                # funçao para transformar o valor monetario que esta com "." e "," em decimal/float para que possamos realizar operações aritméticas
                def make_decimal(string):
                    result = 0
                    if string:
                        # cria duas variaveis e remove a virgula separando o valor que estava antes da virgula em uma variavel e o que estava depois em outra
                        [num, dec] = string.rsplit(',')
                        # remove o ponto e o substitui por espaço em vazio
                        result += int(num.replace('.', ''))
                        # adiciona 
                        result += (int(dec) / 100)
                    return result
                valor_ato = make_decimal(valor_ato)

            # obtendo as palavras chaves
            start = ['Parcela 1 / ']
            end = [' ']
            keyword2 = get_keyword(start, end, text)
            
                
            condicao_pm = ''

            if keyword2 != '1':
                n_parcelas = keyword2
                condicao_pm = 'PM'

            
            start = ['Entrada / Parcela']
            end = ['/']
            keyword3 = get_keyword(start, end, text) 
            keyword3 = int(keyword3)

            if keyword3 != 0:
                qtd_parcelas_ato = keyword3
                qtd_parcelas_ato = int(qtd_parcelas_ato)
            else:
                qtd_parcelas_ato = 1

            # usando outro método para encontrar as palavras pois em alguns casos o outro pode não funcionar corretamente
            pos = text.find('Entrada / Parcela')
            qtd_parcelas = text[pos + 22:]
            qtd_parcelas = qtd_parcelas.split(' ')
            qtd_parcelas = qtd_parcelas[0]
            qtd_parcelas = int(qtd_parcelas)

            pos = text.find('Sinal')
            venc_ato = text[pos + 12:]
            venc_ato = venc_ato.split(' ')
            venc_ato = venc_ato[0]

            pos = text.find('Prestação 1 / ')
            venc_parc = text[pos + 17:]
            venc_parc = venc_parc.split(' ')
            venc_parc = venc_parc[1]

            start = ['Preço']
            end = ['Entrada']
            keyword6 = get_keyword(start, end, text)
            
            #Função para converter moeda que está registrada como string num float 
            def make_decimal(string):
                    result = 0
                    if string:
                        #divide o valor e centavos em duas variaveis
                        [num, dec] = string.rsplit(',')
                        #remove os pontos
                        result += int(num.replace('.', ''))
                        #encontra o decimal
                        result += (int(dec) / 100)
                    return result
            keyword6 = make_decimal(keyword6)

            pg_rest = keyword6 - valor_ato

            indexador_ato = 0
            indexador_pm = 1

            data_base_ato = venc_ato

            date_dt2 = dt.datetime.strptime(venc_parc, '%d/%m/%Y')
            glanceback = dt.timedelta(days=60)
            data_base_pm = date_dt2 - glanceback

            # formatar o datetime para o formato desejado (dia/mês/ano)
            data_pm_1 = dt.datetime.strftime(data_base_pm, "%d/%m/%Y")  
            
            data_base_pm = modificarData_B_PM(data_pm_1)

            # cria uma lista com as palavras extraidas do documento atual
            my_list = [keyword1, ato, valor_ato, qtd_parcelas_ato, venc_ato, indexador_ato, data_base_ato]
            my_list2 = [keyword1, condicao_pm, pg_rest, qtd_parcelas, venc_parc, indexador_pm, data_base_pm]

            # transforma a  lista numa linha de dataframe
            my_list = pd.Series(my_list)
            my_list2 = pd.Series(my_list2)

            # adiciona a lista no meu dataframe.
            if valor_ato != 0:
                my_dataframe = my_dataframe.append(my_list, ignore_index=True)
            else:
                pass
            my_dataframe = my_dataframe.append(my_list2, ignore_index=True)

            print("Condições extraidas com sucesso")
            

    # renomeando as colunas do dataframe
    my_dataframe = my_dataframe.rename(columns={0:'numero contrato', 1:'tipo condicao', 2:'valor total', 3:'qtidade parcelas', 4:'primeiro vencimento', 5:'indexador', 6:'data base'})
                                                    

    #indicando onde o arquivo excel será exportado
    save_path = (r'') # coloque aqui entre as aspas o caminho em que deseja salvar o arquivo
    os.chdir(save_path)
    #exporta o dataframe em .xlsx
    my_dataframe.to_excel('condicao_pagamento.xlsx', sheet_name = 'my dataframe', index=False)
    print("")
    print(my_dataframe)


if __name__ == '__main__':
    main()