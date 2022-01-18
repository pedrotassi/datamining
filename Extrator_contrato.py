# Extrair palavras-chave de multiplos arquivos PDFs numa pasta, cria um dataframe, depois exporta num arquivo xlsx.

import os                        # nos permite trabalhar com pastas do windows
import pandas as pd              # ferramenta de analise de dados
import glob                      # usado no tratamento de diversos arquivos
import pdfplumber                # trabalha informações de pdfs

# função para modificar o inicio da data para o dia 01
def modificarDatas(data):
    data = data.split('/')

    data[0] = '01'
    nova_data = data[0] + '/' + data[1] + '/' + data[2]

    return nova_data


# função para encontrar palavras entre duas palavras
def get_keyword(start, end, text):
    """
    start: palavra anterior a palavra-chave.
    end: palavra logo após a palavra-chave.
    text: representa o texto do pdf em si.
    """
    for i in range(len(start)):
        try:
            field = ((text.split(start[i]))[1].split(end[i])[0])
            return field
        except:
            continue

def main():
    # criando uma dataframe vazio, onde as palavras-chave serão adicionadas em linhas
    my_dataframe = pd.DataFrame()

    # loop para rodar varios pdfs
    for files in glob.glob(r'pdfs\*.pdf'):
        with pdfplumber.open(files) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            text = " ".join(text.split())

            #obtendo as palavras-chave
            start = ['Proposta']
            end = [' ']
            keyword1 = get_keyword(start, end, text)

            if keyword1 == '':
                start = ['Proposta ']
                end = [' ']
                keyword1 = get_keyword(start, end, text)
            
            start = ['CPF:']
            end = ['Nacionalidade']
            keyword2 = get_keyword(start, end, text)

            start = ['Data Contrato ']
            end = [' ']
            keyword3 = get_keyword(start, end, text) 

            start = ['Quadra']
            end = ['Padrão']
            keyword4 = get_keyword(start, end, text)
 
            start = ['Lote']
            end = ['Área']
            keyword5 = get_keyword(start, end, text)

            unidade = 'Q' + keyword4 + 'L' + keyword5
            unidade = unidade.replace(' ', '')

            start = ['Preço']
            end = ['Entrada']
            keyword6 = get_keyword(start, end, text)

            # outra maneira de encontrar palavras pois a outra pode não funcionar com certas condições
            if text.find('Prestação 1 / '):
                pos = text.find('Prestação 1 / ')
                n_parcelas = text[pos + 13:]
                n_parcelas = n_parcelas.split(' ')
                n_parcelas = n_parcelas[1]
                n_parcelas = int(n_parcelas) 
                if n_parcelas > 1:
                    data_reajuste = text[pos + 17:]
                    data_reajuste = data_reajuste.split(' ')
                    data_reajuste = data_reajuste[1]
                    data_reajuste = modificarDatas(data_reajuste)
                else:
                    data_reajuste = ''



            # cria uma lista com as palavras extraidas do documento atual
            my_list = [keyword1, unidade, keyword2, keyword3, keyword6, data_reajuste]

            # transforma a  lista numa linha de dataframe
            my_list = pd.Series(my_list)

            # adiciona a lista no meu dataframe.
            my_dataframe = my_dataframe.append(my_list, ignore_index=True)

            print("Contratos extraidos com sucesso")


    # renomeando as colunas do dataframe
    my_dataframe = my_dataframe.rename(columns={0:'numero contrato', 1:'unidade', 2:'cliente cpf', 3:'data contrato', 4:'preco total', 5:'reajustar a partir de'})
                                                    

    # indicando onde o arquivo excel será exportado
    save_path = (r'')  # coloque aqui entre as aspas o caminho em que deseja salvar o arquivo
    os.chdir(save_path)

   # exporta o dataframe em .xlsx
    my_dataframe.to_excel('importacao_contratos.xlsx', sheet_name = 'my dataframe', index=False)
    print("")
    print(my_dataframe)

if __name__ == '__main__':
    main()