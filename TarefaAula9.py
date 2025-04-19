import pandas as pd
import math
import os

def limpar_tela():
    if os.name == "nt":
        os.system('cls')
    elif os.name == "posix":
        os.system('clear')

def abrir_arquivo():
    limpar_tela()
    print('Buscando e abrindo arquivo "ipca_202503SerieHist.xls"\n')
    xls = pd.ExcelFile('ipca_202503SerieHist.xls')
    data_frame = pd.read_excel(xls, 'Série Histórica IPCA')
    print(data_frame)
    print('\nArquivo aberto com sucesso.') 
    input('\nPressione Enter para continuar...')
    return data_frame

def exibir_um_problema(data_frame, descricao, linha_inicial, linha_final):
    limpar_tela()
    print('Analisando os problemas do arquivo\n')
    print(df.iloc[linha_inicial:linha_final])
    print(f'\n{descricao}')
    input('\nPressione Enter para continuar...')

def exibir_problemas_arquivo(data_frame):
    exibir_um_problema(data_frame, '1) Os títulos das colunas estão distribuído entre as linhas 2 a 5.', 2, 6)
    exibir_um_problema(data_frame, '2) As 7 primeiras linhas não contém dados úteis.', 0, 7)
    exibir_um_problema(data_frame, '3) O cabeçalho do arquivo se repete ao longo dele.', 71, 79)
    exibir_um_problema(data_frame, '4) Existe uma linha em branco separando um ano do próximo.', 7, 21)
    exibir_um_problema(data_frame, '5) Nem toda linha tem a coluna "Ano" preenchida.', 7, 15)

def modificar_nome_colunas(data_frame):
    limpar_tela()
    print('Modificando o nome das colunas de forma explícita, pois seria complicado buscar essa informação da planilha.\n')
    df.columns = ['Ano', 'Mes', 'Numero indice', 'Variacao no Mes', 'Variacao 3 meses', 'Variacao 6 meses', 'Variacao no ano', 'Variacao 12 meses']
    print(df.head(10))
    print('\nNome das colunas modificado com sucesso.') 
    input('\nPressione Enter para continuar...')

def remover_linhas_iniciais(data_frame):
    limpar_tela()
    print('Removendo as 7 primeiras linhas, pois elas não contém dados úteis.\n')
    data_frame.drop(range(0, 7), inplace=True)
    data_frame = data_frame.reset_index(drop=True)
    print(data_frame.head(10))
    print('\nRemoção de linhas executada com sucesso.') 
    input('\nPressione Enter para continuar...')

def remover_cabecalhos_meio_arquivo(data_frame):
    limpar_tela()
    print('Removendo os cabeçalhos que estão no meio do arquivo.\n')
    for i, row in data_frame.iterrows():
        if row['Ano'] == 'SÉRIE HISTÓRICA DO IPCA':
            data_frame.drop(range(i, i+7), inplace=True)
            data_frame.drop(range(i-2, i), inplace=True)    
    data_frame = data_frame.reset_index(drop=True)
    with pd.option_context('display.max_rows', None):
        print(data_frame.iloc[71:79])
    print('\nRemoção dos cabeçalhos do meio do arquivo realizada com sucesso.')
    input('\nPressione Enter para continuar...')

def remover_linha_em_branco(data_frame):
    limpar_tela()
    print('Removendo as linhas em branco que separam um ano do outro.\n')
    for i, row in data_frame.iterrows():
        if isinstance(row['Mes'], (float)) and math.isnan(row['Mes']):
            data_frame.drop(i, inplace=True)
    data_frame = data_frame.reset_index(drop=True)
    print(data_frame.iloc[7:21])
    print('\nRemoção das linhas em branco que separam os anos realizada com sucesso.')
    input('\nPressione Enter para continuar...')

def alimentar_coluna_ano(data_frame):
    limpar_tela()
    print('Alimentando a coluna "Ano" em todas as linhas.\n')
    anoAtual = 0    
    for i, row in df.iterrows():    
        if not math.isnan(row['Ano']):
            anoAtual = row['Ano']
        else:
            row['Ano'] = anoAtual
    print(data_frame.head(10))
    print('\nAjuste da coluna "Ano" realizado com sucesso.')
    input('\nPressione Enter para continuar...')

def analisar_tipo_dados_colunas(data_frame):
    limpar_tela()
    print('Analisando o tipo de cada coluna da planilha.\n')
    print(data_frame.dtypes)
    print('\nPode-se notar que todas as colunas estão com o tipo "object".')
    input('\nPressione Enter para continuar...')

def corrigir_tipo_dados_colunas(data_frame):
    limpar_tela()
    print('Corrigindo o tipo de cada coluna da planilha.\n')
    data_frame['Ano'] = data_frame['Ano'].astype(int)
    data_frame['Mes'] = data_frame['Mes'].astype('string')
    data_frame['Numero indice'] = data_frame['Numero indice'].astype(float)
    data_frame['Variacao no Mes'] = data_frame['Variacao no Mes'].astype(float)
    data_frame['Variacao 3 meses'] = data_frame['Variacao 3 meses'].astype(float)
    data_frame['Variacao 6 meses'] = data_frame['Variacao 6 meses'].astype(float)
    data_frame['Variacao no ano'] = data_frame['Variacao no ano'].astype(float)
    data_frame['Variacao 12 meses'] = data_frame['Variacao 12 meses'].astype(float)
    print(data_frame.dtypes)
    print('\nCorreção do tipo de dados das colunas da planilha realizado com sucesso.".')
    input('\nPressione Enter para continuar...')

def exibir_shape_dataframe(data_frame):
    limpar_tela()
    print('Exibindo shape da planilha.\n')
    print(data_frame.shape)
    input('\nPressione Enter para continuar...')

def exportar_dataframe(data_frame):
    limpar_tela()
    print('Exportando dataframe para o arquivo "serie-histórica-ipca.xlsx".\n')
    data_frame.to_excel('serie-histórica-ipca.xlsx')

    print('Exportando dataframe para o arquivo "serie-histórica-ipca.csv".\n')
    data_frame.to_csv('serie-histórica-ipca.csv')

    print('Exportando dataframe para o arquivo "serie-histórica-ipca.json".\n')
    data_frame.to_json('serie-histórica-ipca.json')

    print('Exportações realizadas com sucesso.\n')
    input('Pressione Enter para finalizar o script...')

if __name__ == '__main__':
    df = abrir_arquivo()      
    exibir_problemas_arquivo(df)
    modificar_nome_colunas(df)
    remover_linhas_iniciais(df)
    remover_cabecalhos_meio_arquivo(df)
    remover_linha_em_branco(df)
    alimentar_coluna_ano(df)
    analisar_tipo_dados_colunas(df)
    corrigir_tipo_dados_colunas(df)
    exibir_shape_dataframe(df)
    exportar_dataframe(df)    