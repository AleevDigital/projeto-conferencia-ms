import openpyxl
import pandas as pd
import os
from openpyxl import load_workbook

def tratar_matriz(arquivo, tipo):
    if tipo == "d4":
        for nome_arquivo in os.listdir(fr'output\D4'):
            caminho_completo = os.path.join(fr'output\D4', nome_arquivo)
            if os.path.isfile(caminho_completo):
                os.remove(caminho_completo)
    elif tipo == "d2":
        for nome_arquivo in os.listdir(fr'output\D2'):
            caminho_completo = os.path.join(fr'output\D2', nome_arquivo)
            if os.path.isfile(caminho_completo):
                os.remove(caminho_completo)

    data = {
        'Conta Contábil': [],
        'Informações Complementares 1':[],
        'Tipo de Informação 1':[],
        'Informações Complementares 2':[],
        'Tipo de Informação 2':[],
        'Informações Complementares 3':[],
        'Tipo de Informação 3':[],
        'Informações Complementares 4':[],
        'Tipo de Informação 4':[],
        'Informações Complementares 5':[],
        'Tipo de Informação 5':[],
        'Informações Complementares 6':[],
        'Tipo de Informação 6':[],
        'Saldo Inicial':[],
        'Crédito':[],
        'Débito':[],
        'Saldo Final':[],
        'Naturezas':[]
    }

    base = pd.read_csv(arquivo,sep = ';', dtype = str, header = None, skiprows=1)
    filename = os.path.basename(arquivo).replace('csv','xlsx')

    base = base.iloc[1:]
    base.columns =['CONTA','IC','TIPO','IC2','TIPO2','IC3','TIPO3','IC4','TIPO4','IC5','TIPO5','IC6','TIPO6','VALOR','TIPO_VALOR','NATUREZA_VALOR']

    base = base.iloc[1:]
    base = base.fillna('-')

    df = base   
    # Agrupe os dados por colunas e crie DataFrames separados para cada grupo
    grouped = df.groupby(['CONTA', 'IC', 'TIPO', 'IC2', 'TIPO2', 'IC3', 'TIPO3', 'IC4', 'TIPO4','IC5','TIPO5','IC6','TIPO6'])
    dataframes = [group for _, group in grouped]

    for dataframe in dataframes:
    #PARA GRUPOS COM QUATRO OPERAÇÕES NA MESMA CONTA CONTÁBIL
        if len(dataframe) == 4:
            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[0] == 'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[0]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[0]}')
                        
            #Se o valor é um  Saldo Final
            if dataframe['TIPO_VALOR'].iloc[0] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[0]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[0]}')
            
            #Se o valor é um crédito ou um débito
            if dataframe['TIPO_VALOR'].iloc[0] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[0]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[0]}')

            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[1] ==  'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[1]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[1]}')
                        
            #Se o valor é um  **Saldo Final**
            if dataframe['TIPO_VALOR'].iloc[1] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[1]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[1]}')
            
            #Se o valor é um crédito ou um débito
            if dataframe['TIPO_VALOR'].iloc[1] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[1]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[1]}')
                        
            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[2] ==  'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[2][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[2]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[2]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[2][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[2] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[2]}')
                        
            #Se o valor é um  Saldo Final
            if dataframe['TIPO_VALOR'].iloc[2] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[2][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[2]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[2]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[2][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[2] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[2]}')
            
            #Se o valor é um crédito ou um débito
            if dataframe['TIPO_VALOR'].iloc[2] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[2][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[2]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[0]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[2][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[2] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[2]}')
                        
            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[3] ==  'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[3][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[3]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[3]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[0]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[3][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[3] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[3]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[3] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[3]}')
                        
            #Se o valor é um  Saldo Final
            if dataframe['TIPO_VALOR'].iloc[3] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[3][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[3]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[3]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[3]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[3]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[3][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[3] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[3]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[3] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[3]}')
            
            #Se o valor é um crédito ou um débito
            if dataframe['TIPO_VALOR'].iloc[3] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[3][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[3]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[3]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[3]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[3]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[3][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[3] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[3]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[3] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[3]}')
            data['Naturezas'].append(f'{dataframe["NATUREZA_VALOR"].iloc[0]},{dataframe["NATUREZA_VALOR"].iloc[2]},{dataframe["NATUREZA_VALOR"].iloc[3]},{dataframe["NATUREZA_VALOR"].iloc[1]}')       
            
            data['Conta Contábil'].append(dataframe['CONTA'].iloc[0])
            data['Informações Complementares 1'].append(dataframe['IC'].iloc[0])
            data['Tipo de Informação 1'].append(dataframe['TIPO'].iloc[0])
            data['Informações Complementares 2'].append(dataframe['IC2'].iloc[0])
            data['Tipo de Informação 2'].append(dataframe['TIPO2'].iloc[0])
            data['Informações Complementares 3'].append(dataframe['IC3'].iloc[0])
            data['Tipo de Informação 3'].append(dataframe['TIPO3'].iloc[0])
            data['Informações Complementares 4'].append(dataframe['IC4'].iloc[0])
            data['Tipo de Informação 4'].append(dataframe['TIPO4'].iloc[0])
            data['Informações Complementares 5'].append(dataframe['IC5'].iloc[0])
            data['Tipo de Informação 5'].append(dataframe['TIPO5'].iloc[0])
            data['Informações Complementares 6'].append(dataframe['IC6'].iloc[0])
            data['Tipo de Informação 6'].append(dataframe['TIPO6'].iloc[0])

        #PARA GRUPOS COM TRÊS OPERAÇÕES NA MESMA CONTA CONTÁBIL
        elif len(dataframe) == 3:
            
            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[0] ==  'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[0]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[0]}')
                        
            #Se o valor é um  Saldo Final
            if dataframe['TIPO_VALOR'].iloc[0] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[0]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[0]}')
            
            #Se o valor é um crédito ou um débito
            if dataframe['TIPO_VALOR'].iloc[0] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[0]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[0]}')
                        
            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[1] ==  'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[1]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[1]}')
                        
            #Se o valor é um  **Saldo Final**
            if dataframe['TIPO_VALOR'].iloc[1] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[1]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[1]}')
            
            #Se o valor é um crédito ou um débito
            if dataframe['TIPO_VALOR'].iloc[1] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[1]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[1]}')
                        
            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[2] ==  'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[2][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[2]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[2]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[2][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[2] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[2]}')
                        
            #Se o valor é um  Saldo Final
            if dataframe['TIPO_VALOR'].iloc[2] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[2][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[2]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[2]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[2][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[2] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[2]}')
            
            #Se o valor é um crédito ou um débito
            if dataframe['TIPO_VALOR'].iloc[2] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[2][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[2]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[2]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[2][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[2] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[2]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[2] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[2]}')

            max_length = max(len(data['Saldo Inicial']), len(data['Saldo Final']), len(data['Crédito']), len(data['Débito']))

            data['Saldo Inicial'] += ['-'] * (max_length - len(data['Saldo Inicial']))
            data['Saldo Final'] += ['-'] * (max_length - len(data['Saldo Final']))
            data['Crédito'] += ['-'] * (max_length - len(data['Crédito']))
            data['Débito'] += ['-'] * (max_length - len(data['Débito']))
                    
            data['Naturezas'].append(f'{dataframe["NATUREZA_VALOR"].iloc[0]},{dataframe["NATUREZA_VALOR"].iloc[2]},{dataframe["NATUREZA_VALOR"].iloc[1]}')
            data['Conta Contábil'].append(dataframe['CONTA'].iloc[0])
            data['Informações Complementares 1'].append(dataframe['IC'].iloc[0])
            data['Tipo de Informação 1'].append(dataframe['TIPO'].iloc[0])
            data['Informações Complementares 2'].append(dataframe['IC2'].iloc[0])
            data['Tipo de Informação 2'].append(dataframe['TIPO2'].iloc[0])
            data['Informações Complementares 3'].append(dataframe['IC3'].iloc[0])
            data['Tipo de Informação 3'].append(dataframe['TIPO3'].iloc[0])
            data['Informações Complementares 4'].append(dataframe['IC4'].iloc[0])
            data['Tipo de Informação 4'].append(dataframe['TIPO4'].iloc[0])
            data['Informações Complementares 5'].append(dataframe['IC5'].iloc[0])
            data['Tipo de Informação 5'].append(dataframe['TIPO5'].iloc[0])
            data['Informações Complementares 6'].append(dataframe['IC6'].iloc[0])
            data['Tipo de Informação 6'].append(dataframe['TIPO6'].iloc[0])

        #PARA GRUPOS COM DOIS OPERAÇÕES NA MESMA CONTA CONTÁBIL
        if len(dataframe) == 2:
            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[0] ==  'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[0]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[0]}')
                        
            #Se o valor é um  Saldo Final
            if dataframe['TIPO_VALOR'].iloc[0] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[0]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[0]}')
            
            #Se o valor é um crédito ou um débito
            if dataframe['TIPO_VALOR'].iloc[0] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[0][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[0]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[0]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[0][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[0] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[0]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[0] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[0]}')
                        
            #Se o valor é um  Saldo Inicial
            if dataframe['TIPO_VALOR'].iloc[1] ==  'beginning_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[1]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Inicial'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Inicial'].append(f'-{dataframe["VALOR"].iloc[1]}')
                        
            #Se o valor é um  **Saldo Final**
            if dataframe['TIPO_VALOR'].iloc[1] ==  'ending_balance':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[1]}')

                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Saldo Final'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Saldo Final'].append(f'-{dataframe["VALOR"].iloc[1]}')
            #Se o valor é um crédito ou um débito 
            if dataframe['TIPO_VALOR'].iloc[1] == 'period_change':
                #Se a conta é devedora
                if dataframe['CONTA'].iloc[1][0] in ['1', '3', '5', '7']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1]=='C':
                        #Caso o valor seja o crédito de uma conta devedora ele é negativo
                        data['Crédito'].append(f'-{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    elif dataframe['NATUREZA_VALOR'].iloc[1]=='D':
                        #Caso o valor seja um débito de uma conta devedora ele é positivo
                        data['Débito'].append(f'{dataframe["VALOR"].iloc[1]}')
                        
                #Se  a conta é credora
                elif dataframe['CONTA'].iloc[1][0] in ['2', '4' , '6' , '8']:
                    #Se o valor é um crédito
                    if dataframe['NATUREZA_VALOR'].iloc[1] =='C':
                        #Caso o valor seja o crédito  de uma conta credora ele é positivo
                        data['Crédito'].append(f'{dataframe["VALOR"].iloc[1]}')
                    #Se o valor é um débito
                    if  dataframe['NATUREZA_VALOR'].iloc[1] == 'D':
                        #Caso o valor seja o débito de uma conta credora ele é negativo
                        data['Débito'].append(f'-{dataframe["VALOR"].iloc[1]}')
            # Verificar o tamanho das listas e adicionar '0,00' onde necessário
            max_length = max(len(data['Saldo Inicial']), len(data['Saldo Final']), len(data['Crédito']), len(data['Débito']))

            data['Saldo Inicial'] += ['-'] * (max_length - len(data['Saldo Inicial']))
            data['Saldo Final'] += ['-'] * (max_length - len(data['Saldo Final']))
            data['Crédito'] += ['-'] * (max_length - len(data['Crédito']))
            data['Débito'] += ['-'] * (max_length - len(data['Débito']))
            
            data['Naturezas'].append(f'{dataframe["NATUREZA_VALOR"].iloc[0]},{dataframe["NATUREZA_VALOR"].iloc[1]}')
            data['Conta Contábil'].append(dataframe['CONTA'].iloc[0])
            data['Informações Complementares 1'].append(dataframe['IC'].iloc[0])
            data['Tipo de Informação 1'].append(dataframe['TIPO'].iloc[0])
            data['Informações Complementares 2'].append(dataframe['IC2'].iloc[0])
            data['Tipo de Informação 2'].append(dataframe['TIPO2'].iloc[0])
            data['Informações Complementares 3'].append(dataframe['IC3'].iloc[0])
            data['Tipo de Informação 3'].append(dataframe['TIPO3'].iloc[0])
            data['Informações Complementares 4'].append(dataframe['IC4'].iloc[0])
            data['Tipo de Informação 4'].append(dataframe['TIPO4'].iloc[0])
            data['Informações Complementares 5'].append(dataframe['IC5'].iloc[0])
            data['Tipo de Informação 5'].append(dataframe['TIPO5'].iloc[0])
            data['Informações Complementares 6'].append(dataframe['IC6'].iloc[0])
            data['Tipo de Informação 6'].append(dataframe['TIPO6'].iloc[0])    

    database = pd.DataFrame(data)
    
    database['Tipo de Informação 1'] = database['Tipo de Informação 1'].replace('PO','Poder ou Orgão').replace('FP','Atributo de Superávit Financeiro').replace('DC','Dívida Consolidada').replace('FR','Fonte ou Destinação de Recursos').replace('CF','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('CO','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('NR','Natureza da Receita').replace('ND','Natureza da Despesa').replace('FS','Classificação Funcional(Função e Subfunção)').replace('AI','Ano de Inscrição de Restos a Pagar').replace('ES','Despesas com MDE e APS')
    database['Tipo de Informação 2'] = database['Tipo de Informação 2'].replace('PO','Poder ou Orgão').replace('FP','Atributo de Superávit Financeiro').replace('DC','Dívida Consolidada').replace('FR','Fonte ou Destinação de Recursos').replace('CF','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('CO','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('NR','Natureza da Receita').replace('ND','Natureza da Despesa').replace('FS','Classificação Funcional(Função e Subfunção)').replace('AI','Ano de Inscrição de Restos a Pagar').replace('ES','Despesas com MDE e APS')
    database['Tipo de Informação 3'] = database['Tipo de Informação 3'].replace('PO','Poder ou Orgão').replace('FP','Atributo de Superávit Financeiro').replace('DC','Dívida Consolidada').replace('FR','Fonte ou Destinação de Recursos').replace('CF','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('CO','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('NR','Natureza da Receita').replace('ND','Natureza da Despesa').replace('FS','Classificação Funcional(Função e Subfunção)').replace('AI','Ano de Inscrição de Restos a Pagar').replace('ES','Despesas com MDE e APS')
    database['Tipo de Informação 4'] = database['Tipo de Informação 4'].replace('PO','Poder ou Orgão').replace('FP','Atributo de Superávit Financeiro').replace('DC','Dívida Consolidada').replace('FR','Fonte ou Destinação de Recursos').replace('CF','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('CO','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('NR','Natureza da Receita').replace('ND','Natureza da Despesa').replace('FS','Classificação Funcional(Função e Subfunção)').replace('AI','Ano de Inscrição de Restos a Pagar').replace('ES','Despesas com MDE e APS')
    database['Tipo de Informação 5'] = database['Tipo de Informação 5'].replace('PO','Poder ou Orgão').replace('FP','Atributo de Superávit Financeiro').replace('DC','Dívida Consolidada').replace('FR','Fonte ou Destinação de Recursos').replace('CF','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('CO','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('NR','Natureza da Receita').replace('ND','Natureza da Despesa').replace('FS','Classificação Funcional(Função e Subfunção)').replace('AI','Ano de Inscrição de Restos a Pagar').replace('ES','Despesas com MDE e APS')
    database['Tipo de Informação 6'] = database['Tipo de Informação 6'].replace('PO','Poder ou Orgão').replace('FP','Atributo de Superávit Financeiro').replace('DC','Dívida Consolidada').replace('FR','Fonte ou Destinação de Recursos').replace('CF','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('CO','Complemento da Fonte de Recursos ou Destinação de Recursos').replace('NR','Natureza da Receita').replace('ND','Natureza da Despesa').replace('FS','Classificação Funcional(Função e Subfunção)').replace('AI','Ano de Inscrição de Restos a Pagar').replace('ES','Despesas com MDE e APS')
    database['Saldo Inicial'] = database['Saldo Inicial'].replace('-','0.00').astype(float)
    database['Saldo Inicial'] = database['Saldo Inicial'].round(2)
    database['Crédito']=database['Crédito'].replace('-','0.00').astype(float)
    database['Débito']=database['Débito'].replace('-','0.00').astype(float)
    database['Saldo Final'] = database['Saldo Final'].replace('-','0.00').astype(float)

    colunas_desejadas= ['Conta Contábil','Tipo de Informação 1','Informações Complementares 1','Tipo de Informação 2','Informações Complementares 2','Tipo de Informação 3','Informações Complementares 3','Tipo de Informação 4','Informações Complementares 4','Tipo de Informação 5','Informações Complementares 5','Tipo de Informação 6','Informações Complementares 6','Saldo Inicial','Crédito','Débito','Saldo Final','Naturezas']

    database = database[colunas_desejadas]

    colunas_numericas = ['Saldo Inicial', 'Crédito', 'Débito', 'Saldo Final']
    database[colunas_numericas] = database[colunas_numericas].astype(float).round(2)

    database.to_excel(fr'output\{filename}', index = False)

    # Carregando o arquivo Excel
    workbook = openpyxl.load_workbook(fr'output\{filename}')
    sheet = workbook.active

    # Ajustando o tamanho das colunas com base no tamanho do conteúdo
    for column in sheet.columns:
        max_length = 0
        column_name = column[0].column_letter  # Nome da coluna (letra)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2  # Adicionando um espaço extra
        sheet.column_dimensions[column_name].width = adjusted_width
        

    # Salvando o arquivo modificado
    if tipo == "d2":
        workbook.save(fr'output\D2\{filename}')
        print(f'Arquivo Salvo: {filename}')
    
    elif tipo == "d4":
        workbook.save(fr'output\D4\{filename}')
        print(f'Arquivo Salvo: {filename}')

    return database