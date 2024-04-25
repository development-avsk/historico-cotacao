from src.api.config_sample import SAMPLE_SPREADSHEET_ID
import re
import time  # Importe a biblioteca time para pausar a execução
from datetime import datetime

def find_column_indices(sheet, column_titles):
    range_row1 = "Petrobras!1:1"
    result_row1 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_row1).execute()
    values_row1 = result_row1.get('values', [])
    
    if not values_row1:
        print('Nenhuma linha de título encontrada na planilha Petrobras.')
        return None
    
    column_indices = []
    for column_title in column_titles:
        try:
            column_index = values_row1[0].index(column_title)
            column_indices.append(column_index)
        except ValueError:
            print(f"O título '{column_title}' não foi encontrado nas colunas.")
            column_indices.append(None)
    
    return column_indices

def index_to_column(index):
    column = ""
    while index >= 0:
        column = chr(index % 26 + 65) + column
        index //= 26
        index -= 1
        if index < 0:
            break
    return column


def update_cells(sheet, range_values, body_values):
    batch_update_request = {
        'value_input_option': 'RAW',
        'data': [{
            'range': range_values,
            'values': body_values['values'],
        }]
    }
    
    request = sheet.spreadsheets().values().batchUpdate(
        spreadsheetId=SAMPLE_SPREADSHEET_ID,
        body=batch_update_request
    )
    
    try:
        response = request.execute()
        print("Atualização em lote das células concluída com sucesso.")
    except Exception as e:
        print("Erro ao atualizar as células em lote:", str(e))
    
    time.sleep(2) 


def descricao_s_fornecedor(sheet):
    column_titles = [
        "DESCRIÇÃO COMPLETA DO ITEM",
        "Descrição Sem Fornecedor:",
    ]
    column_indices = find_column_indices(sheet, column_titles)
    
    # Verificar se todos os títulos foram encontrados
    if None in column_indices:
        print("Alguns títulos não foram encontrados.")
        return
    
    
    # Atribuir os índices das colunas a variáveis
    indice_descricao_completa = column_indices[0]
    indice_descricao_s_fornecedor = column_indices[1]


    # Convertendo o índice da coluna para o nome da coluna
    nome_coluna_descricao_completa = index_to_column(indice_descricao_completa)
    nome_coluna_descricao_s_fornecedor = index_to_column(indice_descricao_s_fornecedor)


    # Recuperar os valores para cada coluna
    range_data_descricao_completa = f"Petrobras!{nome_coluna_descricao_completa}2:{nome_coluna_descricao_completa}"
    result_data_descricao_completa = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_data_descricao_completa).execute()
    values_data_descricao_completa = result_data_descricao_completa.get('values', [])

    range_data_descricao_s_fornecedor = f"Petrobras!{nome_coluna_descricao_s_fornecedor}2:{nome_coluna_descricao_s_fornecedor}"
    result_data_descricao_s_fornecedor = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_data_descricao_s_fornecedor).execute()
    values_data_descricao_s_fornecedor = result_data_descricao_s_fornecedor.get('values', [])


    

    # Verificar se há dados na coluna "Descrição Sem Fornecedor"
    if not values_data_descricao_completa:
        print("Nenhum dado encontrado.")
        return
    
    batch_update = {'valueInputOption': 'RAW', 'data': []}

    # Imprimir os valores da coluna "Descrição Sem Fornecedor" e "DESCRIÇÃO COMPLETA DO ITEM"
    for i, (descricao_completa) in enumerate(zip(
        values_data_descricao_completa,
    ), start=2):

        if not descricao_completa or descricao_completa[0][0].find('\u200B') != -1:
            continue

        elif all([descricao_completa]):     
            descricao_completa_valor = descricao_completa[0][0]
            descricao_completa_s_fornecedor = descricao_completa[0][0].split(";Tp:")[0].strip()
            print(descricao_completa_valor)

            batch_update['data'].append({
                'range': f'Petrobras!{nome_coluna_descricao_s_fornecedor}{i}',
                'values': [[descricao_completa_s_fornecedor]]  
            })

            batch_update['data'].append({
                'range': f'Petrobras!{nome_coluna_descricao_completa}{i}',
                'values': [[f"{descricao_completa_valor}\u200B"]]  
            })

        
              
        # print(f"atualizando a planilha na linha - {i} - OK")


    # Verifica se há dados para atualizar na planilha antes de realizar a atualização em lote
    if batch_update['data']:
        # Atualização em lote das células
        sheet.values().batchUpdate(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            body=batch_update
        ).execute()
    else:
        print("Valores já estão atualizados na planilha.")   

                



      

   




