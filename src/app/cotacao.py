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


def get_column_values(sheet, column_name, sample_spreadsheet_id):
    range_data = f"Petrobras!{column_name}2:{column_name}"
    result_data = sheet.values().get(spreadsheetId=sample_spreadsheet_id, range=range_data).execute()
    values_data = result_data.get('values', [])
    return values_data


def cotacao(sheet, conn, column_titles):
    banco_dados = "cotacao_dev"
    column_indices = find_column_indices(sheet, column_titles)
    
    # Verificar se todos os títulos foram encontrados
    if None in column_indices:
        print("Alguns títulos não foram encontrados.")
        return
    
    data_atual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
   # Atribuir os índices das colunas a variáveis
    indice_descricao, indice_preco_custo, indice_ipi, indice_icms_st, indice_frete_fornecedor, \
    indice_frete_cliente, indice_margem, indice_data_cotacao_bd, indice_data_hist_cotacao_bd, \
    indice_data_responsavel, indice_data_status, indice_data_link, indice_data_prazo_fornecedor = column_indices

    # Convertendo os índices da coluna para o nome da coluna
    nome_coluna_descricao = index_to_column(indice_descricao)
    nome_coluna_preco_custo = index_to_column(indice_preco_custo)
    nome_coluna_ipi = index_to_column(indice_ipi)
    nome_coluna_icms_st = index_to_column(indice_icms_st)
    nome_coluna_frete_fornecedor = index_to_column(indice_frete_fornecedor)
    nome_coluna_frete_cliente = index_to_column(indice_frete_cliente)
    nome_coluna_margem = index_to_column(indice_margem)
    nome_coluna_data_cotacao_bd = index_to_column(indice_data_cotacao_bd)
    nome_coluna_data_hist_cotacao_bd = index_to_column(indice_data_hist_cotacao_bd)
    nome_coluna_data_responsavel = index_to_column(indice_data_responsavel)
    nome_coluna_data_status = index_to_column(indice_data_status)
    nome_coluna_data_link = index_to_column(indice_data_link)
    nome_coluna_data_prazo_fornecedor = index_to_column(indice_data_prazo_fornecedor)

    # Recuperar os valores para cada coluna
    values_data_descricao = get_column_values(sheet, nome_coluna_descricao, SAMPLE_SPREADSHEET_ID)
    values_data_preco_custo = get_column_values(sheet, nome_coluna_preco_custo, SAMPLE_SPREADSHEET_ID)
    values_data_ipi = get_column_values(sheet, nome_coluna_ipi, SAMPLE_SPREADSHEET_ID)
    values_data_icms_st = get_column_values(sheet, nome_coluna_icms_st, SAMPLE_SPREADSHEET_ID)
    values_data_frete_fornecedor = get_column_values(sheet, nome_coluna_frete_fornecedor, SAMPLE_SPREADSHEET_ID)
    values_data_frete_cliente = get_column_values(sheet, nome_coluna_frete_cliente, SAMPLE_SPREADSHEET_ID)
    values_data_margem = get_column_values(sheet, nome_coluna_margem, SAMPLE_SPREADSHEET_ID)
    values_data_responsavel = get_column_values(sheet, nome_coluna_data_responsavel, SAMPLE_SPREADSHEET_ID)
    values_data_link = get_column_values(sheet, nome_coluna_data_link, SAMPLE_SPREADSHEET_ID)
    values_data_prazo_fornecedor = get_column_values(sheet, nome_coluna_data_prazo_fornecedor, SAMPLE_SPREADSHEET_ID)
    

    # Verificar se há dados na coluna "Descrição Sem Fornecedor"
    if not values_data_descricao:
        print("Nenhum dado encontrado.")
        return
    
    batch_update = {'valueInputOption': 'RAW', 'data': []}
    
    # Imprimir os valores da coluna "Descrição Sem Fornecedor" e "Preço Custo do Produto"
    for i, (descricao, preco_custo, ipi, icms_st, frete_fornecedor, frete_cliente, margem,responsavel, link, prazo_fornecedor) in enumerate(zip(
        values_data_descricao,
        values_data_preco_custo,
        values_data_ipi,
        values_data_icms_st,
        values_data_frete_fornecedor,
        values_data_frete_cliente,
        values_data_margem,
        values_data_responsavel,
        values_data_link,
        values_data_prazo_fornecedor
        ), start=2):
        cursor = conn.cursor()

        if all([descricao, preco_custo, ipi, icms_st, frete_fornecedor, frete_cliente, margem]):    
            cursor.execute(f"SELECT COUNT(*) FROM {banco_dados} WHERE descricao = %s", (descricao[0],))
            num_linhas = cursor.fetchone()[0]
            
            if num_linhas < 4:
                preco_custo = preco_custo[0].replace('R$', '')
                icms_st = icms_st = icms_st[0].replace('R$', '')
                frete_fornecedor = frete_fornecedor[0].replace('R$', '')
                frete_cliente = frete_cliente[0].replace('R$', '')

                # print(preco_custo, ipi[0], icms_st, frete_fornecedor, frete_cliente, margem[0])
                cursor.execute(f"SELECT * FROM {banco_dados} WHERE descricao = %s AND preco_custo = %s AND ipi = %s AND icms_st = %s AND frete_fornecedor = %s AND frete_cliente = %s AND margem = %s", 
                            (descricao[0], preco_custo, ipi[0], icms_st, frete_fornecedor , frete_cliente, margem[0]))

                result_conection = cursor.fetchone()
                if not result_conection:
                    print(f"A descrição '{descricao[0]}' não existe na tabela 'cotacao'.")
                    # Insere os dados na tabela cotacao

                    descricao_value = descricao[0] if descricao else None
                    ipi_value = ipi[0] if ipi else None
                    margem_value = margem[0] if margem else None
                    responsavel_value = responsavel[0] if responsavel else None
                    link_value = link[0] if link else None
                    prazo_value = prazo_fornecedor[0] if prazo_fornecedor else None

                    # Executa a inserção na tabela
                    cursor.execute(f"INSERT INTO {banco_dados} (descricao, preco_custo, ipi, icms_st, frete_fornecedor, frete_cliente, margem, data, responsavel, link, prazo) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                                (descricao_value, preco_custo, ipi_value, icms_st, frete_fornecedor, frete_cliente, margem_value, data_atual, responsavel_value, link_value, prazo_value))
                    conn.commit()
                    print(f"Dados da linha {i} inseridos com sucesso na tabela 'cotacao'.")

            else:
                # Identifica o registro mais antigo com base na data de inserção
                cursor.execute(f"SELECT id FROM {banco_dados} WHERE descricao = %s ORDER BY data ASC LIMIT 1", (descricao[0],))
                id_antigo = cursor.fetchone()[0]
                
                # Atualiza o registro mais antigo com os novos valores
                cursor.execute(f"UPDATE {banco_dados} SET preco_custo = %s, ipi = %s, icms_st = %s, frete_fornecedor = %s, frete_cliente = %s, margem = %s, data = %s , responsavel = %s, link = %s, prazo = %s WHERE id = %s",
                            (preco_custo[0], ipi[0], icms_st[0], frete_fornecedor[0], frete_cliente[0], margem[0], data_atual, responsavel[0], link[0], prazo_fornecedor[0], id_antigo))
                conn.commit(),
                print(f"Valor mais antigo atualizado com sucesso para a descrição '{descricao[0]}'.")

        if all([descricao]):    
           if not (preco_custo and ipi and icms_st and frete_fornecedor and frete_cliente and margem):  

            cursor.execute(f"SELECT * FROM {banco_dados} WHERE descricao = %s ORDER BY data DESC LIMIT 1", (descricao[0],))

            descricao_tabela_existe = cursor.fetchall()
            
            if descricao_tabela_existe:
                # Obtenha os valores encontrados no banco de dados
                cotacao_info = descricao_tabela_existe[0]

                # Prepare os valores para atualização na planilha
                update_value_preco_custo = cotacao_info[2].replace('R$', '')
                update_value_ipi = cotacao_info[3]
                update_value_icms_st = cotacao_info[4].replace('R$', '')
                update_value_frete_fornecedor = cotacao_info[5].replace('R$', '')
                update_value_frete_cliente = cotacao_info[6].replace('R$', '')
                update_value_margem = cotacao_info[7]
                update_value_data_cotacao_bd = cotacao_info[8]
                update_value_data_responsavel = cotacao_info[9]
                update_value_data_link = cotacao_info[10]
                update_value_data_prazo_fornecedor = cotacao_info[11]

                atualizacoes = [
                    (nome_coluna_preco_custo, update_value_preco_custo),
                    (nome_coluna_ipi, update_value_ipi),
                    (nome_coluna_icms_st, update_value_icms_st),
                    (nome_coluna_frete_fornecedor, update_value_frete_fornecedor),
                    (nome_coluna_frete_cliente, update_value_frete_cliente),
                    (nome_coluna_margem, update_value_margem),
                    (nome_coluna_data_cotacao_bd, update_value_data_cotacao_bd),
                    (nome_coluna_data_responsavel, update_value_data_responsavel),
                    (nome_coluna_data_status, "Historico-Robo"),
                    (nome_coluna_data_link, update_value_data_link),
                    (nome_coluna_data_prazo_fornecedor, update_value_data_prazo_fornecedor)
                ]

                for coluna, valor in atualizacoes:
                    batch_update['data'].append({
                        'range': f'Petrobras!{coluna}{i}',
                        'values': [[valor]]  
                    })


                cursor.execute(f"SELECT  preco_custo, ipi, icms_st, frete_fornecedor, frete_cliente, margem, data FROM {banco_dados} WHERE descricao = %s ORDER BY data DESC", (descricao[0],))
                historico_cotacao = cursor.fetchall()

                valores_concatenados = []

                if historico_cotacao:
                    for cotacao_info in historico_cotacao:
                        preco_custo, ipi, icms_st, frete_fornecedor, frete_cliente, margem, data = cotacao_info

                        valores_completos = f"preco_custo:{preco_custo} ipi:{ipi} icms_st:{icms_st} frete_fornecedor:{frete_fornecedor} frete_cliente:{frete_cliente} margem:{margem} data:{data}"
                        
                        valores_concatenados.append(valores_completos)

                
                valores_totais = ' ------ '.join(valores_concatenados)

                batch_update['data'].append({
                    'range': f'Petrobras!{nome_coluna_data_hist_cotacao_bd}{i}',
                    'values': [[valores_totais]]  
                })
                
                print(f"atualizando a planilha na linha - {i} - OK")

         
    # Verifica se há dados para atualizar na planilha antes de realizar a atualização em lote
    if batch_update['data']:
        # Atualização em lote das células
        sheet.values().batchUpdate(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            body=batch_update
        ).execute()
    else:
        print("Valores já estão atualizados na planilha.")   

                

