import os.path
import os
import mysql.connector
from oauth2client.service_account import ServiceAccountCredentials 
from googleapiclient.discovery import build
from src.database.env import config  
from src.app.descricao_produto import descricao_s_fornecedor
from src.app.limpar_descricao import limpar_descricao
from src.app.cotacao import cotacao
from src.app.utils.column_titles import column_titles
from src.api.google_sheets import get_client_secret_path

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


# Conectando ao banco de dados
def main():

    creds_path = get_client_secret_path()
    # creds = Credentials.from_service_account_file("client_secret.json", scopes=SCOPES)
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, SCOPES)


    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        conn = mysql.connector.connect(**config)

        print("Conexão ao banco de dados bem-sucedida!")
        print("----------------------------------------\n")
        print("       __DESCRIÇÂO SEM FORNECEDOR___ ")
        descricao_s_fornecedor(sheet)
        print("\n       __LIMPAR DESCRIÇÂO__")
        limpar_descricao(sheet)
        print("\n           __COTAÇÂO__")
        cotacao(sheet, conn, column_titles)

    except KeyboardInterrupt:
        print("Processo interrompido pelo usuário.")
        input("Pressione Enter para sair.")
        conn.close()
        
    finally:
        # Certifique-se de fechar a conexão quando terminar
        if 'conn' in locals() and conn.is_connected():
            conn.close()
            # input("\n\nPressione Enter para sair.")
if __name__ == "__main__":
    main()
