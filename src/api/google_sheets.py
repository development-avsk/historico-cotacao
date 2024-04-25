import os

def get_client_secret_path():

    # Obtém o diretório do script atual
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Constrói o caminho para o arquivo client_secret.json
    client_secret_path = os.path.join(script_dir, "client_secret.json")

    # Retorna o caminho
    return client_secret_path


