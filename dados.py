import pandas as pd
import os

# --- Função para coletar dados do evento ---
def coletar_dados():
    """Coleta os dados do evento via input do usuário e retorna um dicionário de configuração."""
    print("--- Configuração do Evento PythOnRio ---")
    
    while True:
        tipo = input("Digite o tipo de meetup: (1) Meetup ou (2) Tutorial: ")
        if tipo in ["1", "2"]: break
        print("Entrada inválida. Digite 1 ou 2.")

    if tipo == "2":
        titulo_evento = input("Digite o título do tutorial: ")
        texto_evento = f'Tutorial "{titulo_evento}"'
    else:
        texto_evento = "Meetup da Comunidade PythOnRio"

    data_evento = input("Digite a data do evento (ex: 11 de Abril de 2026): ")
    carga_horaria = input("Digite a carga horária (ex: 2 horas): ")
    caminho_planilha = input("Digite o nome da planilha (ex: participantes0426.xlsx): ")

    # Validação básica de arquivo
    if not os.path.exists(caminho_planilha):
        print(f"AVISO: O arquivo {caminho_planilha} não foi encontrado no diretório atual!")

    return {
        "TEXTO_EVENTO": texto_evento,
        "DATA_EVENTO": data_evento,
        "CARGA_HORARIA": carga_horaria,
        "CAMINHO_PLANILHA": caminho_planilha
    }

# Executa ao importar
CONFIG = coletar_dados()