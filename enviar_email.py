import pandas as pd
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from mimetypes import guess_type 
import re # Usado para limpeza de nome de arquivo

# ---- Configuração Inicial e de Caminhos ----

# ATENÇÃO: Verifique se este e-mail e a senha de app estão corretos.
# A senha de app é gerada nas configurações de segurança da sua conta Google.
EMAIL_REMETENTE = 'pythonrio.contato@gmail.com'
SENHA_APP = 'sua senha de app'  #MUDAR AQUI
SMTP_SERVIDOR = 'smtp.gmail.com'
SMTP_PORTA = 587
CC_EMAIL = 'pandaaquatico@gmail.com'
CAMINHO_PLANILHA = 'participantes0825.xlsx'

# Garante que o caminho da pasta 'certificados' está correto
PASTA_CERTIFICADOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'certificados')

# Constantes para as colunas do Excel
# É VITAL que essas constantes correspondam exatamente aos cabeçalhos da sua planilha!
COLUNA_NOME = 'Nome completo:'
COLUNA_EMAIL = 'E-mail:'

# --- Funções de Utilidade ---

def limpar_nome_para_arquivo(nome):
    """
    Limpa o nome do participante para que corresponda exatamente ao nome de arquivo PDF salvo.
    Essa função deve replicar o que foi feito no script de geração de certificados.
    """
    # Remove espaços, pontos e vírgulas para garantir a correspondência do nome do arquivo.
    nome_limpo = str(nome).strip().replace(' ', '_').replace('.', '').replace(',', '')
    return nome_limpo

# ---- Função para enviar e-mail com anexo ----
def enviar_email(destinatario, nome, pdf_filename):
    """Envia um único email com o PDF anexado."""
    
    # 1. Criação da mensagem MIME
    mensagem = MIMEMultipart()
    mensagem['From'] = EMAIL_REMETENTE
    mensagem['To'] = destinatario
    mensagem['Subject'] = 'Certificado de Participação - PythOnRio'

    if CC_EMAIL:
        mensagem['Cc'] = CC_EMAIL

    # Corpo do e-mail
    corpo_email = f"Olá {nome},\nSegue em anexo o seu certificado de participação do meet up da Comunidade PythOnRio.\n \nAtenciosamente, \nComunidade PythOnRio"
    mensagem.attach(MIMEText(corpo_email, 'plain'))
    
    # 2. Adiciona o anexo
    try:
        with open(pdf_filename, "rb") as f:
            pdf_data = f.read()
        
        # Como é um PDF, definimos o tipo MIME explicitamente
        attachment = MIMEApplication(pdf_data, _subtype='pdf')
        
        # Define o nome do arquivo a ser exibido no anexo do email
        attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_filename))
        mensagem.attach(attachment)

    except FileNotFoundError:
        print(f"ERRO: PDF não encontrado em {pdf_filename} para {nome}. Pulando este e-mail.")
        return
    except Exception as e:
        print(f"ERRO ao anexar o PDF {pdf_filename}: {e}. Pulando este e-mail.")
        return

    # 3. Preparação para o envio: lista de destinatários (To + Cc)
    recipients = [destinatario]
    if CC_EMAIL:
        recipients.append(CC_EMAIL)

    # 4. Envio do e-mail usando o bloco with para fechar a conexão automaticamente
    try:
        # Tenta conectar ao servidor SMTP
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_APP)
        
        # Envio
        server.sendmail(EMAIL_REMETENTE, recipients, mensagem.as_string())
        
        # A conexão é fechada automaticamente pelo 'with' ou no final da função
        
        cc_log = f" (CC: {CC_EMAIL})" if CC_EMAIL else ""
        print(f"SUCESSO: E-mail enviado para {destinatario}{cc_log} com anexo: {os.path.basename(pdf_filename)}")

    except smtplib.SMTPAuthenticationError:
        print("ERRO CRÍTICO de autenticação SMTP. Verifique o email e SENHA_APP.")
    except Exception as e:
        print(f"ERRO ao enviar e-mail para {destinatario}: {e}")

def enviar_certificados_em_massa():
    """Função principal para ler a planilha e orquestrar o envio de e-mails."""
    
    try:
        # 1. Leitura da planilha (DataFrame)
        df = pd.read_excel(CAMINHO_PLANILHA)

        # 2. Validação das colunas
        if COLUNA_NOME not in df.columns or COLUNA_EMAIL not in df.columns:
            print(f"ERRO: As colunas '{COLUNA_NOME}' e/ou '{COLUNA_EMAIL}' não foram encontradas na planilha.")
            print(f"Colunas disponíveis: {df.columns.tolist()}")
            return
        
        # 3. Inicia o loop de envio
        total_participantes = len(df.index)
        print(f"Iniciando o envio de e-mails para {total_participantes} participantes...")
        
        for index, row in df.iterrows():

            email_destinatario = row[COLUNA_EMAIL]
            nome_destinatario = str(row[COLUNA_NOME]).strip()
            
            # Verifica se o e-mail está ausente ou inválido
            if pd.isna(email_destinatario) or not nome_destinatario:
                print(f"AVISO: Linha {index + 2} (Nome: {nome_destinatario}) não possui e-mail ou nome válido. Pulando.")
                continue

            # 4. Gera o nome do arquivo PDF (usando a função de limpeza)
            nome_arquivo_pdf = f"{limpar_nome_para_arquivo(nome_destinatario)}_certificado.pdf"
            pdf_filename = os.path.join(PASTA_CERTIFICADOS, nome_arquivo_pdf)
            
            # 5. Envia o e-mail
            enviar_email(email_destinatario, nome_destinatario, pdf_filename)
            
        print("\nProcesso de envio de e-mails concluído.")

    except FileNotFoundError:
        print(f"ERRO: O arquivo '{CAMINHO_PLANILHA}' não foi encontrado. Verifique o caminho.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
