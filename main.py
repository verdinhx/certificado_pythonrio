from criar_certificados import gerar_certificado_massa
from enviar_email import enviar_certificados_em_massa

def solicitar_sn(pergunta: str) -> str:
    """Solicita entrada do usuário até que seja 's' ou 'n' (minúsculo ou maiúsculo).
    Retorna a resposta em minúsculas ('s' ou 'n').
    """
    while True:
        resp = input(pergunta).strip().lower()
        if resp in ('s', 'n'):
            return resp
        print("Resposta inválida. Digite 's' para sim ou 'n' para não.")

# Execução do código
if __name__ == "__main__":
    if solicitar_sn("Deseja gerar os certificados? (s/n) ") == 's':
        gerar_certificado_massa()
    if solicitar_sn("Deseja enviar os certificados por email? (s/n) ") == 's':
        enviar_certificados_em_massa()



