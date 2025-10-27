from criar_certificados import gerar_certificado_massa
from enviar_email import enviar_certificados_em_massa
import os
import pandas as pd

# Execução do código
if __name__ == "__main__":
    gerar_certificado_massa()
    enviar_certificados_em_massa()



