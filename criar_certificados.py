import pandas as pd
import os
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus.flowables import Image
from dados import CONFIG

# --- Configuração Inicial e de Caminhos ---

# Obtém o diretório atual do script para os caminhos relativos
DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))

# Pasta 'certificados' 
PASTA_SAIDA = os.path.join(DIR_SCRIPT, 'certificados')

# Caminho da imagem de fundo 
IMAGEM_FUNDO = os.path.join(DIR_SCRIPT, 'Arquivos', 'fundo.png')

# Caminho para a planilha e dados do evento
CAMINHO_PLANILHA = CONFIG['CAMINHO_PLANILHA']
data = CONFIG['DATA_EVENTO']
carga_horaria = CONFIG['CARGA_HORARIA']
texto_evento_completo = CONFIG['TEXTO_EVENTO']

# --- Configurações de layout e estilos ---
# Dimensões A4 Paisagem 
LARGURA_PAGINA, ALTURA_PAGINA = landscape(A4) 

# Margem em todos os lados
MARGEM = 3.5 * cm

# Estilos de texto
styles = getSampleStyleSheet()

styles.add(ParagraphStyle(name='CertificadoTitulo',
                         parent=styles['Normal'],
                         fontName='Helvetica-bold',
                         fontSize=24,
                         alignment=1, # Centro
                         spaceAfter=0.2 * cm,
                         textColor=HexColor("#185666")))

styles.add(ParagraphStyle(name='CertificadoCorpo',
                         parent=styles['Normal'],
                         fontName='Helvetica',
                         fontSize=16,
                         leading=30,
                         alignment=1, # Centro
                         spaceAfter=0,
                         textColor=HexColor("#185666")))

styles.add(ParagraphStyle(name='CertificadoRodape',
                         parent=styles['Normal'],
                         fontName='Helvetica',
                         fontSize=12,
                         alignment=1, # Centro
                         spaceAfter=0,
                         textColor=HexColor("#185666")))

# --- Função para o background ---
def desenhar_fundo(canvas, doc):
    """Função que desenha a imagem de fundo em cada página."""
    try:
        canvas.drawImage(
            IMAGEM_FUNDO,
            x=0,
            y=0,
            width=LARGURA_PAGINA,
            height=ALTURA_PAGINA,
            mask='auto'
        )
    except Exception as e:
        print(f"ATENÇÃO: Erro ao carregar a imagem de fundo em '{IMAGEM_FUNDO}': {e}")
        # Desenha um quadrado cinza sutil se a imagem falhar
        canvas.setFillColor(colors.lightgrey)
        canvas.rect(0, 0, LARGURA_PAGINA, ALTURA_PAGINA, fill=1, stroke=0)

# --- Função para gerar um certificado individual ---
def gerar_certificado_unitario(nome_participante):
    """Cria e salva um único PDF para o participante na pasta de saída."""

    # Define o nome do arquivo de saída
    nome_arquivo_seguro = f"{nome_participante.replace(' ', '_').replace('.', '').replace(',', '')}_certificado.pdf"
    caminho_saida_pdf = os.path.join(PASTA_SAIDA, nome_arquivo_seguro)
    
    # Configura o SimpleDocTemplate com as margens
    doc = SimpleDocTemplate(
        caminho_saida_pdf,
        pagesize=landscape(A4),
        leftMargin=MARGEM,
        rightMargin=MARGEM,
        topMargin=MARGEM,
        bottomMargin=MARGEM
    )
    
    # Conteúdo do documento (Story)
    story = []

    # Título
    story.append(Spacer(1, 2 * cm)) 
    story.append(Paragraph("CERTIFICADO DE PARTICIPAÇÃO", styles['CertificadoTitulo']))
    story.append(Spacer(3, 2 * cm))

    # Corpo do certificado
    story.append(Paragraph(f"Certificamos que {nome_participante} participou do {texto_evento_completo} da Comunidade PythOnRio, com carga horária de {carga_horaria}, realizado em {data}", styles['CertificadoCorpo']))
    story.append(Spacer(1, 0.5 * cm))

    # Rodapé
    story.append(Spacer(1, 0.5 * cm)) # Espaçamento antes do rodapé
    story.append(Paragraph("PythOnRio - Comunidade de Python do Rio de Janeiro", styles['CertificadoRodape']))
    
    # Constrói o PDF
    doc.build(story, onFirstPage=desenhar_fundo, onLaterPages=desenhar_fundo)
    print(f"Certificado gerado com sucesso para: {nome_participante.strip()}")

# --- Função principal para gerar certificados em massa ---
def gerar_certificado_massa():
    """Função principal para gerar certificados em lote, para todos os participantes presentes."""
    
    # Cria a pasta de certificados se não existir
    if not os.path.exists(PASTA_SAIDA):
        os.makedirs(PASTA_SAIDA)
        print(f"Pasta '{PASTA_SAIDA}' criada.")
    
    try:
        # Leitura e manipulação da planilha
        df = pd.read_excel(CAMINHO_PLANILHA)
        df = df[df["Presença:"] == "Sim"]
        COLUNA_NOME = 'Nome completo:'
        if COLUNA_NOME not in df.columns:
            print(f"ERRO: A coluna '{COLUNA_NOME}' não foi encontrada na planilha. Verifique o nome da coluna.")
            print(f"Colunas disponíveis: {df.columns.tolist()}")
            return

        # Iteração e geração
        nomes_para_certificar = df[COLUNA_NOME].astype(str).dropna().unique()

        if len(nomes_para_certificar) == 0:
            print("ATENÇÃO: Nenhuma linha válida foi encontrada na coluna de nomes para gerar certificados.")
            return

        for nome_participante in nomes_para_certificar:
            gerar_certificado_unitario(nome_participante)
        
        print(f"\n 🐍✨ Processo concluído! Total de {len(nomes_para_certificar)} certificados gerados ✨🐍") 
        
    except FileNotFoundError:
        print(f"ERRO: O arquivo '{CAMINHO_PLANILHA}' não foi encontrado. Verifique o caminho.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
            



