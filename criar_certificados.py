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

# --- Configuração Inicial e de Caminhos ---

# Obtém o diretório atual do script para resolver caminhos relativos de forma segura
DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))

# Pasta de saída: Certificados serão salvos aqui (será criada se não existir)
PASTA_SAIDA = os.path.join(DIR_SCRIPT, 'certificados')

# Caminho da imagem de fundo (assumindo que 'Arquivos' está no mesmo nível do script)
IMAGEM_FUNDO = os.path.join(DIR_SCRIPT, 'Arquivos', 'fundo.png')

# Caminho da planilha
CAMINHO_PLANILHA = 'participantes0825.xlsx'

# Dimensões A4 Paisagem para uso consistente
LARGURA_PAGINA, ALTURA_PAGINA = landscape(A4) 

# Margem desejada em todos os lados
MARGEM = 3.5 * cm

# Dados do evento (pode ser parametrizado conforme necessário)
data = "25 de Agosto de 2025"
carga_horaria = "3 horas"

# --- Estilos de Texto ---
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



# --- Função de Callback para Fundo (Background) ---
def desenhar_fundo(canvas, doc):
    """Função de callback que desenha a imagem de fundo em cada página."""
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


def gerar_certificado_unitario(nome_participante):
    """Cria e salva um único PDF para o participante na pasta de saída."""

    # 1. Define o nome único do arquivo de saída
    nome_arquivo_seguro = f"{nome_participante.replace(' ', '_').replace('.', '').replace(',', '')}_certificado.pdf"
    caminho_saida_pdf = os.path.join(PASTA_SAIDA, nome_arquivo_seguro)
    
    # 2. Configura o SimpleDocTemplate com as margens
    # topMargin, bottomMargin, leftMargin, rightMargin são definidos
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

    # 3. Adiciona o conteúdo formatado
    
    # Título
    story.append(Spacer(1, 2 * cm)) # Espaçamento superior extra para layout
    story.append(Paragraph("CERTIFICADO DE PARTICIPAÇÃO", styles['CertificadoTitulo']))
    story.append(Spacer(3, 2 * cm))

    # Corpo do certificado
    story.append(Paragraph(f"Certificamos que {nome_participante} participou do Meet Up da Comunidade PythOnRio, com carga horária de {carga_horaria}, realizado em {data}", styles['CertificadoCorpo']))
    story.append(Spacer(1, 0.5 * cm))

    # Rodapé
    story.append(Spacer(1, 0.5 * cm)) # Espaçamento antes do rodapé
    story.append(Paragraph("PythOnRio - Comunidade de Python do Rio de Janeiro", styles['CertificadoRodape']))
    
    
    # 4. Constrói o PDF
    # onFirstPage: desenha a imagem de fundo antes do conteúdo do Story
    doc.build(story, onFirstPage=desenhar_fundo, onLaterPages=desenhar_fundo)
    
    print(f"Certificado gerado com sucesso para: {nome_participante.strip()}")


def gerar_certificado_massa():
    """Função principal para gerar certificados em lote, lendo a planilha."""
    
    # 1. Cria a pasta de saída se não existir
    if not os.path.exists(PASTA_SAIDA):
        os.makedirs(PASTA_SAIDA)
        print(f"Pasta '{PASTA_SAIDA}' criada.")
    
    try:
        # 2. Leitura da planilha
        df = pd.read_excel(CAMINHO_PLANILHA)

        # 3. Validação da coluna
        COLUNA_NOME = 'Nome completo:'
        if COLUNA_NOME not in df.columns:
            print(f"ERRO: A coluna '{COLUNA_NOME}' não foi encontrada na planilha. Verifique o nome da coluna.")
            print(f"Colunas disponíveis: {df.columns.tolist()}")
            return

        # 4. Iteração e Geração
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
            



