# Gerar certificados - PythOnRio

## Sumário
1. [Descrição](#Descrição)
2. [Uso](#Uso)
   - [Sobre os arquivos](#Sobre-os-arquivos)
   - [Configurações](#Configurações)

3. [Instalação](#Instalação)


## Descrição

Esse repositório contém os arquivos necessários para gerar os certificados dos encontros da comunidade PythOnRio.

## Uso
Baixe os arquivos deste repositório. No seu compilador, rode apenas o `main.py`. No terminal aparecerá alguns inputs necessários para as configurações iniciais do evento. O primeiro passo é selecionar se você quer gerar um certificado para tutorial ou meetup.
```
--- Configuração do Evento PythOnRio ---
Digite o tipo de meetup: (1) Meetup ou (2) Tutorial:
```
Caso a opção escolhida tenha sido "Tutorial":
```
Digite o título do tutorial:
```
Tendo selecionado qualquer uma das opções, no terminal irá aparecer as seguintes instruções:
```
Digite a data do evento (ex: 11 de Abril de 2026):
Digite a carga horária (ex: 2 horas):
Digite o nome da planilha (ex: participantes0426.xlsx):
```
Se a planilha não for encontrada aparecerá um aviso:
```
AVISO: O arquivo  não foi encontrado no diretório atual!
```
Finalmente, o terminal pede para gerar os certificados. Você pode optar por sim (s) ou não (n).
```
Deseja gerar os certificados? (s/n)
```
Caso selecione sim e a planilha exista, espere que gere os certificados até a mensagem de conclusão:
```
 🐍✨ Processo concluído! Total de 2 certificados gerados ✨🐍
```
Após o "sim" ou "não" para a geração dos certyificados, você poderá escolher enviar os certificados por email ou não.
```
Deseja enviar os certificados por email? (s/n) 
```
Se "sim" for selecionado, espere enviar os emais até o aviso de conclusão:
```
 🐍✨ Processo de envio de e-mails concluído  🐍✨ 
```

### Sobre os arquivos
* [**Arquivos/**](Arquivos) - Nesta pasta você irá encontrar o arquivo de imagem [**fundo.png**](Arquivos/fundo.png) que será o fundo do certificado. 
* [**criar_certificados.py**](criar_certificados.py) - Nesse script será onde os certificados serão gerados. Todo o layout é construído aqui. Título, corpo do texto e outras informações são adicionadas aqui. 
* [**dados.py**](dados.py) - Esse script faz a entrada de dados do evento atrávés de inputs.
* [**enviar_email.py**](enviar_email.py) - Esse script faz a configuração de email e os envios do certificado. É possível adicionar um email para enviar cópias (CC). 
* [**main.py**](main.py) - É o script de execução. Para coletar os dados, gerar os certificados e enviar emails, apenas esse script deverá ser compilado.
* [**participantes0526.xlsx**](participantes0526.xlsx) - Essa é uma planilha de exemplo com a lista de participantes e lista de presença. 


### Configurações

No arquivo [**enviar_email.py**](enviar_email.py), modifique a "configuração inicial do email". 
~~~python
# --- Configuração incial do email ---
# ATENÇÃO: Modifique o e-mail e a senha de app do remetente.
EMAIL_REMETENTE = 'pythonrio.contato@gmail.com'
SENHA_APP = 'sua senha de app'
SMTP_SERVIDOR = 'smtp.gmail.com'
SMTP_PORTA = 587 
~~~
Se quiser receber todos os emails em cópia em um email, adicione este email nessa parte seguinte:
~~~python
# E-mail para cópia (CC) - opcional. Deixe como string vazia '' se não quiser usar CC.
CC_EMAIL = ''
~~~

A planilha de presença dos participantes deve ser em formato .xlsx e conter pelo menos as seguintes colunas: `Presença:`, `Nome completo:`, `E-mail:`. 

## Instalação

Além do Python, é necessário instalar as seguintes bibliotecas via pip:

* [**pandas**](https://pandas.pydata.org/) - Manipulação de dados

* [**reportlab**](https://www.reportlab.com/) - Geração de PDFs

Pode ser feito colando essa linha de comando no seu terminal:
```
pip install pandas reportlab
```
