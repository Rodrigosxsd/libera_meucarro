import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def criar_tabela_dados(dados):
    colunas_exibir = ['PLACA', 'CHASSI', 'MODELO', 'PATIO', 'RESPONSAVEL', 'CPF', 'RG', 'TELEFONE']
    tabela_formatada = dados[colunas_exibir].to_html(index=False, border=1, classes='dataframe')
    return tabela_formatada

def criar_corpo_email(nome_patio, corpo_tabela):
    corpo_email = f"<html><body><p style='font-size:16px'>Bom dia,</p>"
    corpo_email += f"<p style='font-size:16px'>Seguem os dados do responsável pela retirada do veículo vendido no pátio {nome_patio}.</p>"
    corpo_email += "<p style='font-size:16px'>O lojista será orientado a entrar em contato para realizar agendamento.</p>"
    corpo_email += corpo_tabela
    corpo_email += "</body></html>"
    return corpo_email

def enviar_email(remetente, destinatario, assunto, corpo_email, cc=None):
    host = 'smtp.gmail.com'
    porta = 587

    usuario = 'xxxx'  # Substitua pelo seu endereço de email
    senha = 'xxxxxx'  # Substitua pela sua senha

    context = ssl.create_default_context()

    server = smtplib.SMTP(host, porta)
    server.starttls(context=context)
    server.login(usuario, senha)

    limite_tamanho_assunto = 500
    assunto = assunto[:limite_tamanho_assunto]

    mensagem = MIMEMultipart()
    mensagem['From'] = remetente  # Define o remetente como o endereço de email vinculado ao grupo
    mensagem['To'] = destinatario
    if cc:
        cc_str = ', '.join(cc)
        mensagem['Cc'] = cc_str
        # Remove os destinatários e cópias do corpo do e-mail
        corpo_email = corpo_email.replace(destinatario, '').replace(cc_str, '')
    mensagem['Subject'] = assunto
    mensagem.attach(MIMEText(corpo_email, 'html'))
    server.sendmail(usuario, [destinatario] + cc.split(',') if cc else [], mensagem.as_string())

    server.quit()

def limpar_espacos_brancos(emails):
    # Função para remover espaços extras de cada e-mail
    emails_limpos = [email.strip() for email in emails.split(',')]
    return ','.join(emails_limpos)

# Carregando os dados da planilha usando o pandas
dados = pd.read_excel(r"E:\\libera_meucarro\dadoscoletados.xlsx")

# Verificando se todas as colunas desejadas estão presentes no DataFrame
colunas_exibir = ['PLACA', 'CHASSI', 'MODELO', 'PATIO', 'RESPONSAVEL', 'CPF', 'RG', 'TELEFONE']
if not set(colunas_exibir).issubset(dados.columns):
    colunas_faltando = set(colunas_exibir) - set(dados.columns)
    raise ValueError(f"Algumas colunas desejadas não estão presentes no DataFrame: {colunas_faltando}")

# Agrupando os dados por pátio
grupos_patio = dados.groupby('PATIO')

# Endereço de email do grupo remetente
email_remetente_grupo = 'atendimento@napista.com.br'

# Iterando sobre cada grupo (cada pátio) e enviando os e-mails com os dados dos veículos correspondentes
for nome_patio, grupo_dados in grupos_patio:
    corpo_tabela = criar_tabela_dados(grupo_dados)
    corpo_email = criar_corpo_email(nome_patio, corpo_tabela)

    # Coletando todos os e-mails da coluna 'EmailTo' em uma única string
    destinatarios_to = ','.join(grupo_dados['EmailTo'].apply(limpar_espacos_brancos))

    # Verificando se há e-mails CC
    if 'EmailCC' in dados.columns:
        destinatarios_cc = ','.join(grupo_dados['EmailCC'].apply(limpar_espacos_brancos))
    else:
        destinatarios_cc = ''

    # Enviando o e-mail para todos os destinatários To e CC de uma única vez
    enviar_email(email_remetente_grupo, destinatarios_to, f"Dados de Responsáveis pela retirada de veículos - {nome_patio} - naPista", corpo_email, cc=destinatarios_cc)

print("E-mails enviados com sucesso!")
