import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def enviar_email_html(remetente, destinatarios, assunto, corpo_email, cc=None):
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente
    mensagem['To'] = ', '.join(destinatarios)
    if cc:
        cc_str = ', '.join(cc)
        mensagem['Cc'] = cc_str
    mensagem['Subject'] = assunto
    mensagem.attach(MIMEText(corpo_email, 'html'))

    # Configurar o servidor SMTP e enviar o e-mail
    host = 'smtp.gmail.com'  # Substitua pelo servidor SMTP do seu e-mail corporativo
    porta = 587
    usuario = 'xxxxxxxxxx'  # Substitua pelo seu endereço de e-mail corporativo
    senha = 'xxxxxxxxxxxxx'  # Substitua pela senha do seu e-mail corporativo

    context = ssl.create_default_context()

    with smtplib.SMTP(host, porta) as server:
        server.starttls(context=context)
        server.login(usuario, senha)
        destinatarios = destinatarios if destinatarios else []  # Caso não exista destinatários, criar uma lista vazia
        cc = cc if cc else []  # Caso não exista cc, criar uma lista vazia
        server.sendmail(remetente, destinatarios + cc, mensagem.as_string())

# Carregar os dados da planilha para um DataFrame do pandas
df = pd.read_excel('E:\\libera_meucarro\dadoscoletados.xlsx')

# Endereço de e-mail do grupo que será utilizado como remetente
email_remetente_grupo = 'grupo_email@seuemailcorporativo.com'

# Corpo do e-mail em HTML com espaços reservados para as informações do veículo
corpo_email_html = """\
<html>
<body>
    <p>Olá, [LOJISTA]!</p>
    <p>O veículo adquirido de placa [PLACA] e modelo [MODELO], encontra-se disponível para AGENDAMENTO E RETIRADA. Diante disso, pedimos que realize contato com o pátio para agendar a retirada do veículo.</p>
    <p>O veículo encontra-se no PÁTIO:</p>
    <p>Endereço: [ENDERECO_PATIO]</p>
    <p>Telefone: [TELEFONE_PATIO]</p>
    <p>Responsável: [RESPONSAVEL_PATIO]</p>
    <p>Importante certificar, durante o contato para agendamento, se o veículo necessita de algum item para sua retirada (óleo, combustível, cabo para auxiliar bateria, bomba para encher pneu ou plataforma/guincho).</p>
    <p>O DUT será enviado para o endereço sinalizado anteriormente, e não estará disponível na retirada do veículo.</p>
    <p>Mais uma vez agradecemos por usar a nossa plataforma.</p>
    <p>Com o NaPista é fácil de comprar!</p>
    <p>Atendimento naPista</p>
    <p>E-mail: bancos@napista.com.br</p>
    <p>Site: <a href="https://www.napista.com.br/">https://www.napista.com.br/</a></p>
    <p>Instagram: @napistapp</p>
</body>
</html>
"""

# Definir o limite máximo para o tamanho do assunto do e-mail (em caracteres)
limite_tamanho_assunto = 500

# Iterar sobre os dados da planilha e enviar os e-mails para cada lojista
for lojista, lojista_data in df.groupby('LOJISTA'):
    for index, row in lojista_data.iterrows():
        # Preencher as informações do veículo no corpo do e-mail com o tratamento strip()
        corpo_email_preenchido = corpo_email_html.replace('[LOJISTA]', lojista.strip())
        corpo_email_preenchido = corpo_email_preenchido.replace('[PLACA]', row['PLACA'].strip())
        corpo_email_preenchido = corpo_email_preenchido.replace('[MODELO]', row['MODELO'].strip())
        corpo_email_preenchido = corpo_email_preenchido.replace('[ENDERECO_PATIO]', row['ENDERECO_PATIO'].strip())
        corpo_email_preenchido = corpo_email_preenchido.replace('[TELEFONE_PATIO]', str(row['TELEFONE_PATIO']).strip())
        corpo_email_preenchido = corpo_email_preenchido.replace('[RESPONSAVEL_PATIO]', row['LOJISTA'].strip()) # Substituir 'RESPONSAVEL_PATIO' por 'LOJISTA'

        # Montar o assunto do e-mail
        assunto = f"Orientações para Retirada de Veículo - Placa {row['PLACA'].strip()} - NaPista"

        # Limitar o tamanho do assunto
        if len(assunto) > limite_tamanho_assunto:
            assunto = assunto[:limite_tamanho_assunto]

        # Separar os e-mails em cópia (CC) se houver mais de um e-mail
        email_cc_list = row['Emailscard'].split(',') if pd.notnull(row['Emailscard']) else []

        # Separar os e-mails de destinatários se houver mais de um e-mail
        destinatarios_list = row['EMAIL_LOJISTA'].split(',') if pd.notnull(row['EMAIL_LOJISTA']) else []

        # Enviar e-mail para o lojista com as informações do veículo
        enviar_email_html(email_remetente_grupo, destinatarios_list, assunto, corpo_email_preenchido, cc=email_cc_list)

print("Emails enviados para os Lojistas!")
