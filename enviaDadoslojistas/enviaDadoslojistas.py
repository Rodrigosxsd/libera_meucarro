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
    usuario = 'rodrigo.menezes@napista.com.br'  # Substitua pelo seu endereço de e-mail corporativo
    senha = 'Ro055662400'  # Substitua pela senha do seu e-mail corporativo

    context = ssl.create_default_context()

    with smtplib.SMTP(host, porta) as server:
        server.starttls(context=context)
        server.login(usuario, senha)
        server.sendmail(remetente, destinatarios + cc if cc else destinatarios, mensagem.as_string())

# Carregar os dados da planilha para um DataFrame do pandas
df = pd.read_excel(r"E:\\libera_meucarro\dadoscoletados.xlsx")

# Endereço de e-mail do grupo que será utilizado como remetente
email_remetente_grupo = 'bancos@napista.com.br'

# Corpo do e-mail em HTML com espaços reservados para as informações do veículo
corpo_email_html = """\
<html>
<body>
    <p>Olá, [LOJISTA]!</p>
    <p>O veículo adquirido de placa [PLACA] e modelo [MODELO], encontra-se disponível para AGENDAMENTO E RETIRADA. Diante disso, pedimos que realize contato com o pátio para agendar a retirada do veículo.</p>
    <p>O veículo encontra-se no PÁTIO: [PATIO]</p>
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
    <p> <a href="https://wa.me/551131641007"> Whatsapp: (11) 3164-1007</a></p>
</body>
</html>
"""
limite_tamanho_assunto = 500

# Iterar sobre os dados da planilha e enviar os e-mails para cada lojista
for lojista, lojista_data in df.groupby('LOJISTA'):
    destinatarios = lojista_data['EMAIL_LOJISTA'].explode().tolist()  # Lista de destinatários
    cc = lojista_data['EmailCC'].explode().tolist() if 'EmailCC' in df.columns else []

    # Iterar sobre os carros do lojista e enviar e-mails individualmente
    for index, row in lojista_data.iterrows():
        placa = row['PLACA']
        modelo = row['MODELO']
        patio = row['PATIO']
        endereco = row['ENDERECO_PATIO']
        telefone_patio = str(row['TELEFONE_PATIO'])  # Converter para string
        responsavel_patio = row['RESPONSAVEL_PATIO']

        assunto = f"Orientações para Retirada de Veículo - Placa {placa} - NaPista"
        if len(assunto) > limite_tamanho_assunto:
            assunto = assunto[:limite_tamanho_assunto]


        # Substituir os espaços reservados no corpo do e-mail pelas informações do veículo
        corpo_email_html_preenchido = corpo_email_html.replace('[LOJISTA]', lojista)
        corpo_email_html_preenchido = corpo_email_html_preenchido.replace('[PATIO]', patio)
        corpo_email_html_preenchido = corpo_email_html_preenchido.replace('[PLACA]', placa)
        corpo_email_html_preenchido = corpo_email_html_preenchido.replace('[MODELO]', modelo)
        corpo_email_html_preenchido = corpo_email_html_preenchido.replace('[ENDERECO_PATIO]', endereco)
        corpo_email_html_preenchido = corpo_email_html_preenchido.replace('[TELEFONE_PATIO]', telefone_patio)
        corpo_email_html_preenchido = corpo_email_html_preenchido.replace('[RESPONSAVEL_PATIO]', responsavel_patio)

        # Enviar o e-mail para o lojista com as informações do veículo
        enviar_email_html(email_remetente_grupo, destinatarios, assunto, corpo_email_html_preenchido, cc=cc)
print('Orientações para Retirada de Veículo Enviados Para Os lojistas!')        
