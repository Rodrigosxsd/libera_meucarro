import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def criar_tabela_dados(dados):
    # Escolha as colunas que deseja exibir na tabela (exceto 'EmailTo' e 'EmailCC')
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

def enviar_email(destinatario, assunto, corpo_email, cc=None):
    host = 'smtp.gmail.com'
    porta = 587

    usuario = 'rodrigo.menezes@napista.com.br'
    senha = 'Ro055662400'

    context = ssl.create_default_context()

    server = smtplib.SMTP(host, porta)
    server.starttls(context=context)
    server.login(usuario, senha)

    mensagem = MIMEMultipart()
    mensagem['From'] = usuario
    mensagem['To'] = destinatario
    if cc:
        cc_str = ', '.join(cc)
        mensagem['Bcc'] = cc_str
        # Remove os destinatários e cópias do corpo do e-mail
        corpo_email = corpo_email.replace(destinatario, '').replace(cc_str, '')
    mensagem['Subject'] = assunto
    mensagem.attach(MIMEText(corpo_email, 'html'))
    server.sendmail(usuario, [destinatario] + cc, mensagem.as_string())

    server.quit()

# Carregando os dados da planilha usando o pandas
dados = pd.read_excel(r"E:\\libera_meucarro\dadoscoletados.xlsx")

# Verificando se todas as colunas desejadas estão presentes no DataFrame
colunas_exibir = ['PLACA', 'CHASSI', 'MODELO', 'PATIO', 'RESPONSAVEL', 'CPF', 'RG', 'TELEFONE']
if not set(colunas_exibir).issubset(dados.columns):
    colunas_faltando = set(colunas_exibir) - set(dados.columns)
    raise ValueError(f"Algumas colunas desejadas não estão presentes no DataFrame: {colunas_faltando}")

# Agrupando os dados por pátio
grupos_patio = dados.groupby('PATIO')

# Iterando sobre cada grupo (cada pátio) e enviando os e-mails com os dados dos veículos correspondentes
for nome_patio, grupo_dados in grupos_patio:
    corpo_tabela = criar_tabela_dados(grupo_dados)
    corpo_email = criar_corpo_email(nome_patio, corpo_tabela)

    # Enviando e-mails para os destinatários listados na planilha (coluna 'EmailTo') e também adicionando cópia para 'EmailCC'
    for _, row in grupo_dados.iterrows():
        destinatario_email = row['EmailTo']
        cc_email = row['EmailCC'] if 'EmailCC' in dados.columns else None
        enviar_email(destinatario_email, f"Dados de Responsáveis pela retirada de veículos - {nome_patio} - naPista", corpo_email, cc=[cc_email])

print("E-mails enviados com sucesso!")