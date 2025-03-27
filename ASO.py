import pandas as pd
import smtplib
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurações do e-mail
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
EMAIL_SENDER = "fabrizio.apparicio@decisionbr.com.br"
EMAIL_PASSWORD = "kbrhrwbpmrbzjfvy"

def enviar_email(destinatario, nome):
    """Envia um e-mail de notificação."""
    assunto = "Lembrete: ASO Anual"
    mensagem = f"Olá {nome},\n\nEste é um lembrete de que você precisa realizar seu exame ASO anual.\n\nAtenciosamente,\nEquipe de Saúde Ocupacional"

    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = destinatario
    msg["Subject"] = assunto
    msg.attach(MIMEText(mensagem, "plain"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, destinatario, msg.as_string())
        print(f"✅ E-mail enviado para {nome} ({destinatario})")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para {nome}: {e}")

def verificar_aniversario():
    """Lê a planilha e verifica se algum funcionário faz aniversário de empresa hoje."""
    df = pd.read_excel("funcionarios.xlsx", dtype={"Data_Admissao": str})
    hoje = datetime.date.today()

    for _, row in df.iterrows():
        nome = row["Nome"]
        email = row["Email"]
        data_admissao = pd.to_datetime(row["Data_Admissao"], dayfirst=True).date()  # Agora interpreta DD/MM/YYYY

        if data_admissao.month == hoje.month and data_admissao.day == hoje.day:
            enviar_email(email, nome)

# Executa a verificação
verificar_aniversario()
