import pandas as pd
import smtplib
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

# Configurações do e-mail
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
EMAIL_SENDER = "fabrizio.apparicio@decisionbr.com.br"
EMAIL_PASSWORD = "kbrhrwbpmrbzjfvy"

def enviar_email(destinatario, nome):
    """Envia um e-mail de notificação com assinatura HTML."""
    assunto = "Lembrete: ASO Anual"

    # Assinatura HTML com tabela e imagem
    assinatura_html = """
    <table>
        <tbody><tr>
            <td>
            <a href="https://www.decisionbr.com/">
                <img src="https://drive.google.com/uc?id=1WwhrTiAc83EOzGETQd-NNMr9VGjS5wyx" alt="image">
            </a>
            </td>
            <td style="margin: 20px; font-size: 12px; color: grey; font-family: Verdana, sans-serif;">
            <p style="margin-bottom: -20px;">________________________________________________________________</p>
                <p><p style="margin-bottom: 10px; color: #4D2DD0; font-size: 14px;">Fabrizio | Desenvolvimento Humano e Organizacional</p>
                Decision | Unidade São Paulo<br> 
                Telefone: (19) 3252-2838 | (11) 98280-2722<br>
                <a style="color: blue;" href="https://www.decisionbr.com/">www.decisionbr.com.br</a>
            </p>
            </td>
        </tr></tbody>
    </table>

    """

    mensagem = f"Olá {nome},<br><br>Este é um lembrete de que você precisa realizar seu exame ASO anual.<br><br>{assinatura_html}"

    # Criando a mensagem com MIME
    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = destinatario
    msg["Subject"] = assunto
    msg.attach(MIMEText(mensagem, "html"))

    # # Adicionando imagem à assinatura
    # with open("logo_empresa.png", "rb") as img_file:
    #     img = MIMEImage(img_file.read())
    #     img.add_header("Content-ID", "<logo>")
    #     msg.attach(img)

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

    print(f"📆 Hoje é: {hoje}")

    for _, row in df.iterrows():
        nome = row["Nome"]
        email = row["Email"]
        data_admissao = pd.to_datetime(row["Data_Admissao"], dayfirst=False).date()  # Agora interpreta DD/MM/YYYY

        if data_admissao.month == hoje.month and data_admissao.day == hoje.day:
            print(f"Verificando: {nome} - {email} - Admissão em {data_admissao}")
            enviar_email(email, nome)

# Executa a verificação
verificar_aniversario()
