import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from config import SMTP_HOST, SMTP_PORT, EMAIL_REMETENTE, EMAIL_SENHA

def enviar_email(destino, assunto, corpo, anexos=None):
    """
    Envia e-mail via SMTP com suporte a múltiplos anexos.
    Configurações lidas do config.py.
    """
    if not EMAIL_REMETENTE or not EMAIL_SENHA:
        raise ValueError("E-mail remetente ou senha não configurados no config.py")

    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = destino
    msg['Subject'] = assunto

    # Assinatura Com System
    rodape = "\n\n--\nCom System\nAgência de Atendimento Digital"
    msg.attach(MIMEText(corpo + rodape, 'plain'))

    if anexos:
        for caminho in anexos:
            if os.path.exists(caminho):
                with open(caminho, "rb") as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(caminho))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho)}"'
                    msg.attach(part)

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(EMAIL_REMETENTE, EMAIL_SENHA)
        server.send_message(msg)

def enviar_email_relatorio(destino, assunto, corpo, anexos=None):
    """Alias para enviar_email, usado em algumas partes do sistema."""
    return enviar_email(destino, assunto, corpo, anexos)
