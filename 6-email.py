import smtplib
import ssl
import mimetypes
from email.message import EmailMessage
import os

# 1- Dados do E-mail
with open("senha", "r") as f:
    password = f.read().strip()

from_email = "damiaorochaduda@gmail.com"
to_email = "damiaorochaduda@gmail.com"
subject = "Automação Planilha"
body = """
Olá. Segue em anexo a automação da planilha
para a empresa XYZ Automação.

Qualquer dúvida estou à disposição!
"""

# 2- Montando estrutura do e-mail
message = EmailMessage()
message["From"] = from_email
message["To"] = to_email
message["Subject"] = subject  # Corrigido para "Subject" com a primeira letra maiúscula

message.set_content(body)
safe = ssl.create_default_context()

# 3- Adicionando Anexo
anexo = "test.xlsx"
# Verificando se o arquivo existe
if not os.path.isfile(anexo):
    print(f"Arquivo não encontrado: {anexo}")
else:
    # Definindo manualmente o tipo MIME para arquivos .xlsx
    mime_type = "application"
    mime_subtype = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    with open(anexo, "rb") as a:
        message.add_attachment(
            a.read(),
            maintype=mime_type,
            subtype=mime_subtype,
            filename=anexo
        )

# 4- Envio do E-mail
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=safe) as smtp:
    smtp.login(from_email, password)
    smtp.send_message(message)  # Usando send_message ao invés de sendmail
