import smtplib
from email.message import EmailMessage

def send_test_email():
    print("Iniciando o script de envio de e-mail...")

    msg = EmailMessage()
    msg['Subject'] = 'Teste de Envio de E-mail'
    msg['From'] = 'gustavoheenri02@gmail.com'
    msg['To'] = 'gustavoheenri02@gmail.com'
    msg.set_content('Olá, este é um e-mail de teste para verificar o funcionamento do envio de e-mail.')

    try:
        print("Conectando ao servidor SMTP...")
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            print("Realizando login...")
            smtp.login('gustavoheenri02@gmail.com', 'ayvo ibvf hwte hkdq')  # Substitua pela sua senha de app
            print("Enviando e-mail...")
            smtp.send_message(msg)
        print("E-mail enviado com sucesso!")
    except smtplib.SMTPAuthenticationError as e:
        print("Erro de autenticação:", e)
    except Exception as e:
        print("Erro ao enviar e-mail:", e)

if __name__ == "_main_":
    send_test_email()