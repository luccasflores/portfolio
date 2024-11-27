import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import threading
import pandas as pd

# SMTP server settings
smtp_servers = [
    ('smtp.provedor.com', 587, 'email@email.com', 'Senha'),
    ('smtp.provedor.com', 587, 'email@email.com', 'Senha'),
    ('smtp.provedor.com', 587, 'email@email.com', 'Senha')
]

# Email body with HTML content
corpo_email = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Reduza Custos com RPA</title>
</head>
<body>
  <table align="center" width="600" cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif; border-collapse: collapse; background-color: #f9f9f9; border: 1px solid #ddd;">
    <tr> 
      <td style="padding: 20px; text-align: center; background-color: #6c63ff; color: #fff;">
        <h1 style="margin: 0; font-size: 24px;">Um robô pode ser a solução!</h1>
      </td>
    </tr>
    <tr>
      <td style="padding: 20px; text-align: left;">
        <p>Olá,</p>
        <p>Meu nome é <strong>Luccas Flores</strong>, sou especialista em Automação de Processos Robóticos (RPA), análise de dados e inteligência artificial. Estou entrando em contato para apresentar como soluções de RPA podem transformar sua empresa, otimizando tarefas, reduzindo custos e liberando recursos humanos para atividades estratégicas.</p>
        <p>Com a automação, sua empresa pode:</p>
        <ul>
          <li><strong>Reduzir custos operacionais</strong> automatizando tarefas repetitivas.</li>
          <li><strong>Minimizar erros humanos</strong>, aumentando a confiabilidade nos processos.</li>
          <li><strong>Aumentar a eficiência</strong> das equipes e o tempo dedicado a decisões importantes.</li>
        </ul>
        <p>Já ajudamos empresas a reduzir até <strong>95%</strong> em custos, otimizando a carga de trabalho e melhorando a precisão dos processos internos.</p>
        <p>Gostaria de agendar uma breve conversa para entender seus desafios e demonstrar como minha experiência pode beneficiar sua empresa. Entre em contato diretamente pelo WhatsApp clicando no botão abaixo:</p>
        <p style="text-align: center;">
          <a href="https://wa.me/5548999540943" style="display: inline-block; padding: 10px 20px; background-color: #6c63ff; color: #fff; text-decoration: none; font-size: 16px; border-radius: 5px;">Falar no WhatsApp</a>
        </p>
        <p>Estou à disposição para transformar seus processos em oportunidades de crescimento!</p>
        <p>Atenciosamente,</p>
        <p><img src="cid:assinatura" alt="Assinatura" style="width: 600px; height: 300px;"></p>
      </td>
    </tr>
    <tr>
      <td style="text-align: center; padding: 10px; font-size: 12px; color: #aaa;">
        <p>Este é um email automatizado. Para deixar de receber mensagens, clique <a href="#" style="color: #6c63ff;">aqui</a>.</p>
      </td>
    </tr>
  </table>
</body>
</html>
"""




def enviar_email(destinatario, nome, corpo_email, smtp_server, smtp_port, smtp_user, smtp_password):
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)

        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = destinatario
        msg['Subject'] = f'A/C de {nome}'

        # Adicionando o corpo do email
        msg.attach(MIMEText(corpo_email.format(nome=nome), 'html'))

        # Anexando a imagem da logo
        with open("logo.png", 'rb') as img_file:
            img = MIMEImage(img_file.read())
            img.add_header('Content-ID', '<logo>')
            img.add_header('Content-Disposition', 'inline', filename='logo.png')
            msg.attach(img)

        # with open("imagem.png", 'rb') as img_file:
        #     img = MIMEImage(img_file.read())
        #     img.add_header('Content-ID', '<imagem>')
        #     img.add_header('Content-Disposition', 'inline', filename='imagem.png')
        #     msg.attach(img)

        with open("assinatura.png", 'rb') as img_file:
            img = MIMEImage(img_file.read())
            img.add_header('Content-ID', '<assinatura>')
            img.add_header('Content-Disposition', 'inline', filename='assinatura.png')
            msg.attach(img)
        #
        # with open("Apresentacao.pdf", 'rb') as pdf_file:
        #     pdf_attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
        #     pdf_attachment.add_header('Content-Disposition', 'attachment', filename="Apresentacao.pdf")
        #     msg.attach(pdf_attachment)

        server.sendmail(smtp_user, destinatario, msg.as_string())
        server.quit()
        print(f"Email enviado para {destinatario}")
        return True
    except Exception as e:
        print(f'Erro ao enviar email para {destinatario}: {str(e)}')
        return False


def send_emails():
    df = pd.read_excel('lead.xlsx')  # Supondo que 'lead.xlsx' tenha colunas 'Email' e 'Nome'
    num_emails = len(df)
    batch_size = 200

    for i in range(0, num_emails, batch_size):
        # Determine o servidor e login a serem usados
        smtp_server, smtp_port, smtp_user, smtp_password = smtp_servers[(i // batch_size) % len(smtp_servers)]

        for index in range(i, min(i + batch_size, num_emails)):
            row = df.iloc[index]
            destinatario = row['Email']
            nome = row['Nome']
            enviar_email(destinatario, nome, corpo_email, smtp_server, smtp_port, smtp_user, smtp_password)


def start_sending():
    threading.Thread(target=send_emails).start()


# Start sending emails
start_sending()
