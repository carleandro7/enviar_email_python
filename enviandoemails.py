import csv
import smtplib
from email.message import EmailMessage

# Configurações SMTP Hostinger
SMTP_SERVER = 'smtp.hostinger.com'
SMTP_PORT = 587
EMAIL_REMETENTE = 'suporte@inforpiaui.app.br'
SENHA = 'senha_aqui'  # Use senha segura ou variável de ambiente

# Arquivo CSV de entrada
ARQUIVO_ENTRADA = 'destinatarios.csv'

# Resultados
enviados = []
falharam = []

# Carregar destinatários
def carregar_destinatarios(arquivo):
    with open(arquivo, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        return list(reader)

# Criar mensagem personalizada
def criar_mensagem(nome_completo, email):
    msg = EmailMessage()
    msg['Subject'] = 'Assunto do E-mail'
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = email
    msg.set_content(f'Olá {nome_completo}, este é um e-mail automático enviado via Python.')
    return msg

# Enviar e-mails via SMTP
def enviar_emails(destinatarios):
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=20) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(EMAIL_REMETENTE, SENHA)

            for d in destinatarios:
                nome_completo = d['nome_completo']
                email = d['email']

                try:
                    msg = criar_mensagem(nome_completo, email)
                    server.send_message(msg)
                    enviados.append({'nome_completo': nome_completo, 'email': email})
                    print(f"✅ Enviado para: {nome_completo} <{email}>")

                except smtplib.SMTPRecipientsRefused as e:
                    falharam.append({'nome_completo': nome_completo, 'email': email, 'erro': 'Destinatário recusado'})
                except smtplib.SMTPSenderRefused as e:
                    falharam.append({'nome_completo': nome_completo, 'email': email, 'erro': 'Remetente recusado'})
                except smtplib.SMTPDataError as e:
                    falharam.append({'nome_completo': nome_completo, 'email': email, 'erro': 'Erro ao enviar dados'})
                except Exception as e:
                    falharam.append({'nome_completo': nome_completo, 'email': email, 'erro': str(e)})
                    print(f"❌ Erro inesperado: {nome_completo} <{email}> -> {e}")

    except Exception as e:
        print("❌ Erro geral ao conectar ao SMTP:", e)

# Salvar resultado em CSV
def salvar_csv(lista, nome_arquivo, campos):
    with open(nome_arquivo, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        writer.writerows(lista)

# Execução
destinatarios = carregar_destinatarios(ARQUIVO_ENTRADA)
enviar_emails(destinatarios)

salvar_csv(enviados, 'emails_enviados.csv', ['nome_completo', 'email'])
salvar_csv(falharam, 'emails_falharam.csv', ['nome_completo', 'email', 'erro'])

print("\n📁 Relatórios salvos: 'emails_enviados.csv' e 'emails_falharam.csv'")
