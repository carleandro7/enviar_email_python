import imaplib
import email
from email.header import decode_header
import re
import csv

# Configurações do e-mail (Hostinger)
IMAP_SERVER = 'imap.hostinger.com'
EMAIL = 'suporte@inforpiaui.app.br'
SENHA = 'senha_aqui'

# Conectar ao servidor IMAP
def conectar_email():
    print("🔗 Conectando ao servidor de e-mail...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, SENHA)
    return mail

# Procurar e-mails com falhas
def buscar_bounces(mail):
    print("📥 Buscando e-mails com falhas de entrega...")
    mail.select("inbox")
    # Procurar por assuntos típicos de erro
    status, messages = mail.search(None,
        '(OR SUBJECT "Mail Delivery Failed" SUBJECT "Undelivered Mail Returned to Sender")'
    )

    bounces = []
    for num in messages[0].split():
        res, data = mail.fetch(num, "(RFC822)")
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                assunto, encoding = decode_header(msg["Subject"])[0]
                if isinstance(assunto, bytes):
                    assunto = assunto.decode(encoding or "utf-8")

                corpo = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            corpo = part.get_payload(decode=True).decode(errors="ignore")
                            break
                else:
                    corpo = msg.get_payload(decode=True).decode(errors="ignore")

                # Tentar extrair e-mail rejeitado
                emails_rejeitados = re.findall(r'[\w\.-]+@[\w\.-]+', corpo)
                bounces.append({
                    'assunto': assunto,
                    'email_rejeitado': emails_rejeitados,
                })

    return bounces

# Execução
# Salvar em CSV
def salvar_em_csv(lista, nome_arquivo='emails_nao_entregues.csv'):
    with open(nome_arquivo, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['assunto', 'email_rejeitado'])
        writer.writeheader()
        writer.writerows(lista)

# Execução
mail = conectar_email()
falhas = buscar_bounces(mail)
mail.logout()

print("\n📩 E-mails com erro de entrega encontrados:")
for f in falhas:
    print(f"- Assunto: {f['assunto']}")
    print(f"  → E-mail rejeitado: {f['email_rejeitado']}")

salvar_em_csv(falhas)
print("\n📁 Relatório salvo: emails_nao_entregues.csv")
