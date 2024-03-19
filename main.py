from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
import sys
from docxtpl import DocxTemplate
import mammoth
from dotenv import load_dotenv


load_dotenv()


def get_html(template_file, context):
    docx_template = DocxTemplate(template_file)
    docx_template.render(context)
    docx_template.save("generated_document.docx")

    with open("generated_document.docx", "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        return result.value  # The generated HTML


def send_email(sender_email, sender_password, recipient_email, subject, html_content):
    smtp_server = "smtp.office365.com"
    smtp_port = 587

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient_email
    message["Subject"] = subject

    message.attach(MIMEText(html_content, "html"))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(message)
        print("Email sent successfully!")
        server.quit()


if __name__ == "__main__":
    if len(sys.argv) > 3:
        filename = sys.argv[1]
        recipient_email = sys.argv[2]
        title = sys.argv[3]
        filename = "hello.docx"

        context = {"title": title, "name": "Huy"}
        html_content = get_html(filename, context)

        sender_email = os.getenv("SENDER_EMAIL")
        sender_password = os.getenv("SENDER_PASSWORD")
        subject = "Test HTML Email"

        send_email(sender_email, sender_password,
                   recipient_email, subject, html_content)
        print("Done!")
    else:
        print("Error: need filename, recipient_email, title ")
