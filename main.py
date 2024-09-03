from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_mail(datos):
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = datos[0].value
    msg['Subject'] = datos[6].value
    body = f"Dear {datos[1].value} {datos[2].value},\n\nWe want to thank you for being part of our community and allowing us to keep you informed about the latest news and updates from our brand.\n\nTo ensure that you continue receiving our emails and stay updated on all our news, promotions, and special events, we want to confirm that we have your contact information up-to-date. Below, we share the details we have on file:\n\n- Full Name: {datos[1].value} {datos[2].value}\n- Contact Phone: {datos[3].value}\n- Mailing Address: {datos[4].value}\n- Email Address: {datos[0].value}\n- Subscription Date: {datos[5].value}\n\nIf any of this information has changed or if you wish to update your profile, please feel free to contact us by replying to this email or by accessing your account on our website.\n\nWe are here to help and ensure that you receive the best possible experience with this example brand.\n\nThank you for your continued support.\n\nBest regards,\n\n"


    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Seguridad TLS
        server.login(email_address, email_password)
        server.send_message(msg)
        server.quit()
        print("the email was succesfully send it")
    except Exception as e:
        print(f"Error :{e}")



#sign in
smtp_server = 'smtp.gmail.com'
smtp_port = 587
email_address = 'your email'
email_password = 'the password of your email or the email password for this app'


# Cargar el libro de trabajo
wb = load_workbook('datos.xlsx')
sheet = wb.active
for i in sheet:
    send_mail(i)
    
