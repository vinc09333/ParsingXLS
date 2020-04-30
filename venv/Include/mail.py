from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import loggs

    #Отправка передаваемого в msg сообщения на почту
def err_mail_msg(msg, param):
    smtp_server = "192.168.100.238"
    smtp_port = "25"
    to = ['am.fesenko@agrohold.ru']
    frm = "xls.scrip@agrohold.ru"

    if msg:
        message = MIMEMultipart("alternative")
        if param == '-s':
            message["Subject"] = f'Отчет об ошибках функции get_svin на момент {datetime.today().strftime(" %Y/%m/%d; %H:%M;")}'
            
        if param == '-k':
            message["Subject"] = f'Отчет об ошибках функции get_korm на момент {datetime.today().strftime(" %Y/%m/%d; %H:%M;")}'

        message["From"] = frm
        message["To"] = ', '.join(to)
        html = \
            "<html>" \
            "<body>" \
            "<p>" \
            "<strong>" \
            + msg + \
            "</strong></p>" \
            "</body>" \
            "</html>"
        # Сделать их текстовыми\html объектами MIMEText
        message.attach(MIMEText(msg, "plain"))
        message.attach(MIMEText(html, "html"))
        mailObj = smtplib.SMTP(smtp_server, smtp_port)
        mailObj.sendmail(frm, to, message.as_string())
        mailObj.quit()