import datetime
import email
import imaplib
import config
import xlsxwriter


def login():
    mail = imaplib.IMAP4_SSL(config.imap_server, config.imap_port)
    mail.login(config.username, config.password)
    read_mail(mail)


def read_mail(mail):
    mail.select('INBOX/GSD')
    """ if you are not sure about the folder kindly use the line to check the folders
      print(mail.list()) """
    today = datetime.datetime.today()
    cutoff = today - datetime.timedelta(days=2)
    r, d = mail.search(None, "ALL", 'SINCE', cutoff.strftime('%d-%b-%Y'))
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('daily_tracker.xlsx')
    worksheet = workbook.add_worksheet()
    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0
    for num in d[0].split():
        result, data = mail.fetch(num, "(RFC822)")
        raw_email_string = data[0][1].decode('utf-8')
        msg = email.message_from_string(raw_email_string)
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                mail_content = str(part.get_payload(decode=True))
                content_list = mail_content.split("\\r\\n")
                worksheet.write(row, col, msg['Date'])
                for i in range(len(content_list)):
                    worksheet.write(row, col + 1, content_list[i])
                    col += 1
                row += 1
                col = 0
    workbook.close()


if __name__ == "__main__":
    login()
