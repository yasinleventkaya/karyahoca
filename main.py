import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas

excel_file_name = "auto_mail.xlsx"
sheet_name = "Sayfa1"
names_header = "names"
emails_header = "mails"
grades_header = "grades"

sender_address = ''
sender_pass = ''

email_subject = 'Sınav Notu'
email_text = "Selam %s,\n\nSınav notun şu şekildedir: %s\n\nİyi çalışmalar dilerim."


def get_student_grade_pairs():
    df = pandas.read_excel(excel_file_name, sheet_name=sheet_name)
    names = list(filter(lambda x: (type(x) is str) and (x.strip() != ""), df[names_header].tolist()))
    emails = list(filter(lambda x: (type(x) is str) and (x.strip() != ""), df[emails_header].tolist()))
    grades = list(filter(lambda x: (type(x) is float) and (not pandas.isna(x)), df[grades_header].tolist()))
    # print("--")
    # print(emails)
    # print("--")
    # print(grades)
    # print("--")
    # print(names)
    # print("--")

    if (len(names) == len(emails)) and (len(grades) == len(emails)):
        mail_grade_pair = []
        for i in range(len(grades)):
            pair = {
                "name": names[i],
                "email": emails[i],
                "grade": grades[i]
            }
            # print(pair)
            mail_grade_pair.append(pair)
        return mail_grade_pair
    else:
        print("Mail - Not bilgilerinde hata olabilir, dosyayı kontrol et.")


def send_mails(pairs):
    try:
        session = smtplib.SMTP('smtp.gmail.com', 587)  # use gmail with port
        session.starttls()  # enable security
        session.login(sender_address, sender_pass)  # login with mail_id and password
        for pair in pairs:
            receiver_address = pair.get("email")
            message = MIMEMultipart()
            message['From'] = sender_address
            message['To'] = receiver_address
            message['Subject'] = email_subject
            mail_content = email_text % (pair.get("name"), str(pair.get("grade")))
            message.attach(MIMEText(mail_content, 'plain'))
            text = message.as_string()
            session.sendmail(sender_address, receiver_address, text)
            print('Mail Sent to %s' % pair.get("email"))
        session.quit()
    except Exception as e:
        print(e)


def start_auto_mailing():
    pairs = get_student_grade_pairs()
    if pairs:
        send_mails(pairs)


if __name__ == '__main__':
    start_auto_mailing()
