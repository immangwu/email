import smtplib
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

username= 'immanual.me@srit.org'#str(input('your username:'));
password='trounxlykdbnzfem'#str(input('your password'));
From =username

##subject
Subject='Certificate for Seminar on Patent filing process & PCT'

wb=xl.load_workbook(r'C:\Users\Immanual mech\Desktop\Events\ip seminar\EMail\Book1.xlsx')
sheet1=wb.get_sheet_by_name('Sheet1')
names=[]
emails=[]
files=[]
for cell in sheet1['A']:
    emails.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

for cell in sheet1['C']:
    files.append(cell.value)

    
for i in range(len(emails)):
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(username,password)
    msg=MIMEMultipart()
    msg['From']=username
    msg['To']=names[i]
    msg['subject']=Subject
    text=''' Hello {},
            Thank you for Participating in Seminar on Patent filing process & PCT,
            Here with i attached the certificate for  {}.

                                                                        Thank You

            Resource Person Contact Details:

            A.K.Balaji,
            Advocate & IP Attorney.
            Contact Mail Id : advocateakb@gmail.com
            Session Video Link: https://www.youtube.com/watch?v=a5YhEjYuuM8

            Nithya.S 
            Advocate & Patent Agent.
            Contact Mail Id : nithya.ipattorney@gmail.com
            Session Video Link: https://www.youtube.com/watch?v=Zn5a9NkAsGI

            For any Queries and Changes in Certificate Kindly Contact

            Mr.R.Immanual
            Mob No:9677817992
            Email Id : immanual.me@srit.org


            

'''.format(names[i],names[i])
    a=str(files[i])
    ##Attachement1
    filename="Certificate_"+a+'.pdf'
    attachement=open(filename,'rb')
    part=MIMEBase('application','octet-stream')
    part.set_payload((attachement).read())
    encoders.encode_base64(part)
    part.add_header('content-Disposition','attachement;filename='+filename)
    msg.attach (part)
##    ####Attachement2
##    filename2=a+'.jpg'
##    attachement2=open(filename2,'rb')
##    part2=MIMEBase('application','octet-stream')
##    part2.set_payload((attachement2).read())
##    encoders.encode_base64(part2)
##    part2.add_header('content-Disposition','attachement;filename='+filename2)
##    msg.attach (part2)

    
    msg.attach(MIMEText(text,'plain'))
    message=msg.as_string()
    server.sendmail(username,emails[i],message)
    print(i)
    print('mail sent to',emails[i])
    server.quit()
print('all emails sent sucessfully')

