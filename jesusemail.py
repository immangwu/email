import smtplib
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

username= 'immangwu@gmail.com'#str(input('your username:'));
password='JESUS143su'#str(input('your password'));
From =username

##subject
Subject='Jesus Loves Me This I Know'

wb=xl.load_workbook(r'C:\Users\Immanual mech\Desktop\emailpy\Book1.xlsx')
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
    text=''' Hello{}, I want to tell Jesus Loves {}'''.format(names[i],names[i])
    a=str(files[i])
    ##Attachement1
    filename=a+'.xlsx'
    attachement=open(filename,'rb')
    part=MIMEBase('application','octet-stream')
    part.set_payload((attachement).read())
    encoders.encode_base64(part)
    part.add_header('content-Disposition','attachement;filename='+filename)
    msg.attach (part)
    ####Attachement2
    filename2=a+'.jpg'
    attachement2=open(filename2,'rb')
    part2=MIMEBase('application','octet-stream')
    part2.set_payload((attachement2).read())
    encoders.encode_base64(part2)
    part2.add_header('content-Disposition','attachement;filename='+filename2)
    msg.attach (part2)

    
    msg.attach(MIMEText(text,'plain'))
    message=msg.as_string()
    server.sendmail(username,emails[i],message)
    print(i)
    print('mail sent to',emails[i])
    server.quit()
print('all emails sent sucessfully')

