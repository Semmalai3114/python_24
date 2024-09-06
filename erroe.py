#Modules
import os
from openpyxl import load_workbook
import glob
import smtplib
from email.message import EmailMessage
import email.utils
from tabulate import tabulate
#latest file access

source=r'\\localhost\cobian\*'
list_of_files = glob.glob(source) 
latest_file = max(list_of_files, key=os.path.getctime)
file_name=os.path.basename(latest_file)
k=source+str(file_name)
k=k.replace("\\","\\")                                                                               
k=k.replace("*","")
f=open(k,'r',encoding='utf-16-le')
a=f.readlines()
non=[]
non2=[]
mydata=[]

def Automation(words,serial,b,c=0,c1=0,h='',r='',backup='',status='',d1='',d2='' ):
#counting number of lines in file 
    for j in a:
        c1+=1
#taking first and last lines
    for i in range(0,c1):
        c+=1
        if((words in a[i]) and ("ERR" not in a[i])):
            b.insert(c1,c-1)
            h=a[int(b[0])]
            r=a[int(b[-1])]
            continue
    
#split string and insert into list
    d=r.split()
    if("Errors:" in d): 
        sindex=d.index("Errors:")
        n=int((d[sindex+1][0]))
    if("Errors:" in d):
        eindex=d.index("size:")
        n2=float(d[eindex+1])
        n3=round(n2)
        n1=str((d[eindex+1]))+" "+str((d[eindex+2]))
        if(n3>0):
            status="Failed"
            non.append('---->Cobain backup has been compelted!')
        elif(n3==0):
            non.append('---->Cobain backup has been Failed!')
            status="Done"
        backup=n1
#For identify Backup Name not in folder 
    d2=h.split()
    d1=r.split()
    if(len(d1)==0 ):
        start_date="not in run"
        start_time="not in run"
        End_time="not in run"
    else:
        start_date=d1[0]
        start_time=d2[1]
        End_time=d1[1]
#Adding detials in Excel sheet
    existing_file = r'\\localhost\cobian_excel\Book1.xlsx' # 'D:/cobian_excel/Book1.xlsx'
    new_data = [[start_date,serial,start_time,End_time,words,backup,status," ","Python-Automation"]]
    wb = load_workbook(existing_file)
    ws = wb.active
    for row in new_data:
        ws.append(row)
        wb.save(existing_file)
#Permission For Update
try:
    w=["SP_APP","CE_Working","EBOOK_EPUB-Outgoing","EBOOK","IT Management","SP QC_PDF","CE_Track","SP Tracking","Devtool","BR"]
    for i in range(0,len(w)):
        Automation(w[i],i+1,b=[])
        if(i==len(w)-1):
#Mail sending 
            EMAIL_ADDRESS = 'alerts@cps-india.com'
            EMAIL_PASSWORD = 'chennai@123'
            EMAIL_TO="ithelpdesk.chn@cps-india.com"
            msg = EmailMessage()
            msg['Subject'] = 'Cobain backup'
            msg['From'] = EMAIL_ADDRESS
            msg['To'] = EMAIL_TO
            msg['Date'] = email.utils.formatdate(localtime=True)
            msg['Message-ID'] = email.utils.make_msgid()
            w1=["SP_APP","CE_Working","EBOOK_EPUB-Outgoing","EBOOK","IT Management","SP QC_PDF","CE_Track","SP Tracking","Devtool","BR"]
            for l in range(len(w1)-1):
                non2.append(str(w1[l]))
                non5=[non2[l],non[l]]
                mydata.append(non5)
            head = ["Backup Name", "Report"]
            msg.set_content(tabulate(mydata, headers=head))
            Email=smtplib.SMTP("mail.cps-india.com", 587)
            Email.starttls()
            Email.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            Email.send_message(msg)
            

except PermissionError as er:
    content= er
    print(er)
    EMAIL_ADDRESS = 'alerts@cps-india.com'
    EMAIL_PASSWORD = 'chennai@123'
    EMAIL_TO="ithelpdesk.chn@cps-india.com"
    msg = EmailMessage()
    msg['Subject'] = 'PermissionError'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = EMAIL_TO
    msg['Date'] = email.utils.formatdate(localtime=True)
    msg['Message-ID'] = email.utils.make_msgid()
    msg.set_content(str(content )+str('\n kindly close excel file'))
    Email=smtplib.SMTP("mail.cps-india.com", 587)
    Email.starttls()
    Email.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    Email.send_message(msg)
    print("success")
