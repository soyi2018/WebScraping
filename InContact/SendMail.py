import smtplib
from os.path import basename
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate, formataddr

def send_mail(send_from, send_to, send_cc, send_bcc, subject, html, files=None,
              server="externalmailmimecast.ashna.udcsms.com"):
    
    msgRoot = MIMEMultipart('related')
    msgRoot['From'] = send_from
    msgRoot['To'] = send_to
    msgRoot['Cc'] = send_cc
    msgRoot['Bcc'] = send_bcc
    msgRoot['Date'] = formatdate(localtime=True)
    msgRoot['Subject'] = subject
    rcpt = send_to.split(',')+send_cc.split(',')+send_bcc.split(',')
    
    msgAlt = MIMEMultipart('alternative')
    msgRoot.attach(msgAlt)
    
    if '<html>' in html:
        msgAlt.attach(MIMEText(html,'html'))      
        imgpath = r'images'
        imgs = ['Ashfield']
        for img in imgs:
            with open(imgpath+'\\'+img+'.png','rb') as fp:
                msgImage = MIMEImage(fp.read())
            msgImage.add_header('Content-ID','<'+img+'>')
            msgRoot.attach(msgImage)
                
    else:
        msgAlt.attach(MIMEText(html,'plain'))
    
    for f in files or []:
        with open(f, "rb") as att:
            part = MIMEApplication(att.read(),Name=basename(f))
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msgRoot.attach(part)

    smtp = smtplib.SMTP(server)
    smtp.ehlo()
    smtp.starttls()
    smtp.sendmail(send_from, rcpt, msgRoot.as_string())
    smtp.close()