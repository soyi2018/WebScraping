import pandas as pd
import numpy as np
import urllib
import glob
import pyodbc
import sqlalchemy
from datetime import datetime
import logging
import time
import os
import smtplib
from os.path import basename
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate, formataddr


inbound = r'G:\ACMS Nurse Partner\Amgen ENPP\Contact_Center_IC_Reports\Inbound\Daily'
outbound = r'G:\ACMS Nurse Partner\Amgen ENPP\Contact_Center_IC_Reports\Outbound'
archive = r'G:\ACMS Nurse Partner\Amgen ENPP\Contact_Center_IC_Reports\Archive\Daily'
log = r'G:\ACMS Nurse Partner\Amgen ENPP\Contact_Center_IC_Reports\Log'
t = datetime.today().strftime('%Y%m%d%H%M')
logname = log + r'\Load_Amgen_CC_IC_Report_Daily_' + t + '.log'
logging.basicConfig(filename=logname,
                    filemode='a',
                    format='%(asctime)s, %(msecs)d %(name)s %(levelname)s %(message)s',
                    level=logging.ERROR)

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=usivy-sql02;'
                      'Database=BI_TotalCare_ENPP;'
                      'Trusted_Connection=yes;')
params = urllib.parse.quote_plus("DRIVER={SQL Server Native Client 11.0};SERVER=usivy-sql02;DATABASE=BI_TotalCare_ENPP;Trusted_Connection=yes")
engine = sqlalchemy.create_engine("mssql+pyodbc:///?odbc_connect={}".format(params))

### Load data and transfer data files.
try:
        # Check the files in the inbound folder
    message = ''
    query = '''Select Filename, Keyword, FileType, SheetName, StandardName, DataSource, Frequency
                   from [BI_Eisai_Dayvigo].[dbo].[Eisai_Dayvigo_Data_Feed] where [IsActive] = 1'''
    files = pd.read_sql(query, conn)
    kwlist = files['Filename'].tolist()
    #print(kwlist)

    if len(os.listdir(inbound)) != 0 and len(os.listdir(inbound)) < 2:
            f = []
            for kw in kwlist:
                k = 0
                for name in os.listdir(inbound):
                    if kw.upper() in name.upper():
                        k = 1
                if k == 0:
                    f.append(kw)
            if len(f) > 1:
                message = "Hi, the files '{}' are missing from data feed list. Please check the files. Thanks --BI Team".format(', '.join(f))
              #  send_mail(From, To, CC, BCC, Subject, message)
    #print(message)
    elif len(os.listdir(inbound)) != 0:
        message = "Hi, there are no files in the data feed. Please check. Thanks --BI Team"
      #  send_mail(From, To, CC, BCC, Subject, message)
    if len(os.listdir(inbound)) ==2:
        if len(os.listdir(outbound)) != 0:
                for fpath in glob.glob(outbound + r'\*.*'):
                    os.remove(fpath)
                # Load.Convert data files
                tbl1 = ''
                tbl2 = ''
                i,j=0,0
                EM = 'No'
        EM = 'No'
        for name in os.listdir(inbound):
            if 'IC_Reports_AgentSummary'.upper() in name.upper():
                # print(name)
                cur = conn.cursor()
                cur.execute('truncate table [BI_TotalCare_ENPP].[dbo].[Amgen_ENPP_Contact_Center_IC_Report_Daily_Summary_Stg]')
                cur.commit()
                df = pd.read_excel(os.path.join(inbound, name), sheet_name = 'Sheet1')
             #   print(df)
                filetime = datetime.fromtimestamp(os.stat(os.path.join(inbound, name)).st_mtime).strftime('%Y-%m-%d %H:%M:%S')
               # print(filetime)
                df['Src_FileName'] = name
                df['Src_FileDate'] = filetime
                df['Src_FileLoadDate'] = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
              #  print(df)
                df.to_sql('Amgen_ENPP_Contact_Center_IC_Report_Daily_Summary_Stg',engine,schema='dbo',if_exists='append',index=False)  
                cur.execute("declare @EM varchar(4000); Exec dbo.usp_Load_Contact_Center_IC_Reports 'IC_Report_Daily_Summary', @Error = @EM OUTPUT; select @EM")
               # rows = cur.fetchall()
                #EM = rows[0][0]
                if EM =='No':
                    cur.commit()
                else:
                    message = "Hi, there is an error when Loading Contact Center IC Reports file. Please check it. Error details: {}".format(EM)
                    Subject = 'Load_Contact_Center_IC_Reports_Failed, Please Check'
                 #   send_mail(From, To, CC, BCC, Subject, message)
                    break
            if 'IC_Reports_AgentUnavailableTime'.upper() in name.upper():
                # print(name)
                cur = conn.cursor()
                cur.execute('truncate table [BI_TotalCare_ENPP].[dbo].[Amgen_ENPP_Contact_Center_IC_Reports_AgentUnavailableTime_Daily_Stg]')
                cur.commit()
                df = pd.read_excel(os.path.join(inbound, name), sheet_name = 'Sheet1', dtype={'Percent': float})
                df['Agent Name (ID)']= df['Agent Name (ID)'].fillna(method='ffill')
              #  print(df)
                filetime = datetime.fromtimestamp(os.stat(os.path.join(inbound, name)).st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                #print(filetime)
                df['Src_FileName'] = name
                df['Src_FileDate'] = filetime
                df['Src_FileLoadDate'] = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
              #  print(df)
                #df.rename(columns={'Product/Medical Device Tech. Complaint':'Product/Medical Device Tech# Complaint'}, inplace=True)                 
                df.to_sql('Amgen_ENPP_Contact_Center_IC_Reports_AgentUnavailableTime_Daily_Stg',engine,schema='dbo',if_exists='append',index=False)  
                cur.execute("declare @EM varchar(4000); Exec dbo.usp_Load_Contact_Center_IC_Reports 'AgentUnavailableTime', @Error = @EM OUTPUT; select @EM")
                #rows = cur.fetchall()
                #EM = rows[0][0]
                if EM =='No':
                    cur.commit()
                else:
                    message = "Hi, there is an error when Loading Contact Center IC Reports file. Please check it. Error details: {}".format(EM)
                    Subject = 'Load_Contact_Center_IC_Reports_Failed, Please Check'
                  #  send_mail(From, To, CC, BCC, Subject, message)
                    break
except:
    logging.exception("Error Details:")
    Subject = 'Trak360_Data_Of_Eisai_Dayvigo_Loaded_To_SDG_Failed, Please Check'
    message = '''Hi, there is an error during the data loading process. Please check the log file to debug in the folder G:\Commercial Sales\BI Supplement\Eisai_Dayvigo\Archive\Log. Thanks. '''
   # send_mail(From, To, CC, BCC, Subject, message)
        
conn.close()   