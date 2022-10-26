import time
import json
import pandas as pd
import shutil
import glob
import os
import re
import smtplib
from os.path import basename
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.header import Header
from email.utils import COMMASPACE, formatdate, formataddr
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

from SendMail import send_mail


From = formataddr((str(Header('InContactDailyFeed','utf-8')),'ETLAlert@xxx.com'))
To = ''
CC = ''
BCC = ''
Subject = 'Summary of InContactDailyFeed Reports Process Details'
eSubject = 'InContactDailyFeed Reports Processed with Error'

download = r'G:\ACMS Nurse Partner\Amgen ENPP\Contact_Center_IC_Reports\Download\Daily'
inbound = r'G:\ACMS Nurse Partner\Amgen ENPP\Contact_Center_IC_Reports\Inbound\Daily'
# outbound = r'G:\ACMS Nurse Partner\Amgen ENPP\InContactDailyFeed\Outbound'
archive = r'G:\ACMS Nurse Partner\Amgen ENPP\Contact_Center_IC_Reports\Archive\Daily'
driverpath = r'G:\ChromeDriver\chromedriver.exe'
names = ['Agent Summary','Agent Unavailable Time']

### Check if any file exists in the download folder
if len(os.listdir(download)) == 0:
    # Set the max number of attempts
    atts = 3
    n = 0
    
    while n < atts:
        try:
            ### Delete any file in the download folder if exists
            if len(os.listdir(download)) != 0:
                for fname in glob.glob(download+r'\*'):
                    os.remove(fname)

            ### Connect to the website
            chrome_options = webdriver.ChromeOptions()
            prefs = {'download.default_directory':download}
            chrome_options.add_experimental_option('prefs',prefs)
            ###chrome_options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
            ###chrome_options.add_argument('--disable-gpu')
            ###chrome_options.add_argument('--headless')
            driver = webdriver.Chrome(executable_path=driverpath, options=chrome_options)

            i = 0
            for name in names:
                driver.get('https://home-c16.incontact.com/inContact/Manage/Reports/PrebuiltReports.aspx')
                driver.implicitly_wait(30)

                # Login to the account
                if i == 0:
                    login = json.load(open('login.json'))
                    username = login['username']
                    password = login['password']

                    user_box = driver.find_element_by_class_name('userName')
                    next_button = driver.find_element_by_id('ctl00_BaseContent_btnNext')
                    user_box.send_keys(username)
                    driver.implicitly_wait(10)
                    next_button.click()
                    driver.implicitly_wait(10)

                    pwd_box = driver.find_element_by_class_name('password')
                    submit_button = driver.find_element_by_id('ctl00_BaseContent_mslp_btnLogin')
                    pwd_box.send_keys(password)
                    driver.implicitly_wait(10)
                    submit_button.click()
                    driver.implicitly_wait(60)

                # Search the report list on the page
                items = driver.find_element_by_tag_name('tbody').find_elements_by_class_name('grid-column-1')
                for item in items:
                    report = item.find_element_by_tag_name('a').get_attribute('innerHTML').strip()

                    if report == name:
                        item.find_element_by_tag_name('a').click()
                        driver.implicitly_wait(10)

                        driver.find_element_by_class_name('ui-daterangepicker-prev').click()
                        driver.implicitly_wait(10)
                        
                        ### Check the report date (date range = yesterday)
                        yesterday = datetime.today().date() + timedelta(days=-1)
                        datevalue = driver.find_element_by_class_name('datetimepicker-inputbox').get_attribute('value').strip()
                        pattern = re.compile(r'\d{1,2}/\d{1,2}/\d{2,4}')
                        match = pattern.search(datevalue)
                        rdate = datetime.strptime(match.group(),'%m/%d/%Y')
                        
                        if rdate.date() == yesterday:
                            span_elements = driver.find_elements_by_tag_name('span')
                            for span in span_elements:
                                showoption = span.get_attribute('innerHTML').strip()
                                if showoption == 'Show Options':
                                    span.click()
                                    driver.implicitly_wait(10)
                                    time.sleep(2)

                                    t = 0
                                    td_elements = driver.find_elements_by_tag_name('td')
                                    for td in td_elements:
                                        if td.get_attribute('class').strip() == 'Header' and td.find_element_by_tag_name('span').get_attribute('innerHTML').strip() == 'Teams':
                                            td_elements[t+1].find_element_by_class_name('AddItem').click()
                                            driver.implicitly_wait(10)
                                            time.sleep(2)

                                            selects = driver.find_elements_by_class_name('msi-control-results')
                                            for select in selects:
                                                try:
                                                    # Only one option available or first option is 'Ashfield Team'
                                                    option = select.find_element_by_tag_name('option').get_attribute('innerHTML').strip()
                                                    if option == 'Ashfield Team':
                                                        select.find_element_by_tag_name('option').click()
                                                        driver.implicitly_wait(3)

                                                        driver.find_element_by_xpath('//*[@id="modalImageColumn"]/div/div[3]/div[2]/div/button[1]/span').click()
                                                        driver.find_element_by_xpath('//*[@id="modalImageColumn"]/div/div[6]/div[2]/button[1]/span').click()
                                                        driver.implicitly_wait(3)

                                                        spans = driver.find_elements_by_tag_name('span')
                                                        for span in spans:
                                                            if span.get_attribute('innerHTML').strip() == 'Run Report':
                                                                span.click()
                                                                driver.implicitly_wait(10)
                                                                time.sleep(2)

                                                                # Download the excel report
                                                                save_btn = driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_BaseContent_ReportContent_reportViewerControl_ctl05_ctl04_ctl00_ButtonImg"]')
                                                                save_btn.click()
                                                                driver.implicitly_wait(10)

                                                                save_options = driver.find_elements_by_tag_name('a')
                                                                for save_option in save_options:
                                                                    try:
                                                                        save_option_fmt = save_option.get_attribute('innerHTML').strip()   
                                                                        if save_option_fmt == 'Excel':
                                                                            save_option.click()
                                                                            time.sleep(10)
                                                                            i = i + 1

                                                                            break
                                                                    except:
                                                                        continue

                                                                break

                                                        break

                                                except:
                                                    continue

                                            break

                                        t = t + 1

                                    break

                            break
                            
                        else:
                            message = 'Hi, there is an error in the date range of the report. Please check the website. Thanks.'
                            send_mail(From, To, CC, BCC, eSubject, message)
#                             print(message)
            
            time.sleep(5)
            driver.quit()
            
            # Check if all the reports have been downloaded and rename file names         
            if i == len(names):
                fdate = rdate.strftime('%m%d%Y')
                for file in glob.glob(download+r'\*.xlsx'):
                    root,ext = os.path.splitext(file)
                    destfile = root + '_' + fdate + ext
                    os.rename(file,destfile)
                time.sleep(2)
                    
                files = os.listdir(download)
                # Move the files from the download folder to the inbound folder
                for fp in files:
                    shutil.move(os.path.join(download,fp),os.path.join(inbound,fp))
                time.sleep(2)
                
                message = f"Hi, InContactDailyFeed reports ({i} files) - {', '.join(files)} have been downloaded from the website successfully. Thanks."
                send_mail(From, To, CC, BCC, Subject, message)
#                 print(message)
            else:
                message = f"Hi, {len(names)} InContactDailyFeed reports are supposed to download from the website, but only {i} report(s) are downloaded. Please check it. Thanks."
                send_mail(From, To, CC, BCC, eSubject, message)
#                 print(message)
            
            break
            
        except Exception as e:
            driver.quit()
            n = n + 1
            if n >= atts:
                message = f'Hi, {n} attempts to download InContactDailyFeed reports failed. Please check the Internet connection. Thanks. Error details: {e}'
                send_mail(From, To, CC, BCC, eSubject, message)
#                 print(message)
else:
    send_mail(From, To, CC, BCC, eSubject, 'Hi, there are files already exists in the download folder. Please empty the download folder before you run this program. Thanks.')
#     print('File Exists!')
        
