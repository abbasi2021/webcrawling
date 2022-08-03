# -*- coding: utf-8 -*-
"""
@author: Z.Abbasi
"""


import schedule
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import requests
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from persiantools.jdatetime import JalaliDate
from datetime import datetime,date,timedelta
import pandas as pd
user='#########'
password='########'
executable_path=r"C:\Users\z.abbasi\Documents\chromedriver.exe"
def ttac(user,password,executable_path):
    executable_path=executable_path
    options=webdriver.ChromeOptions()

    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
#    options.add_argument('--headless')
    driver = webdriver.Chrome(executable_path)
    url="https://report.ttac.ir/IFDADashboard#/home/CustomPayments_OS"
    driver.get(url)
    username=driver.find_element_by_id("username")
    username.send_keys(user)
    pw=driver.find_element_by_id("password")
    pw.send_keys(password)
    enter_key=driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div[2]/form/div[3]/input")
    enter_key.click()
    today=JalaliDate(datetime.today().date()) 
    today=str(today)
    year,month,day=today.split('-')
    yesterday = datetime.today().date()- timedelta(1)
    yesterday_j=JalaliDate(yesterday)
    yesterday_j=str(yesterday_j)
    y_year,y_month,y_day=yesterday_j.split('-')
    print(y_year,y_month,y_day)
    for i in range(1,17):

        link_excel="""https://report.ttac.ir/CustomPayment_OS/ExportToExcelForCustomPaymentStatus?clientFromCreationDate={}%2F{}%2F{}&clientToCreationDate={}%2F{}%2F{}&companyId=&technicalAssistant_NationalCode=&officeId=&entranceCustomsCode=&registeredProduct_Name=&isFirst=false&hsType={}""".format(y_year,y_month,y_day,year,month,day,i)
        driver.get(link_excel)
    driver.quit()
    
    df1=pd.read_excel(r'C:\Users\z.abbasi\Downloads\CustomPayment.xlsx')
    for i in range(1,16):
        df2=pd.read_excel(r'C:\Users\z.abbasi\Downloads\CustomPayment (%i).xlsx'%i)
        df1=pd.concat([df1,df2],ignore_index=True)
    df1.to_excel(r"C:\Users\z.abbasi\Downloads\ttac_%s.xlsx"%today)
    
schedule.every().day.at("17:34").do(ttac,user='ma.ghasemi',password='@Mary1837608',executable_path=r"C:\Users\z.abbasi\Documents\chromedriver.exe")
while True:
    schedule.run_pending()
    #time.sleep(1)