import pandas as pd
import os
import sys
import math
import datetime
import re
import itertools
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import requests
import time

import requests

def write_to_delimited_file(name,journal):
    write_string = ''
    lis_names=name.split(".,")
    for i in lis_names:
        write_string=journal+"|"+i+"\n"
        try:
            file = open("D:\\List_of_Proffessors.txt" , 'a')
            file.write(write_string)
            file.close()
        except:
            print('File contained unrecordable content')


pandas_file = pd.read_csv("C:\\Users\\TradingLab13\\Desktop\\Journal_not_found.CSV", encoding = "ISO-8859-1")
driver = webdriver.Chrome("C:\\Users\\TradingLab13\\Desktop\\chromedriver.exe")
driver.maximize_window()


for i in range(len(pandas_file.index)):
    name=pandas_file.iloc[i,1]
    driver.get("https://www.scopus.com/sources.uri?zone=TopNavBar&origin=sourceinfo")
    driver.find_element_by_id("source-name").send_keys(name.strip())
    driver.find_element_by_id("searchTermsSubmit").click()
    table=driver.find_element_by_id("sourceList")
    t_body=table.find_element_by_tag_name('tbody')
    try:
        link=t_body.find_element_by_link_text(name.strip()).click()
    except:
        print(name+" not found")
        file = open("D:\\count_of_files.txt" , 'a')
        file.write(name+"|"+"Not Found\n")
        file.close()
        continue
    print("t s")
    time.sleep(2)
    print("t e")
    driver.find_element_by_id("viewSourceDocs").click()
    head=driver.find_element_by_class_name("documentHeader")
    count=head.find_element_by_class_name("resultsCount").text
    file = open("D:\\count_of_files.txt" , 'a')
    file.write(name+"|"+count+"\n")
    file.close()
    button=driver.find_element_by_id("resultsPerPage-button")
    button.click()
    driver.find_element_by_id("ui-id-4").click()
    while(1):
        try:
            table=driver.find_element_by_id("srchResultsList")
        except:
            break
        t_body=table.find_element_by_tag_name('tbody')
        time.sleep(2)
        t_r=t_body.find_elements_by_tag_name('tr')
        i=0
        while(i<len(t_r)):
            t_d=t_r[i].find_elements_by_tag_name('td')
            present_names=t_d[1].text
            write_to_delimited_file(present_names,name)
            print(present_names)
            i=i+3
        try:
            driver.find_element_by_xpath('//*[@title="Next page"]').click()
        except:
            break
    
    
    