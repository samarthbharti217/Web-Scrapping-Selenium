# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
import re
import smtplib

flippage=requests.get("https://www.flipkart.com/moto-g5-plus-gold-32-gb/p/itmes2zjvwfncxxr?pid=MOBEQHMGED7F9CZ2&srno=b_1_1&otracker=hp_omu_Top%20Offers_2_Now%20%E2%82%B914%2C999*_61747eaa-eb5e-4b80-bea5-19f892ce5999&lid=LSTMOBEQHMGED7F9CZ2KHTBI8")
flipsoup=BeautifulSoup(flippage.content,'html.parser')
flipprice =flipsoup.find(class_="_1vC4OE _37U4_g")
flipprice=str(flipprice)
flipprice = re.sub('<div class="_1vC4OE _37U4_g" data-reactid="230"><!-- react-text: 231 -->₹<!-- /react-text --><!-- react-text: 232 -->', '', flipprice)
flipprice = re.sub('<!-- /react-text --></div>', '', flipprice)
delivery=flipsoup.find(class_="_29Zp1s")
delivery=str(delivery)
delivery = re.sub('<div class="_29Zp1s" data-reactid="384"><!-- react-text: 385 --><!-- /react-text --><!-- react-text: 386 -->', '', delivery)
delivery = re.sub('Enter pincode for exact delivery dates/charges<!-- /react-text --><!-- react-text: 387 --><!-- /react-text --></div>', '', delivery)
print("Flipkart price= "+flipprice)
print(delivery)
 
#server = smtplib.SMTP('smtp.gmail.com', 587)
#server.starttls()
#server.login("sb887@snu.edu.in", "S@m9972713843")
 
#msg = "Flipkart price= "+flipprice+"\n"+delivery;
#server.sendmail("sb887@snu.edu.in", "samarthbharti97@gmail.com", msg)
#server.quit()
