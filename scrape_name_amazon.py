import requests
from bs4 import BeautifulSoup
import re
#import selenium

page=requests.get("http://www.amazon.in/s/ref=nb_sb_noss_2?url=search-alias%3Daps&field-keywords=mobiles")
soup=BeautifulSoup(page.content,'html.parser')
while 1:
    containers =soup.find_all(class_="s-item-container")
    for x in containers:
        center=x.find('div',class_="a-row a-spacing-none")
        names=x.find('h2').get_text()
        price=x.find(class_="a-size-base a-color-price s-price a-text-bold")
        price=str(price)
        print(names)
        print(price)
    link=soup.find(id="pagnNextLink")["href"]
    if link is None:
        break
    link="www.amazon.in"+link
    page=requests.get(link)
    soup=BeautifulSoup(page.content,'html.parser')