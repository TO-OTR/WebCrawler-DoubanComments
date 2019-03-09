#-*- coding: utf-8 -*-
import urllib
import re
import HTMLParser
import xlwt
import requests
import time
import random
from bs4 import BeautifulSoup
from xlwt import Workbook
from fake_useragent import UserAgent
    

    

excel_name = u'douban_hot_review.xls'
sheet_name = u'豆瓣影评'
column = [u'标题',u'作者',u'影片',u'影评']

douban_excel = Workbook(excel_name)
douban_excel = Workbook(encoding='utf-8')
douban_sheet = douban_excel.add_sheet(sheet_name,cell_overwrite_ok=True)
douban_sheet.write(0,0,u'标题')
douban_sheet.write(0,1,u'作者')
douban_sheet.write(0,2,u'影片')
douban_sheet.write(0,3,u'影评')
douban_excel.save(excel_name)




def get_movie_review():
    #html = get_html(url, 1, 3)
  
    
#    titles = soup.select('.main-bd h2')
#    i = 0
#    for row in (range(1+page*10,11+page*10)):
#        douban_sheet.write(row,0,titles[i].text)
#        i = i + 1 #print(titles[i].text)
        
        
    
#    names = soup.select('.name')
#    i = 0
#    for row in (range(1+page*10,11+page*10)):
#        douban_sheet.write(row,1,names[i].text)
#        i = i + 1

    titles = soup.select('.main-bd h2')
    row = 1+page*10
    for i in range(10):
        douban_sheet.write(row,0,titles[i].text)
        row = row + 1
    
    
    names = soup.select('.name')
    row = 1+page*10
    for i in range(10):
        douban_sheet.write(row,1,names[i].text)
        row = row + 1
    
    
    films = soup.select('img')
    row = 1+page*10
    for i in (range(0,40,4)):
        douban_sheet.write(row,2,films[i]['title'])
        row = row + 1 #print(films[i]['title'])
              
    reviews = soup.select('.review-short')
    row = 1+page*10
    for review in reviews:
        reviewid = review['data-rid']#文章id
        #print(reviewid)
        reviewres = requests.get('https://movie.douban.com/review/'+reviewid)#访问文章所在网页
        reviewres.encoding = 'utf-8'
        reviewsoup = BeautifulSoup(reviewres.text,'html.parser')
        reviewtext = []
        for p in reviewsoup.select('.review-content p'):
            if len(p) > 0 :
                reviewtext.append(p.text.strip())
        douban_sheet.write(row,3,reviewtext)
        #print(reviewtext)
        row = row+1
    douban_excel.save(excel_name)
    time.sleep(random.random()*3)

for page in range(0,19,1):    
    urls = 'https://movie.douban.com/review/best/?start='
    url = (urls+str(page*10))
    print(url)
    ua=UserAgent()
    headers={'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
             'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9,zh-TW;q=0.8,fr;q=0.7',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Cookie':'gr_user_id=1001704b-7ce9-486c-96e3-e25cd73cac0c; _vwo_uuid_v2=6DA47C7BCC83936C435503463C7D3683|5bd8d91f871acedfccdc5c8b76d4b16c; __utmv=30149280.13030; douban-fav-remind=1; viewed="19162451_3117248"; ll="108296"; bid=B4ze21QchNk; __utmz=223695111.1543828035.32.30.utmcsr=python.jobbole.com|utmccn=(referral)|utmcmd=referral|utmcct=/88325/; ct=y; ps=y; push_noty_num=0; push_doumail_num=0; douban-profile-remind=1; ap_v=0,6.0; __utma=30149280.1416874245.1477755688.1544196340.1544203769.63; __utmc=30149280; __utmz=30149280.1544203769.63.49.utmcsr=baidu|utmccn=(organic)|utmcmd=organic; __utmt=1; __utmb=30149280.5.6.1544203769; _pk_ref.100001.4cf6=%5B%22%22%2C%22%22%2C1544203787%2C%22http%3A%2F%2Fpython.jobbole.com%2F88325%2F%22%5D; _pk_id.100001.4cf6=e58636e3d86c5459.1477755686.44.1544203787.1544197968.; _pk_ses.100001.4cf6=*; __utma=223695111.1764043885.1477755688.1544196340.1544203787.44; __utmb=223695111.0.10.1544203787; __utmc=223695111',
            'Host': 'movie.douban.com',
            'Upgrade-Insecure-Requests': '1',
            "User-Agent":ua.random}
    print(ua.random)
    res = requests.get(url=url,headers=headers,allow_redirects=False)
    print(res.text)
    soup = BeautifulSoup(res.text,'html.parser')
    res.encoding = 'utf-8'    
    get_movie_review()   
