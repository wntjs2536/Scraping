from bs4 import BeautifulSoup
from selenium import webdriver
from PIL import Image

import shutil
import xlsxwriter as xw
import requests
import urllib.request
import time    #지연시간 모듈
import sys  #파일 읽기,쓰기 모듈
import re
import os


#--------------------- 경로 지정 ----------------------------------------
now = time.strftime('%y-%m-%d_%H-%M-%S') #현재시간

driver = webdriver.Chrome('./chromedriver.exe') #크롬드라이버 경로


print('--------------------------------------------------------------')
print('G 마켓의 분야별 Best Seller 상품 정보 추출하기')
print('==============================================================')
print('')
print('')
print('')

save_path = './output/'
os.makedirs(save_path + now + '-' + 'G마켓')
save_path = (save_path + now + '-' + 'G마켓/')

driver.get('http://corners.gmarket.co.kr/Bestsellers')   #타겟 주소
time.sleep(2)


#--------------------- 문서 스크롤 (이미지, 상품 로딩) 후 크롤링 시작 ----------------------------------------

page_length = driver.execute_script("return document.body.scrollHeight") 
count = 0
while count < page_length:
    count +=900
    driver.execute_script('window.scrollTo(0, '+str(count)+')')
    print('상품 이미지를 로드 중 입니다. 기다려 주세요.'+str(count/page_length*100)+"%진행")
    time.sleep(1)
    
soup =  BeautifulSoup(driver.page_source, 'html.parser')

f = open(save_path + 'G마켓 베스트 상품.txt', 'a', encoding='utf-8') #쓰기 모드

os.makedirs(save_path + 'image/')

seller_rank = []
seller_name = []
seller_img = []
seller_price = []
seller_dis_price = []
seller_dis_rate = []

for i in range(0, 200):
    
    seller_rank_temp = soup.select('#no'+ str(i+1))[0].text   #순위                                
    seller_img_temp = soup.select('#gBestWrap > div > div:nth-child(5) > div:nth-child(3) > ul > li:nth-child('+ str(i+1) +') > div.thumb > a')
    seller_img_temp = seller_img_temp[0].find('img')["src"] #이미지 주소 추출
    urllib.request.urlretrieve(seller_img_temp, save_path + 'image/' + str(i+1) + '.jpg')   #이미지 다운로드                   
    
    seller_name_temp = soup.select('#gBestWrap > div > div:nth-child(5) > div:nth-child(3) > ul > li:nth-child('+ str(i+1) +') > a')[0].text   #상품명
#--------------------- 원가만 있을 때 원가, 최저가, 할인율 예외처리 ------------------------------------------------
    try:
        seller_price_temp = soup.select('#gBestWrap > div > div:nth-child(5) > div:nth-child(3) > ul > li:nth-child('+ str(i+1) +') > div.item_price > div.o-price > span > span')[0].text   #원가
    except: #원가만 있을 때
        seller_price_temp = soup.select('#gBestWrap > div > div:nth-child(5) > div:nth-child(3) > ul > li:nth-child('+ str(i+1) +') > div.item_price > div.s-price > strong > span > span')[0].text 
        
    seller_dis_price_temp = soup.select('#gBestWrap > div > div:nth-child(5) > div:nth-child(3) > ul > li:nth-child('+ str(i+1) +') > div.item_price > div.s-price > strong > span > span')[0].text   #할인가

    try: 
        seller_dis_rate_temp = soup.select('#gBestWrap > div > div:nth-child(5) > div:nth-child(3) > ul > li:nth-child('+ str(i+1) +') > div.item_price > div.s-price > span > em')[0].text   #할인율

    except: # 할인율 찾을 수 없을 때
        seller_dis_rate_temp = '할인율 없음.'

    print("-----------------------------------------------------------------------------------------------")
    print("판매 순위: " + seller_rank_temp)
    print("상품명: " + seller_name_temp)
    print("원가: " + seller_price_temp)
    print("최저가: " + seller_dis_price_temp)
    print("할인율: " + seller_dis_rate_temp)
    print("")

    f.write("-----------------------------------------------------------------------------------------------")
    f.write("판매 순위: "+ seller_rank_temp + "\n")
    f.write("상품명: "+ seller_name_temp + "\n")
    f.write("원가: "+ seller_price_temp + "\n")
    f.write("최저가: "+ seller_dis_price_temp + "\n")
    f.write("할인율: "+ seller_dis_rate_temp + "\n")
    f.write("\n")

    seller_rank.append(seller_rank_temp)
    seller_name.append(seller_name_temp)
    seller_img.append(seller_img_temp)
    seller_price.append(seller_price_temp)
    seller_dis_price.append(seller_dis_price_temp)
    seller_dis_rate.append(seller_dis_rate_temp)
    
f.close()

#--------------------- xls, csv 파일 생성 ------------------------------------------------
wb = xw.Workbook(save_path + 'G마켓 베스트 상품.xls')
ws = wb.add_worksheet()

ws.set_default_row(80) # 행 크기

ws.write(0,0, "판매 순위")
ws.write(0,1, "이미지")
ws.write(0,2, "제품명")
ws.write(0,3, "원가")                        
ws.write(0,4, "최저가")
ws.write(0,5, "할인율")

for i in range(0,200):
    ws.set_column(0, i+1, 30) # 열 크기
    
    image = Image.open(save_path + 'image/'+str(i+1)+'.jpg') # 이미지 로드
    resize_image = image.resize((100,100)) # 이미지 크기조정 
    resize_image.save(save_path + 'image/'+str(i+1)+'.jpg') # 이미지 저장

    ws.write(i+1, 0, seller_rank[i])
    ws.insert_image(i+1, 1, save_path + 'image/'+str(i+1)+'.jpg')
    ws.write(i+1, 2, seller_name[i])
    ws.write(i+1, 3, seller_price[i])
    ws.write(i+1, 4, seller_dis_price[i])
    ws.write(i+1, 5, seller_dis_rate[i])

wb.close()

shutil.copyfile(save_path+'G마켓 베스트 상품.xls', save_path+"G마켓 베스트 상품.csv") # csv 파일 생성
driver.close( )
print('크롤링 종료')
os.system("pause")
