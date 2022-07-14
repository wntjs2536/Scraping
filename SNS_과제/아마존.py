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
driver.get('https://www.amazon.com/bestsellers?ld=NSGoogle')   #타겟 주소

print("=" *80)
print("아마존 닷컴의 분야별 Best Seller 상품 정보 추출하기")
print("=" *80)

select_num = input('''
    1.Amazon Devices & Accessories     2.Amazon Launchpad            3.Appliances
    4.Apps & Games                     5.Arts, Crafts & Sewing       6.Audible Books & Originals
    7.Automotive                       8.Baby                        9.Beauty & Personal Care      
    10.Books                           11.CDs & Vinyl                12.Camera & Photo             
    13.Cell Phones & Accessories       14.Clothing, Shoes & Jewelry  15.Collectible Currencies       
    16.Computers & Accessories         17.Digital Music              18.Electronics                
    19.Entertainment Collectibles      20.Gift Cards                 21.Grocery & Gourmet Food     
    22.Handmade Products               23.Health & Household         24.Home & Kitchen             
    25.Industrial & Scientific         26.Kindle Store               27.Kitchen & Dining           
    28.Magazine Subscriptions          29.Movies & TV                30.Musical Instruments        
    31.Office Products                 32.Patio, Lawn & Garden       33.Pet Supplies               
    34.Prime Pantry                    35.Smart Home                 36.Software                   
    37.Sports & Outdoors               38.Sports Collectibles        39.Tools & Home Improvement   
    40.Toys & Games                    41.Video Games

    1.위 분야 중에서 자료를 수집할 분야의 번호를  선택하세요: ''')

select_count = int(input('2.해당 분야에서 크롤링 할 건수는 몇건입니까?(1-100 건 사이 입력): '))
if select_count > 100:
      print ("베스트 항목은 최대 100개 입니다.")
      select_count = 100


print("\n")

if select_num == '1' :
      select_title='Amazon Devices and Accessories'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[1]/a""").click( )

elif select_num =='2' :
      select_title='Amazon Launchpad'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[2]/a""").click( )

elif select_num =='3' :
      select_title='Appliances'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[3]/a""").click( )

elif select_num =='4' :
      select_title='Apps and Games'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[4]/a""").click( )

elif select_num =='5' :
      select_title='Arts and Crafts and Sewing'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[5]/a""").click( )

elif select_num =='6' :
      select_title='Audible Books and Originals'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[6]/a""").click( )

elif select_num =='7' :
      select_title='Automotive'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[7]/a""").click( )

elif select_num =='8' :
      select_title='Baby'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[8]/a""").click( )

elif select_num =='9' :
      select_title='Beauty and Personal Care'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[9]/a""").click( )

elif select_num =='10' :
      select_title='Books'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[10]/a""").click( )

elif select_num =='11' :
      select_title='CDs and Vinyl'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[11]/a""").click( )

elif select_num =='12' :
      select_title='Camera and Photo'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[12]/a""").click( )

elif select_num =='13' :
      select_title='Cell Phones and Accessories'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[13]/a""").click( )

elif select_num =='14' :
      select_title='Clothing and Shoes and Jewelry'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[14]/a""").click( )

elif select_num =='15' :
      select_title='Collectible Currencies'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[15]/a""").click( )

elif select_num =='16' :
      select_title='Computers and Accessories'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[16]/a""").click( )

elif select_num =='17' :
      select_title='Digital Music'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[17]/a""").click( )

elif select_num =='18' :
      select_title='Electronics'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[18]/a""").click( )

elif select_num =='19' :
      select_title='Entertainment Collectibles'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[19]/a""").click( )

elif select_num =='20' :
      select_title='Gift Cards'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[20]/a""").click( )

elif select_num =='21' :
      select_title='Grocery and Gourmet Food'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[21]/a""").click( )

elif select_num =='22' :
      select_title='Handmade Products'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[22]/a""").click( )

elif select_num =='23' :
      select_title='Health and Household'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[23]/a""").click( )

elif select_num =='24' :
      select_title='Home and Kitchen'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[24]/a""").click( )

elif select_num =='25' :
      select_title='Industrial and Scientific'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[25]/a""").click( )

elif select_num =='26' :
      select_title='Kindle Store'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[26]/a""").click( )

elif select_num =='27' :
      select_title='Kitchen and Dining'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[27]/a""").click( )

elif select_num =='28' :
      select_title='Magazine Subscriptions'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[28]/a""").click( )

elif select_num =='29' :
      select_title='Movies and TV'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[29]/a""").click( )

elif select_num =='30' :
      select_title='Musical Instruments'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[30]/a""").click( )

elif select_num =='31' :
      select_title='Office Products'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[31]/a""").click( )

elif select_num =='32' :
      select_title='Patio and Lawn and Garden'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[32]/a""").click( )

elif select_num =='33' :
      select_title='Pet Supplies'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[33]/a""").click( )

elif select_num =='34' :
      select_title='Prime Pantry'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[34]/a""").click( )

elif select_num =='35' :
      select_title='Smart Home'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[35]/a""").click( )

elif select_num =='36' :
      select_title='Software'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[36]/a""").click( )

elif select_num =='37' :
      select_title='Sports and Outdoors'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[37]/a""").click( )

elif select_num =='38' :
      select_title='Sports Collectibles'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[38]/a""").click( )

elif select_num =='39' :
      select_title='Tools and Home Improvemen'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[39]/a""").click( )

elif select_num =='40' :
      select_title='Toys and Games'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[40]/a""").click( )

elif select_num =='41' :
      select_title='Video Games'
      driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[41]/a""").click( )

save_path = './output/'
os.makedirs(save_path + now + '-아마존 닷컴-' + select_title)
save_path = (save_path + now + '-아마존 닷컴-' + select_title +'/')
os.makedirs(save_path +'image/')

driver.execute_script("window.scrollBy(0,9300);")
time.sleep(1)
soup = BeautifulSoup(driver.page_source, 'html.parser')


seller_rank = []
seller_name = []
seller_price = []
seller_review_count = []
seller_review_rate = []
seller_img = []

f = open(save_path + '아마존 베스트 상품.txt', 'a', encoding='utf-8')
roof_count = 0
for i in range (0, select_count):
      roof_count += 1

      if roof_count != 51 :

            print("--------------------------------------------------------------------------------------")
            f.write("--------------------------------------------------------------------------------------\n")
            
            try:
                  seller_rank_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > div > span.a-size-small.aok-float-left.zg-badge-body.zg-badge-color > span')[0].text #순위
                  seller_rank.append(seller_rank_temp)
                  print("판매순위: " +seller_rank_temp)
                  f.write("판매순위: " +seller_rank_temp +"\n")
            except:
                  roof_count +=1
                  
                  seller_rank_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > div > span.a-size-small.aok-float-left.zg-badge-body.zg-badge-color > span')[0].text #순위
                  seller_rank.append(seller_rank_temp)
                  print("판매순위: " +seller_rank_temp)
                  f.write("판매순위: " +seller_rank_temp+"\n")
            try:
                  seller_name_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > a > div')[0].text #상품명
                  seller_name.append(seller_name_temp)
                  print("제품명: " +seller_name_temp)
                  f.write("제품명: " +seller_name_temp+"\n")
            
                  seller_review_count_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > div.a-icon-row.a-spacing-none > a.a-size-small.a-link-normal')[0].text #리뷰 수
                  seller_review_count.append(seller_review_count_temp)
                  print("상품평 수: " +seller_review_count_temp)
                  f.write("상품평 수: " +seller_review_count_temp+"\n")

            except:
                  seller_review_count.append("상품평 없음")
                  print("상품평 수: 없음")
                  f.write("상품평 수: 없음"+"\n")
                  pass                  
            try:

                  seller_review_rate_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > div.a-icon-row.a-spacing-none > a:nth-child(1) > i')[0].text #평점
                  seller_review_rate.append(seller_review_rate_temp)
                  print("평점: " +seller_review_rate_temp)
                  f.write("평점: " +seller_review_rate_temp+"\n")

            except:
                  seller_review_rate.append('상품 점수 없음')
                  print("평점: 없음")
                  f.write("평점: 없음"+"\n")

            try:
                  seller_img_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > a > span > div') #이미지 영역 추출
                  seller_img_temp = seller_img_temp[0].find('img')["src"] #이미지 링크 추출
                  urllib.request.urlretrieve(seller_img_temp, save_path + 'image/' + str(i+1) + '.jpg')
                  seller_img.append(seller_img_temp)
                  ("이미지 주소:" +seller_img_temp)

            except:
                  print("이미지 찾을 수 없음")
                  seller_img.append("Mssing img")
                  
            try:
                  seller_price_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > div.a-row > a > span')[0].text #가격                                 
                  seller_price.append(seller_price_temp)
                  print("가격: " +seller_price_temp)
                  f.write("가격: " +seller_price_temp +"\n")
      
            except IndexError:
                  seller_price_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > div.a-row > a > span > span')[0].text
                  seller_price.append(seller_price_temp)
                  print("가격: " +seller_price_temp)
                  f.write("가격: " +seller_price_temp +"\n")
                  
      else:

            print("--------------------------------------------------------------------------------------")
            f.write("--------------------------------------------------------------------------------------\n")
            
            try:
                  seller_rank_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > div > span.a-size-small.aok-float-left.zg-badge-body.zg-badge-color > span')[0].text #순위
                  seller_rank.append(seller_rank_temp)
                  print("판매순위: " +seller_rank_temp)
                  f.write("판매순위: " +seller_rank_temp +"\n")
            except:
                  roof_count +=1
                  
                  seller_rank_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > div > span.a-size-small.aok-float-left.zg-badge-body.zg-badge-color > span')[0].text #순위
                  seller_rank.append(seller_rank_temp)
                  print("판매순위: " +seller_rank_temp)
                  f.write("판매순위: " +seller_rank_temp+"\n")
            try:
                  seller_name_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > a > div')[0].text #상품명
                  seller_name.append(seller_name_temp)
                  print("제품명: " +seller_name_temp)
                  f.write("제품명: " +seller_name_temp+"\n")
            
                  seller_review_count_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > div.a-icon-row.a-spacing-none > a.a-size-small.a-link-normal')[0].text #리뷰 수
                  seller_review_count.append(seller_review_count_temp)
                  print("상품평 수: " +seller_review_count_temp)
                  f.write("상품평 수: " +seller_review_count_temp+"\n")

            except:
                  seller_review_count.append("상품평 없음")
                  print("상품평 수: 없음")
                  f.write("상품평 수: 없음"+"\n")
                  pass                  
            try:

                  seller_review_rate_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > div.a-icon-row.a-spacing-none > a:nth-child(1) > i')[0].text #평점
                  seller_review_rate.append(seller_review_rate_temp)
                  print("평점: " +seller_review_rate_temp)
                  f.write("평점: " +seller_review_rate_temp+"\n")

            except:
                  seller_review_rate.append('상품 점수 없음')
                  print("평점: 없음")
                  f.write("평점: 없음"+"\n")

            try:
                  seller_img_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > a > span > div') #이미지 영역 추출
                  seller_img_temp = seller_img_temp[0].find('img')["src"] #이미지 링크 추출
                  urllib.request.urlretrieve(seller_img_temp, save_path + 'image/' + str(i+1) + '.jpg')
                  seller_img.append(seller_img_temp)
                  ("이미지 주소:" +seller_img_temp)

            except:
                  print("이미지 찾을 수 없음")
                  seller_img.append("Mssing img")
                  
            try:
                  seller_price_temp = soup.select('#zg-ordered-list > li:nth-child('+str(roof_count)+') > span > div > span > div.a-row > a > span')[0].text #가격                                 
                  seller_price.append(seller_price_temp)
                  print("가격: " +seller_price_temp)
                  f.write("가격: " +seller_price_temp +"\n")
      
            except IndexError :
                  seller_price.append("가격 없음")
                  print("가격: 없음" +seller_price_temp)
                  
            driver.find_element_by_xpath("""//*[@id="zg-center-div"]/div[2]/div/ul/li[3]/a""").click( )
            roof_count = 0
            time.sleep(2)
            soup =  BeautifulSoup(driver.page_source, 'html.parser')
            
f.close( )
#--------------------- xls, csv 파일 생성 ------------------------------------------------
wb = xw.Workbook(save_path + '아마존 베스트 상품.xls')
ws = wb.add_worksheet()

ws.set_default_row(80) # 행 크기

ws.write(0,0, "판매순위")
ws.write(0,1, "이미지")
ws.write(0,2, "상품명")
ws.write(0,3, "가격")                        
ws.write(0,4, "상품평 수")
ws.write(0,5, "상품 평점")

for i in range(0,select_count):
      
      ws.set_column(0, i+1, 30) # 열 크기
      
      image = Image.open(save_path + 'image/'+str(i+1)+'.jpg') # 이미지 로드
      resize_image = image.resize((100,100)) # 이미지 크기조정 
      resize_image.save(save_path + 'image/'+str(i+1)+'.jpg') # 이미지 저장
            
      ws.write(i+1, 0, seller_rank[i])
      ws.insert_image(i+1, 1, save_path + 'image/'+str(i+1)+'.jpg')
      ws.write(i+1, 2, seller_name[i])
      ws.write(i+1, 3, seller_price[i])
      ws.write(i+1, 4, seller_review_count[i])
      ws.write(i+1, 5, seller_review_rate[i])

wb.close()

shutil.copyfile(save_path+'아마존 베스트 상품.xls', save_path+"아마존 베스트 상품.xls") # csv 파일 생성
driver.close( )
print('크롤링 종료')
os.system("pause")
