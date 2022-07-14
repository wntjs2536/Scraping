from bs4 import BeautifulSoup
from selenium import webdriver
import pandas as pd #csv 모듈
import time    #지연시간 모듈
import sys  #파일 읽기,쓰기 모듈
import re
import os

#--------------------- 경로 지정 ----------------------------------------
now = time.strftime('%y-%m-%d_%H-%M-%S') #현재시간
driver = webdriver.Chrome('./chromedriver.exe') #크롬드라이버 경로

print('--------------------------------------------------------------')
print('연습문제 8-1 :네이버 영화 리뷰 정보 수집하기')
print('==============================================================')
print('')
print('')
print('')

sreach_input_str = str(input("검색명: "))
review_input_count = int(input("리뷰건수: "))
save_path = ('./output/')
os.makedirs(save_path + now + ' ' + sreach_input_str)
save_path = (save_path + now + ' ' + sreach_input_str)
print("결과 저장경로: " + save_path)

driver.get('https://movie.naver.com')   #타겟 주소
time.sleep(2)

#--------------------- 검색 영역----------------------------------------
elemnet = driver.find_element_by_id('ipt_tx_srch') 
elemnet.send_keys(sreach_input_str) 
driver.find_element_by_class_name('btn_srch').click() 
driver.find_element_by_xpath('//*[@id="old_content"]/ul[1]/li[2]/a').click()
time.sleep(2)

#--------------------- 지정 영화 진입 영역----------------------------------------
driver.find_element_by_class_name('result_thumb').click()
driver.find_element_by_class_name('end_sub_tab').click()
time.sleep(2)

#--------------------- 리뷰 아이프레임 진입 영역----------------------------------------
iframe = driver.switch_to.frame("pointAfterListIframe")
soup =  BeautifulSoup(driver.page_source, 'html.parser')

review_count = soup.find('div', class_='score_total').find('strong', class_='total').find('em').get_text()
review_count = re.findall("\d",review_count)
review_count = (''.join(review_count))
review_count = int(review_count)

f = open(save_path+'/'+ now + ' ' + sreach_input_str +'.txt', 'a', encoding='utf-8') #쓰기 모드
print(save_path+'/'+ now + ' ' + sreach_input_str +'.txt')

if review_count < review_input_count:
    print("--------------------------------------------------------------")
    print("리뷰건수 초과" + str(review_count) + "건만 수집합니다.")
    print("==============================================================")
    
    f.write("--------------------------------------------------------------\n")
    f.write("리뷰건수 초과" + str(review_count) + "건만 수집합니다.\n")
    f.write("==============================================================\n")
    review_input_count = review_count
    
print("")
f.write("")

roof_count = 0
page_count = 0

review_point = []
review_contents = []
review_id = []
review_date = []
review_good = []
review_bad = []

#old_content > ul.search_list_1 > li:nth-child(1)


for i in range (0, review_input_count):
    roof_count += 1
    
    print("--------------------------------------------------------------")
    print("총 " + str(review_input_count) + "건 중" + str(i+1) + " 번째 리뷰 데이터를 수집합니다.")
    print("==============================================================")
    print("")

    f.write("--------------------------------------------------------------\n")
    f.write("총 " + str(review_input_count) + "건 중" + str(i+1) + " 번째 리뷰 데이터를 수집합니다.\n")
    f.write("==============================================================\n")
    f.write("\n")

    if roof_count != 10 :
#-------------------------------------- 리뷰 아이프레임 진입 영역----------------------------------------
        review_point_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child(' + str(roof_count) + ')> div.star_score > em')[0].text
        review_contents_temp = soup.select('#_filtered_ment_' + str(roof_count-1))[0].text
        review_contents_temp = review_contents_temp.strip()
        review_id_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child('+ str(roof_count) +') > div.score_reple > dl > dt > em:nth-child(1) > a > span')[0].text                   
        review_date_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child('+ str(roof_count) +') > div.score_reple > dl > dt > em:nth-child(2)')[0].text
        review_good_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child('+ str(roof_count) +') > div.btn_area > a._sympathyButton > strong')[0].text
        review_bad_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child('+ str(roof_count) +') > div.btn_area > a._notSympathyButton > strong')[0].text
        
        print("별점:" + review_point_temp)
        f.write("별점:" + review_point_temp + "\n")
        print("리뷰내용:" + review_contents_temp)
        f.write("별점:" + review_contents_temp + "\n")
        print("작성자: " + review_id_temp)
        f.write("별점:" + review_id_temp + "\n")
        print("작성일자:" + review_date_temp)
        f.write("별점:" + review_date_temp + "\n")
        print("공감: " + review_good_temp)
        f.write("별점:" + review_good_temp + "\n")
        print("비공감: " + review_bad_temp)
        f.write("별점:" + review_bad_temp + "\n")
        print("")
        
        review_point.append(review_point_temp)
        review_contents.append(review_contents_temp)
        review_id.append(review_id_temp)
        review_date.append(review_date_temp)
        review_good.append(review_good_temp)
        review_bad.append(review_bad_temp)
        
    else:
#-------------------------------------- 페이지 넘기기 전 마지막 내용 수집 후 페이지 넘김 ----------------------------------------        
        review_point_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child(' + str(roof_count) + ')> div.star_score > em')[0].text
        review_contents_temp = soup.select('#_filtered_ment_' + str(roof_count-1))[0].text
        review_contents_temp = review_contents_temp.strip()
        review_id_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child('+ str(roof_count) +') > div.score_reple > dl > dt > em:nth-child(1) > a > span')[0].text                   
        review_date_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child('+ str(roof_count) +') > div.score_reple > dl > dt > em:nth-child(2)')[0].text
        review_good_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child('+ str(roof_count) +') > div.btn_area > a._sympathyButton > strong')[0].text
        review_bad_temp = soup.select('body > div > div > div.score_result > ul > li:nth-child('+ str(roof_count) +') > div.btn_area > a._notSympathyButton > strong')[0].text
        
        print("별점:" + review_point_temp)
        f.write("별점:" + review_point_temp + "\n")
        print("리뷰내용:" + review_contents_temp)
        f.write("별점:" + review_contents_temp + "\n")
        print("작성자: " + review_id_temp)
        f.write("별점:" + review_id_temp + "\n")
        print("작성일자:" + review_date_temp)
        f.write("별점:" + review_date_temp + "\n")
        print("공감: " + review_good_temp)
        f.write("별점:" + review_good_temp + "\n")
        print("비공감: " + review_bad_temp)
        f.write("별점:" + review_bad_temp + "\n")
        print("")
        
        review_point.append(review_point_temp)
        review_contents.append(review_contents_temp)
        review_id.append(review_id_temp)
        review_date.append(review_date_temp)
        review_good.append(review_good_temp)
        review_bad.append(review_bad_temp)

#-------------------------------------- 카운트 초기화  ---------------------------------------- 
        driver.find_element_by_link_text('''다음''').click()
        time.sleep(2)
        soup =  BeautifulSoup(driver.page_source, 'html.parser')
        roof_count = 0
f.close()

save_file = pd.DataFrame()
save_file['별점']=pd.Series(review_point)
save_file['리뷰내용']=pd.Series(review_contents)
save_file['작성자']=pd.Series(review_id)
save_file['작성일자']=pd.Series(review_date)
save_file['공감횟수']=pd.Series(review_good)
save_file['비공감횟수']=pd.Series(review_bad)
save_file.to_csv(save_path+ '/' + now + ' ' + sreach_input_str +'.csv', encoding="utf-8-sig",index=True)
save_file.to_excel(save_path+ '/' + now + ' ' + sreach_input_str +'.xls' ,index=True)

driver.close( )
print('크롤링 종료')
os.system("pause")
