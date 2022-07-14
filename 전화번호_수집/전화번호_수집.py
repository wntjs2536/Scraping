from selenium import webdriver
import pyautogui

from bs4 import BeautifulSoup
import pandas as pd

import tkinter
from tkinter import *
import tkinter.ttk as ttk
from tkinter import scrolledtext
import tkinter.messagebox

import threading
import time
import sys

now = time.strftime('%y-%m-%d_%H-%M-%S')

def th():
    th = threading.Thread(target=main)
    th.daemon = True
    th.start()
    
def main():
    global select_city_name, select_district_name
    
    if select_city_num == 0:
        for select_c in range(1, len(city_combox['values'])):
            select_city_name = city_combox['values'][select_c]
    
        city_combox.set('')
        district_combox.set('')


    elif select_district_num == 0:
        for select_d in range(1, len(district_combox['values'])):
            select_district_name = district_combox['values'][select_d]
            crawling()
            
        city_combox.set('')
        district_combox.set('')

    else:
        crawling()
        city_combox.set('')
        district_combox.set('')

            
def crawling():
    global driver, now
    
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-gpu')
    options.add_argument("--auto-open-devtools-for-tabs")
    #options.add_argument('--headless')
    driver = webdriver.Chrome('./chromedriver.exe', options=options)
    driver.get('https://www.114.co.kr/search/map?query='+str(select_city_name)+' '+str(select_district_name)+'')
    
    del_p_list()
    progressbar['value'] = 0

    start_button['state'] = 'disabled'
    end_button['state'] = 'normal'

    city_combox.configure(state='disabled')
    district_combox.configure(state='disabled')

    log_box.configure(state='normal')
    log_box.insert(END, "-"*45+"\n")
    log_box.insert(END, "\n작업시작\n")
    log_box.insert(END, "시작 시간: "+now+"\n")
    log_box.insert(END, "-"*45+"\n")
    log_box.insert(END, "선택 지역: "+str(select_city_name)+"\n선택 행정구: "+str(select_district_name))
    log_box.configure(state='disabled')
    log_box.see("end")

    for kind_numbox in range(1, 3):
        time.sleep(2)

        if kind_numbox == 1:
            driver.find_element_by_xpath('/html/body/div[1]/section/div/div/div[1]/ul/li[2]').click()
            element = driver.find_element_by_xpath('//body')
            time.sleep(2)

            pyautogui.hotkey('ctrl','shift','i')
            pyautogui.hotkey('ctrl','shift','i')
            time.sleep(1)
            soup =  BeautifulSoup(driver.page_source, 'html.parser')
            

        if kind_numbox == 2:
            driver.find_element_by_xpath('/html/body/div[1]/section/div/div/div[1]/ul/li[3]').click()
            element = driver.find_element_by_xpath('//body')
            time.sleep(2)

            pyautogui.hotkey('ctrl','shift','i')
            pyautogui.hotkey('ctrl','shift','i')
            time.sleep(1)
            soup =  BeautifulSoup(driver.page_source, 'html.parser')

            
        soup =  BeautifulSoup(driver.page_source, 'html.parser')
        all_company_list = soup.findAll('span', attrs={'class':'fc1'})[0].text
        all_company_list = all_company_list.replace(',','')
        all_company_list = int(all_company_list)
        progressbar['maximum'] = all_company_list

        log_box.configure(state='normal')    
        log_box.insert(END, "\n\n"+"총 : "+str(all_company_list)+" 건의 업체 검색 됨\n")
        log_box.insert(END, "-"*48+"\n")
        log_box.configure(state='disabled')
        log_box.see("end")
        
        all_page = int(all_company_list / 10)
        if all_company_list % 10 > 0:
            all_page += 1
            
        for count_page in range(0, all_page):

            soup =  BeautifulSoup(driver.page_source, 'html.parser')
            page_company_list = soup.findAll('li', attrs={'class':'map-item'})
            
            if count_page == int(all_company_list / 10) and all_company_list % 10 > 0 :
                all_page = int(all_company_list % 10)

            for count_list in range(0, len(page_company_list)):
                temp_name = page_company_list[count_list].findAll('strong', attrs={'class':'tit'})

                name = temp_name[0].findAll('span')[0].text
                try:
                    Sectors = temp_name[0].findAll('span', attrs={'class':'stit'})[0].text
                except:
                    Sectors = ""
                tel = page_company_list[count_list].findAll('p', attrs={'class':'tel'})[0].text

                try:
                    try:
                        addrass = page_company_list[count_list].findAll('p', attrs={'class':'add'})[0].text
                    except:
                        addrass = page_company_list[count_list].findAll('p', attrs={'class':'add2'})[0].text
                except :
                    adrass = " "
                    
                p_list.insert('', 'end', values=[select_city_name, select_district_name, name, Sectors, tel, addrass])
                p_list.yview_moveto(1)
                
                progressbar.step(1)

            now_p = int((count_page+1)*100 / all_page)
            progress_all_page.configure(text= '(' +str(select_city_name)+ ' '+str(select_district_name) + ') ' +str(all_page)+" / "+str(count_page+1))
            progress_p.configure(text= str(now_p)+' %')
            
            count_page += 1

            if count_page % 5 == 0:
                driver.find_element_by_class_name('paging').find_element_by_css_selector('.page.next').click()
                #time.sleep(0.5)
                
            else:
                try:
                    driver.find_element_by_class_name('paging').find_element_by_partial_link_text(str(count_page + 1)).click()
                    #time.sleep(0.5)
                except:
                    pass

    driver.quit()
    
    start_button['state'] = 'normal'
    end_button['state'] = 'disabled'

    city_combox.configure(state='normal')
    district_combox.configure(state='normal')

    progress_p.configure(text= '0 %')
    progress_all_page.configure(text= '0 / 0')

    now = time.strftime('%y-%m-%d_%H-%M-%S')
    
    log_box.configure(state='normal')    
    log_box.insert(END, "\n작업종료")
    log_box.insert(END, "\종료 시간: "+now+"\n")
    log_box.configure(state='disabled')
    log_box.see("end")

    save_p_list()
    
def on_closing():
    if tkinter.messagebox.askokcancel("종료", "정말로 종료 하시겠습니까?\n진행중인 사항은 저장되지 않습니다.",icon='warning'):
       window.destroy()
       driver.quit()

def stop():
    if tkinter.messagebox.askokcancel("취소", "정말로 취소 하시겠습니까?\n진행중인 사항은 저장되지 않습니다.",icon='warning'):
        progressbar['value'] = 0
        progress_p.configure(text= '0 %')
        progress_all_page.configure(text= '0 / 0')
        
        end_button['state'] = 'disabled'
        start_button['state'] = 'normal'
        
        city_combox.configure(state='normal')
        district_combox.configure(state='normal')

        
        del_p_list()
        driver.quit()

def del_p_list():
    for i in p_list.get_children():
       p_list.delete(i)
       
    log_box.configure(state='normal')
    log_box.delete("2.0","end")
    log_box.configure(state='disabled')
    window.update()

def save_p_list():
    all_p_list = []
    
    for child in p_list.get_children():
        all_p_list.append(p_list.item(child)["values"])
        
    save_pd = pd.DataFrame(all_p_list,columns=['지역', '행정구', '업체명', '업종','전화번호','주소'])
    save_pd.to_excel('./'+now+'_'+select_city_name+'_'+select_district_name +'.xlsx',index=True)

def del_com(event):
    district_combox.set('')

def select_combox(index, value, op):

    global select_city_num, select_city_name, select_district_name, select_district_num
    
    select_city_num = city_combox.current()
    select_city_name = city_combox.get()

    select_district_num = district_combox.current()
    select_district_name = district_combox.get()

    print(select_city_num, select_city_name, select_district_name, select_district_num)

    if district_combox.get() != '' and city_combox.bind("<<ComboboxSelected>>", del_com):
        print("지역 변경")
        
    if select_city_num == 0:
        district_combox.set('')
        district_combox.configure(state='disabled')

    if select_city_num == 1:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','강남구','강동구','강북구','강서구','관악구','광진구','구로구','금천구',
                                     '노원구','도봉구','동대문구','동작구','마포구','서대문구','서초구','성동구',
                                     '성북구','송파구','양천구','영등포구','용산구','은평구','종로구','중구','중랑구')
        
    if select_city_num == 2:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','강서구','금정구','기장군','남구','동구','동래구','부산진구','북구',
                                     '사상구','사하구','서구','수영구','연제구','영도구','중구','해운대구')
    if select_city_num == 3:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','남구','달서구','달성군','동구','북구','서구','수성구','중구')

    if select_city_num == 4:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','강화군','계양구','남구','남동구','동구','미추홀구','부평구','서구','연수구','옹진군','중구')

    if select_city_num == 5:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','광산구','남구','동구','북구','서구')

    if select_city_num == 6:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','대덕구','동구','서구','유성구','중구')

    if select_city_num == 7:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','남구','동구','북구','울주군','중구')
        
    if select_city_num == 8:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','강릉시','고성군','동해시','삼척시','속초시','양구군','양양군','영월군','원주시','인제군',
                                     '정선군','철원군','춘천시','태백시','평창군','홍천군','화천군','횡성군')

    if select_city_num == 9:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','가평군','고양시 덕양구','고양시 일산동구','고양시 일산서구','과천시','광명시','광주시','구리시',
                                     '군포시','김포시','남양주시','동두천시','부천시','부천시 소사구','부천시 오정구','부천시 원미구',
                                     '성남시 분당구','성남시 수정구','성남시 중원구','수원시 권선구','수원시 영통구','수영시 장안구','수원시 팔달구','시흥시','안산시 단원구',
                                     '안산시 상록구','안성시','안양시 동안구','안양시 만안구','양주시','양평군','여주시','연천군',
                                     '오산시','용인시 기흥구','용인시 수지구','용인시 처인구','의왕시','의정부시','이천시','파주시',
                                     '평택시','포천시','하남시','화성시')
        
    if select_city_num == 10:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','거제시','거창군','고성군','김해시','남해군','밀양시','사천시','산청군',
                                     '양산시','의령군','진주시','창녕군','창원시 마산합포구','창원시 마산회원구','창원시 성산구','창원시 의창구','창원시 진해구','통영시',
                                     '하동군','함안군','함양군','합천군')
    if select_city_num == 11:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','경산시','경주시','고령군','구미시','군위군','김천시','문경시',
                                     '봉화군','상주시','성주군','안동시','영덕군','영양군','영주시','영천시',
                                     '예천군','울릉군','울진군','의성군','청도군','청송군','칠곡군','포항시 남구',
                                     '포항시 북구')

    if select_city_num == 12:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','강진군','고흥군','곡성군','광양시','구례군','나주시','담양군','목포시','무안군','보성군','순천시',
                                     '신안군','여수시','영광군','영암군','완도군','장성군','장흥군','진도군','함평군','해남군','화순군')

    if select_city_num == 13:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','고창군','군산시','김제시','남원시','무주군','부안군',
                                     '순창군','완주군','익산시','임실군','장수군','전주시 덕진구','전주시 완산구','정읍시','진안군')

    if select_city_num == 14:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','서귀포시','제주시')

    if select_city_num == 15:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','계룡시','공주시','금산군','논산시','당진시','보령시',
                                     '부여군','서산시','서천군','아산시','예산군','천안시 동남구',
                                     '천안시 서북구','청양군','태안군','홍성군')
        
    if select_city_num == 16:
        district_combox.configure(state='normal')
        district_combox['values'] = ('전체','괴산군','단양군','보은군','영동군','옥천군','음성군',
                                     '제천시','증평군','진천군','청주시 상당구','청주시 서원구','청주시 청원구','청주시 흥덕구','충주시')

    if select_city_num == 17:
        district_combox.configure(state='disabled')
        district_combox['values'] = ('')
                         
window=tkinter.Tk()
window.title("전화번호 수집.exe")
window.resizable(False, False)

main_frame = Frame(window)
main_frame.pack()

top_frame = Frame(main_frame)# 상단 프레임
top_frame.pack(side = "top", fill = 'x')

top_left_frame = LabelFrame(top_frame, text="필터링")# 상단 프레임
top_left_frame.pack(side = "left", padx= 15)

top_right_frame = LabelFrame(top_frame, text="진행도")# 상단 프레임
top_right_frame.pack(side = "right", padx= 15)

bottom_frame = LabelFrame(main_frame, text="출력") # 하단 프레임
bottom_frame.pack(side = "bottom", fill = 'x')

bottom_right_frame = Frame(bottom_frame) # 하단 오른쪽 프레임
bottom_right_frame.pack(side = 'right')

bottom_left_frame = Frame(bottom_frame)
bottom_left_frame.pack(side = 'left')

progressbar = tkinter.ttk.Progressbar(top_right_frame, maximum=0, mode="determinate", length=320)
progressbar.pack(side = "top", fill = "x")

progress_all_page = tkinter.Label(top_right_frame, text = " 0 / 0")
progress_all_page.pack(side = "left")

progress_p = tkinter.Label(top_right_frame, text = "0 %")
progress_p.pack(side = "right")

select_city = StringVar()
select_city.trace('w',select_combox)

city_lable = tkinter.Label(top_left_frame, text = "지역 선택 ")
city_lable.pack(side = "left", padx= 15)

city_combox = ttk.Combobox(top_left_frame, state="readonly", values = ["전체","서울","부산","대구","인천","광주","대전","울산","강원","경기","경남","경북","전남","전북","제주","충남","충북","세종시"], width = 10, textvariable=select_city)
city_combox.pack(side = "left", padx= 15)

select_district = StringVar()
select_district.trace('w',select_combox)

district_lable = tkinter.Label(top_left_frame, text = "행정구")
district_lable.pack(side = "left", padx= 15)

district_combox = ttk.Combobox(top_left_frame, state="readonly", textvariable=select_district)
district_combox.pack(side = "left", padx= 15)

start_button = tkinter.Button(top_left_frame, overrelief="solid", width=15, text = "시작", command = th)
start_button.pack(side = "right" , padx= 10)

end_button = tkinter.Button(top_left_frame, overrelief="solid", width=15, text = "취소", command = stop)
end_button.pack(side = "right" , padx= 10)
end_button['state'] = 'disabled'

p_list = ttk.Treeview(bottom_left_frame, selectmode='browse', columns=[1, 2, 3, 4, 5, 6], height=10, show="headings")
p_list.pack(anchor= "nw")

vsb = Scrollbar(bottom_right_frame, orient="vertical", command=p_list.yview)
vsb.pack(side="left", fill = 'y')

hsb = Scrollbar(bottom_left_frame, orient="horizontal", command=p_list.xview)
hsb.pack(side = "bottom", fill="x")

p_list.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

p_list.heading(1, text="지역")
p_list.heading(2, text="행정구")
p_list.heading(3, text="업체명")
p_list.heading(4, text="업종")
p_list.heading(5, text="전화번호")
p_list.heading(6, text="주소")

p_list.column(1, width=100)
p_list.column(2, width=100)
p_list.column(3, width=200)
p_list.column(4, width=100)
p_list.column(5, width=150)
p_list.column(6, width=200)

p_var = DoubleVar()

log_box = scrolledtext.ScrolledText(bottom_right_frame, height=19, width=48)
log_box.pack(side="bottom", padx = 2)
log_box.insert(END, "Log\n")
log_box.insert(END, "-"*48)
log_box.configure(state='disabled')

window.protocol("WM_DELETE_WINDOW", on_closing)
window.mainloop()
