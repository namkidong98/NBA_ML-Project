# %%
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import requests                  #파이썬과 URL이 소통하는 창구
from bs4 import BeautifulSoup    #HTML 객체를 parsing하기 위한 라이브러리
import time 
import openpyxl

# %%
def crawl_info(num_of_players):
    for i in range(1, num_of_players + 1): # 1부터 해당 페이지에서 가져올 선수의 숫자만큼 --> row에 해당
        player_data = [] # 해당 선수의 column별 데이터를 저장할 빈 리스트를 선언
        
        for j in range(1, 19): # 1부터 18까지 --> column에 해당
            # 각 선수당 column의 개수만큼 데이터를 모은다
            try:
                data = driver.find_element(By.CSS_SELECTOR, f"div.Crom_container__C45Ti.crom-container > table > tbody > tr:nth-child({i}) > td:nth-child({j})").text 
            except:
                data = "없음"
            
            player_data.append(data) # 수집한 데이터를 차례대로 리스트에 추가하고
            
        # 한 선수에 대한 데이터를 다 모았으면 다음 선수 넘어가기 전에 
        ws.append(player_data) # 해당 데이터를 엑셀 데이터에 쌓기
        time.sleep(1)    

# %%
options = ["2-4+Feet+-+Tight", "4-6+Feet+-+Open", "6%2B+Feet+-+Wide+Open"]
distance = ["2-4", "4-6", "6+"]
years = ["2022-23", "2021-22", "2020-21", "2019-20", "2018-19", "2017-18", "2016-17", "2015-16", "2014-15", "2013-14"]

for year in years: # 연도 선택
    for idx in range(3): # 옵션의 인덱스 선택
        # 엑셀파일 생성
        wb = openpyxl.Workbook(f"{year}_NBA_Shot_Closest Defender_{distance[idx]} Feet.xlsx")        
        ws = wb.create_sheet("Sheet1")             
        ws.append(['Player','Team','Age','GP','G','Freq%','FGM','FGA','FG%','EFG%','2FG Freq%', '2FGM', '2FGA', '2FG%', '3FG Freq%', '3PM', '3PA', '3P%']) # 첫 번째 줄에 column명을 기입

        # 연도와 옵션을 기준으로 크롤링할 웹 페이지의 url 설정
        url = f"https://www.nba.com/stats/players/shots-closest-defender?CloseDefDistRange={options[idx]}&PerMode=Totals&Season={year}&dir=A&sort=PLAYER_NAME"

        # 웹 크롤링 시작
        driver = webdriver.Chrome()
        driver.implicitly_wait(10)  # 웹페이지 로딩 될때까지 5초는 기다림
        driver.maximize_window()    # 화면 최대화
        driver.get(url)        
        time.sleep(3)
        
        driver.execute_script("window.scrollTo(0, 300)") # 스크롤을 next page button이 보일때까지 내리도록
        time.sleep(2)
        next_page = driver.find_element(By.XPATH, "//button[@title='Next Page Button']") # 다음 페이지 버튼을 할당
        num_data = int(list(driver.find_element(By.CSS_SELECTOR, "div.Pagination_content__f2at7.Crom_cromSetting__Tqtiq > div:nth-child(1)").text.split())[0]) # 전체 데이터 개수를 가져오고
        print(num_data) # 총 데이터의 개수를 출력
        num_page = (num_data // 50) + 1 # 한 페이지에 50개씩 데이터가 있다

        for cur_page in range(num_page):
            one_page_num_data = 50
            if cur_page == num_page - 1: # 마지막 페이지이면
                one_page_num_data = num_data % 50 # 마지막 페이지의 데이터 개수로 바꿈
    
            crawl_info(one_page_num_data) # 해당 데이터 개수만큼 선수 데이터를 긁어옴
    
            if cur_page != num_page-1: # 마지막 페이지 전까지만 다음 페이지로 넘기기
                time.sleep(3)
                next_page.click()
            time.sleep(3)
                     
        driver.quit()
        
        wb.save(f"{year}_NBA_Shot_Closest Defender_{distance[idx]} Feet.xlsx") # 상단에서 만든 엑셀 파일명과 동일하게 해서 저장
        
        print(f"{year}_NBA_Shot_Closest Defender_{distance[idx]} Feet.xlsx", "is created") # 엑셀 파일이 만들어지고 있는지 체크
        time.sleep(5) # 다음 driver를 열기 전까지의 시간 여유를 제공


