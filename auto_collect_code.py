#필요 라이브러리 호출
import requests
import pandas as pd
from pandas import DataFrame
from bs4 import BeautifulSoup
import re
from datetime import datetime
import os
import time
import warnings

warnings.filterwarnings(action='ignore') 



#파일 불러오기
file_name = '파일명입력(확장자까지).xlsm' #확장자까지 포함한 파일명 입력
df = pd.read_excel(file_name)
df.columns = ['검색어','제목', '네이버수집일', '링크', '요약내용', '자동수집 사용여부', '주제영역', '언론사', '날짜', '기사사진', '연도', '월', '주차', '중복체크', '관련회사']
df_new = df #유지 보수를 위해 df_new에 저장, 후에 df형태가 달라질 수 있어서..
#딕셔너리 호출, 링크정보와 언론사 정보를 매핑한 파일: link_name.xlsx
link_name = pd.read_excel('link_name.xlsx')
link_name = link_name.set_index('link')
link_name_dict = link_name.to_dict()
link_name_dict = link_name_dict['name']


#시작시간 저장
start = time.time()


for i in range(len(df_new)):
    url = df_new.iloc[i,3]
    data = requests.get(url,verify=False)
    soup = BeautifulSoup(data.text, 'html.parser')
    
    try: image = soup.select_one('meta[property="og:image"]')['content']  #관련사진 찾기
    except: image = '검색불가'  #로직으로 못 찾아내면 검색불가 입력

    try: name = link_name_dict[url.split('/')[2]] #언론사 찾기 1. 저장해놨던 딕셔너리로 찾기
    except: name= '검색불가' #못 찾으면 검색불가 입력
    
    if name =='검색불가': #언론사 찾기 2. 메타데이터 og:site_name으로 찾기
        try: name = soup.select_one('meta[property="og:site_name"]')['content']
        except: name = '검색불가'
        
    if name == '검색불가': #언론사 찾기 3. 메타데이터 copyright로 찾기
        try: name = soup.select_one('meta[name="copyright"]')['content']
        except: name = '검색불가'   
    if name == '검색불가': #언론사 찾기 4. 메타데이터 Copyright(앞 대문자)로 찾기
        try: name = soup.select_one('meta[name="Copyright"]')['content']
        except: name = '검색불가'
    
    #로직으로 찾아낸 값 입력
    df_new.iloc[i,9] = image 
    df_new.iloc[i,7] = name
    #초기화
    image = None
    name = None

    
#종료시간 저장
end = time.time()

print(end - start)

print('초 소요됨')

#'파일명.xlsx' 형태로 저장
new_file_name = file_name.split('.')[0] + '_자동수집완료.xlsx'  
df_new.to_excel(new_file_name)


