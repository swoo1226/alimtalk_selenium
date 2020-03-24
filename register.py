from bs4 import BeautifulSoup
import urllib.request
import urllib.parse
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import time
import sys
import os
import openpyxl
from selenium.webdriver.common.alert import Alert

options = Options()
options.headless = True
browser = webdriver.Chrome(
    executable_path="./chromedriver", options=options)
# browser.get('https://test.pocketsurvey.co.kr/login')

# print(sys.argv)
# # 1.엠앤와이즈 로그인 페이지 이동
path = "./chromedriver"
# driver = webdriver.Chrome(path)
browser.get('https://alimtalk.carrym.com/Gate/')
time.sleep(1)
channel_id_input = browser.find_element_by_id("textfield-1013-inputEl")
print(channel_id_input)
channel_id_input.send_keys('pocketsurvey')
id_input = browser.find_element_by_id("textfield-1014-inputEl")
id_input.send_keys('admin')
pw_input = browser.find_element_by_id("textfield-1016-inputEl")
pw_input.send_keys("At@pizza1!")
login_btn = browser.find_element_by_id("button-1021")

# # 2.마이페이지 이동
login_btn.click()
time.sleep(1)

# # 3.알림톡 템플릿 관리 페이지 이동
manage_template = browser.find_element_by_id("ext-element-37")
manage_template.click()
time.sleep(1)

# # 4.엑셀 업로드 팝업
# # 팝업창 열기
upload_xls = browser.find_element_by_id("button-1092")
upload_xls.click()

# 업로드 파일 가공
filename = sys.argv[1] + ".xlsx"

# 엑셀파일 열기
book = openpyxl.load_workbook(filename)

# 맨 앞의 시트 추출하기
sheet = book.worksheets[0]

# 시트의 각 행을 순서대로 추출하기
data = []
for row in sheet.rows:
    # data.append([row[0].value, row[8].value])
    if(row[3].value == "BA" and row[4].value != None):
        print('버튼 값 입력할 곳')
        row[8].value = '[{{"name": "참여하기","type": "WL", "url_mobile": "http://plus.kakao.com/talk/bot/@#{{{}}}/참여하기#{{참여코드}}"}}]'.format(
            row[0].value)
        data.append([row[0].value, row[8].value])
print('data?', data)
book.save(filename)
# 버튼 json 작성하기
absolute_file_path = os.path.abspath(filename)
print(filename)
choose_file = browser.find_element_by_id("filefield-1100-button-fileInputEl")
choose_file.send_keys(absolute_file_path)
time.sleep(3)
upload = browser.find_element_by_id("button-1101")
upload.click()
time.sleep(3)
confirm = browser.find_element_by_id("button-1005")
# confirm.click()
browser.close()
