from selenium import webdriver
from selenium.webdriver.common.by import By
# from timer import *
import openpyxl
import time
import pyautogui

def altTab():
    pyautogui.keyDown('alt')
    time.sleep(.2)
    pyautogui.press('tab')
    time.sleep(.2)
    pyautogui.keyUp('alt')

def toClaimX(driver):
    driver.get("https://www.claimx.de/claimx")
    driver.implicitly_wait(30)
    

def login(driver, id, password):
    # id 입력
    idForm = driver.find_element(by=By.XPATH, value="/html/body/table/tbody/tr[3]/td[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/form/table/tbody/tr[1]/td[2]/input")
    idForm.send_keys(id)
    time.sleep(0.5)

    # password 입력
    pwForm = driver.find_element(by=By.XPATH, value="/html/body/table/tbody/tr[3]/td[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/form/table/tbody/tr[2]/td[2]/input")
    pwForm.send_keys(password)

    ## clipboard 사용하여 입력
    # clipboard.copy(id)
    # idForm.click()
    # pyautogui.hotkey('ctrl', 'v')
    # time.sleep(0.5)

    # clipboard.copy(password)
    # pwForm.click()
    # pyautogui.hotkey('ctrl', 'v')
    
    print(f"'{id}'")
    print(f"'{password}'")
    input("press Enter: ")
    pwForm.click()
    pwForm.send_keys(Keys.ENTER)
    # 로그인 버튼 클릭
    # driver.find_element(by=By.XPATH, value="/html/body/table/tbody/tr[3]/td[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/form/table/tbody/tr[3]/td/input").click()

    # 문 버튼이 나타나면 클릭, 없으면 그냥 패스
    try:
        driver.find_element(by=By.XPATH, value="/html/body/table/tbody/tr[3]/td[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/a[1]/img").click()

    except:
        pass

def clickClaim(driver):
    # claim 버튼 클릭
    driver.find_element(by=By.XPATH, value='/html/body/table/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/table[1]/tbody/tr/td[1]/table/tbody/tr/td[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/a').click()

def memo(file_name, row, msg):
    # 엑셀 파일 오픈
    wb = openpyxl.load_workbook(f"./upload/{file_name}")

    # 시트 설정
    sheet = wb.worksheets[0]

    if row == 0:
        sheet.cell(row = 5, column = 24).value = msg
    else:
        # cid값 저장
        sheet.cell(row = 5 + int(row["No."]), column = 24).value = msg

    # 파일 저장 후 닫기
    wb.save(f"./upload/{file_name}")
    wb.close()

def writeLog(logFile, msg):
    logFile.write(f"{msg}\n")

from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from random import random, randint
import time
import openpyxl

import pyautogui
import clipboard

# from inputFunctions import *
# from common import *
# from timer import *

# cid가 이미 입력되었는지 검사
def checkCID(row):
    # CID가 입력되지 않았으면 
    if str(row["CID"]) == "nan":
        return False

    # 이미 CID가 입력되어 있으면
    else:
        return True


# Vehicle Logistics 입력한 경우 query로 찾아가기
def query(driver, row):
    driver.implicitly_wait(3)
    driver.find_element(by=By.LINK_TEXT, value="query").click() # query 버튼 클릭
    waitLoading()

    vinForm = driver.find_element(by=By.NAME, value="field_akrefitem")
    vinForm.clear()
    time.sleep(0.5)

    cidForm = driver.find_element(by=By.NAME, value="field_aksiditem")
    cidForm.clear()
    time.sleep(1)

    # 만약 경고창이 뜨는 경우
    # try:
    #     driver.find_element(by=By.XPATH, value="/html/body/div[2]/div[1]/button/span[1]").click()
    # except:
    #     pass

    cidForm.send_keys(str(row["CID"]))
    time.sleep(1)
    cidForm.send_keys(Keys.ENTER)

    driver.implicitly_wait(30)
    tdTag = driver.find_element(by=By.XPATH, value='//*[@id="show_header_reference"]/table/tbody/tr/td/b')
    tdTag.find_element(by=By.TAG_NAME, value='')
    # uploadBtn.click()

    ### Vehicle Logistics에서 submit한 후의 화면이 나타난다!


def getCid(file_name, driver, row):
    # 텍스트 클릭
    line = driver.find_element(by=By.XPATH, value='//*[@id="show_header_reference"]/table/tbody/tr/td/b').get_attribute('innerText')
    line = line.strip()

    # cid 값 변수 저장
    cid = line[-7:]

    # 엑셀 파일 오픈
    wb = openpyxl.load_workbook(f"./upload/{file_name}")

    # 시트 설정
    sheet = wb.worksheets[0]

    # cid값 저장
    sheet.cell(row = 5 + int(row["No."]), column = 23).value = cid

    # 파일 저장 후 닫기
    wb.save(f"./upload/{file_name}")
    wb.close()

def uploadArchive(driver,logFile, fileList, selectionList):
    div = 1
    for i in range(0, 7):
        for file in fileList[i]:
            writeLog(logFile, file)

            # file selection 클릭
            driver.implicitly_wait(5)
            fileSelection = driver.find_element(by=By.XPATH, value='//*[@id="mainpart"]/table[5]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr/td/img')
            print('Button found')
            fileSelection.click()
            time.sleep(2)

            # 파일 선택
            pyautogui.write(file)
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(2)

            # document key 선택
            documentKey = driver.find_element(by=By.XPATH, value=f'//*[@id="previews"]/div[{div}]/div[2]/div[1]/div[2]/div/button')
            div += 1
            action = ActionChains(driver)
            action.move_to_element(documentKey).perform()
            time.sleep(0.5)
            documentKey.click()
            time.sleep(1)

            # document key 입력 후 선택
            clipboard.copy(selectionList[i])
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pyautogui.press("enter")
            waitLoading()


    driver.find_element(by=By.XPATH, value='//*[@id="actions"]/div[1]/button[1]').click()

    driver.implicitly_wait(5)
    
    # 업로드 화면이 닫혔는지 확인할 때까지 
    while(True):
        time.sleep(2)
        try:
            driver.switch_to.window(driver.window_handles[1])
            continue
        except:
            driver.switch_to.window(driver.window_handles[0])
            waitLoading()
            break
    driver.implicitly_wait(30)