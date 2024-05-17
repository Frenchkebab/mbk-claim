# input파일 처리 라이브러리
import columns
import pandas as pd
import os
import openpyxl
from email import policy
from email.parser import BytesParser
from glob import glob
import base64

# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.ui import WebDriverWait as wait
# import datetime
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import Select
# from selenium.webdriver.common.by import By

def attributeWrite(file_name):
    
    wb = openpyxl.load_workbook(f"./upload/{file_name}")

    sheet = wb.worksheets[0]

    for i in range (1, 25):
        sheet.cell(row = 5, column = i).value = columns.column[f"{i}"]

    wb.save(f"./upload/{file_name}")
    wb.close()

# 엑셀 데이터 입력
def userInput():
    f = open('./settings/setting.txt', 'r', encoding='utf-8')
    id = f.readline().strip()
    password = f.readline().strip()
    f.readline()
    work_type = f.readline().split(':')[1].strip()              # 작업 종류
    file_name = f.readline().split(':')[1].strip()              # 파일 이름
    start_idx = int(f.readline().split(':')[1].strip()) - 1     # 시작 인덱스
    last_idx = int(f.readline().split(':')[1].strip()) - 1      # 마지막 인덱스
    minSecond = int(f.readline().split(':')[1].strip())         # 최소 시간
    maxSecond = int(f.readline().split(':')[1].strip())         # 최대 시간
    f.close()
    
    # 변수 여러개 묶어서 리턴
    user_input = [id, password, work_type, file_name, start_idx, last_idx, minSecond, maxSecond]
    return user_input

def printUserInput(id, password, work_type, file_name, start_idx, last_idx, minSecond, maxSecond):
    print(f"아이디 : {id}")
    print(f"패스워드 : {password}")
    print(f"작업 종류: {work_type}")
    print(f"파일 이름: {file_name}")
    print(f"시작 No.: {start_idx + 1}")
    print(f"마지막 No.: {last_idx + 1}")
    print(f"최소 대기시간(초) : {minSecond}")
    print(f"최대 대기시간(초) : {maxSecond}")

def readXslx(file_name):
    df = pd.read_excel(f'./upload/{file_name}',
                                        skiprows=4,
                                        na_values='',
                                        dtype = {
                                            "No.": str,
                                            "Commission No.": str,
                                            "VIN No.": str,
                                            "Repair No.": str,
                                            "B/L no.": str,
                                            "Closing Date": str,
                                            "Incident date": str,
                                            "Damage Code01": str,
                                            "Damage Code02": str,
                                            "Damage Code03": str,
                                            "Damage Code04": str,
                                            "Damage Code05": str,
                                            "Damage Code06": str,
                                            "Damage Code07": str,
                                            "Damage Code08": str,
                                            "Damage Code09": str,
                                            "Damage Code010": str,
                                            "Sub Total": str,
                                            "Date of booking": str,
                                            "Reclamation date": str,
                                            "Audit type": str,
                                            "Claim Entry ID-Number": str,
                                            "CID": str
                                        })
    return df

# 데이터프레임의 각 행을 dictionary로 만들고, 각 데이터타입을 정렬
def dfToDictArr(df, start_idx, last_idx):
    dataArr = []
    
    for i in range(start_idx, last_idx + 1):
        rowDict = df.loc[i].to_dict() # 각 행을 딕셔너리로 저장
        rowDict['Closing Date'] = rowDict['Closing Date'][:10]
        rowDict['Incident date'] = rowDict['Incident date'][:10]
        rowDict['Date of booking'] = rowDict['Date of booking'][:10]
        rowDict['Reclamation date'] = rowDict['Reclamation date'][:10]
        dataArr.append(rowDict)

    return dataArr  # 데이터 사전을 원소로 갖는 리스트

def searchFileName(dirName, row):
    currentAbsPath = os.path.dirname(os.path.realpath(__file__))
    dirAbsPath = currentAbsPath + f"\\upload\\{dirName}"
    fileList = os.listdir(dirAbsPath)

    if dirName == "LIST":
        result = [f"{dirAbsPath}\\{fileList[0]}"]
    else:
        result = []
        for file in fileList:
            if row["VIN No."] in file:
                result.append(f"{dirAbsPath}\\{file}")
        
    return result

def doesBLContainMolu(dirName, row):
    currentAbsPath = os.path.dirname(os.path.realpath(__file__))
    dirAbsPath = currentAbsPath + f"\\upload\\{dirName}"
    fileList = os.listdir(dirAbsPath)

    # Example:
    # 'W1N9M0JB1PN028391_MOLU18004491675' -> returns True
    for file in fileList:
        if file.startswith(row["VIN No."]):
            fileNameStrings = file.split("_")
            return fileNameStrings[1].startswith("MOLU")

def isHyundaiGlovis(dirName, row):
    currentAbsPath = os.path.dirname(os.path.realpath(__file__))
    dirAbsPath = currentAbsPath + f"\\upload\\{dirName}"
    fileList = os.listdir(dirAbsPath)

    # Example:
    # 'W1K1K5KB3PF200197_HDGLMXKR0523894A' -> returns True
    for file in fileList:
        if file.startswith(row["VIN No."]):
            fileNameStrings = file.split("_")
            return fileNameStrings[1].startswith("HDGLMXKR")
    


import re 

def searchEmail(row):
    # 해당 VIN No.와 동일한 EMAIL파일의 경로를 찾는다.
    currentAbsPath = os.path.dirname(os.path.realpath(__file__))
    file_list = list(glob(f"{currentAbsPath}\\upload\\EMAIL\\*.eml"))

    result = []

    for file in file_list:
        if row["VIN No."] in file:
            result.append(file)

        else:
            with open(file, 'rb') as fp:
                msg = BytesParser(policy=policy.default).parse(fp)
                txt = ""
                
                # 이메일 제목 추출
                subject = msg['subject']
                
                # 제목에 VIN No. 포함 여부 확인
                if subject.find(row["VIN No."]) > -1:
                    print(f'제목에 VIN No. 포함 ({row["VIN No."]}):',  subject)
                    result.append(file)
                    continue
                
                # 본문 텍스트 추출
                try: 
                    txt = msg.get_body(preferencelist=('plain')).as_string()
                except:
                    txt = msg.get_body(preferencelist=('plain', 'html')).get_content()
                    
                # 헤더 부분 제거 (정규 표현식 사용)
                pattern = re.compile(r'^Content-Type:.*?\n^Content-Transfer-Encoding:.*?\n\n', re.DOTALL | re.MULTILINE)
                base64_encoded_str = pattern.sub('', txt).strip()

                # 여러 줄로 되어 있는 Base64 인코딩 본문을 한 줄로 합치기
                base64_encoded_str = ''.join(base64_encoded_str.splitlines())
                
                # Base64 디코딩
                try:
                    decoded_bytes = base64.b64decode(base64_encoded_str)
                    decoded_str = decoded_bytes.decode('utf-8')  # 또는 'ascii'로 디코딩 가능
                except (base64.binascii.Error, UnicodeDecodeError) as e:
                    print("디코딩 중 오류 발생:", str(e))
                    print("바이너리 데이터로 처리해야 할 가능성이 있습니다.")

                # 본문에 VIN No. 포함 여부 확인
                if decoded_str.find(row["VIN No."]) > -1:
                    print(f'본문에 VIN No. 포함 ({row["VIN No."]}): ', decoded_str)
                    result.append(file)

    return result