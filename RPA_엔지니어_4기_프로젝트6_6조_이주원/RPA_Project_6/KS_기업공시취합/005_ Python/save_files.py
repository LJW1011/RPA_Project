import pandas as pd
import win32com.client as win32
import os
import re
from PyPDF2 import PdfMerger
from winreg import *
import tkinter.messagebox as msbox


def save_file(path):

    # 한글 보안모듈 레지스트리 등록
    hwp_auto_path = r"Software\Hnc\HwpAutomation\Modules"
    reg_handle = ConnectRegistry(None, HKEY_CURRENT_USER)
    try:
        print("한글 보안모듈 레지스트리 등록")
        key = OpenKey(reg_handle, hwp_auto_path, 0, KEY_WRITE)
        SetValueEx(key, "FilePathCheckerModule", 0, REG_SZ,
            path+r"\005_ Python\FilePathCheckerModuleExample.dll")
    except FileNotFoundError:
        print("AutoConfigURL not found")

    # 한글파일열기
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
    print("템플릿파일 열기")
    hwp.Open(path+r"\001_Template\회사개요서.hwp")

    # 누름틀 필드 가져오기
    field = hwp.GetFieldList().split("\x02")

    # 엑셀파일 읽기
    print("엑셀파일 읽는중")
    dir = path + "\\003_Result\\"
    file = "신규상장.xlsx"
    df = pd.read_excel(dir+file, engine='openpyxl', sheet_name=None)
    df = pd.concat(df, ignore_index=True)
    df_2 = df.rename(columns={"회사명": "기업체명", "매출액(수익) 추출": "매출액"})

    # 누름틀에 맞는 데이터 가져오기
    for i in range(len(df_2)):
        data = []
        num = 0
        for j in field:
            data.append(df_2[j][i])

        # 한글파일 작성
        for j in data:
            hwp.PutFieldText(field[num], j)
            num += 1

        # 다른이름으로 저장
        filename = data[field.index('기업체명')]
        file_path = dir + filename
        hwp.SaveAs(file_path+".hwp")
        hwp.SaveAs(file_path+".pdf", "PDF")
        print(filename+"저장")

    # 한글 종료
    hwp.Quit()

    # pdf 병합
    pdfs = [f for f in os.listdir(dir) if re.match('.*[.]pdf', f)]
    merger = PdfMerger()

    for pdf in pdfs:
        merger.append(dir+"/"+pdf)

    merger.write(dir+"/RESULT.pdf")
    print("RESULT.pdf 파일 생성")
