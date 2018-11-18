# python 3

import xlrd
import tkinter as tk
from tkinter import filedialog
import re

TEST_ROW = 22
TEST_COLUM = 6
TEST_VALUE = 'I. ТОВАРНЫЙ РАЗДЕЛ'
#Begin of cells with additional info
INFO_ROW = 28
INFO_COLUMN = 13
CONSIGNEE_ROW= 16
CONSIGNEE_COLUMN = 3
AGREEMENT_ROW = 17
AGREEMENT_COLUMN = 3
TTN_NUMBER_ROW = 3
TTN_NUMBER_COLUMN = 13
MATCH_PATTERN = r'(\d{8})'


###########################################

def CheckCorrectExcelFile(path_to_file):
    """
    Checks if the opened Excel file is a file with TTN

    returns True if correct
    otherwise returns False
    """

    res =False

    try:
        with xlrd.open_workbook(path_to_file,
                                encoding_override='cp1251') as book:
            sheet = book.sheets()[0]
            if sheet.cell(TEST_ROW, TEST_COLUM).value == TEST_VALUE:
                res = True
            else:
                res = False
    except:
        pass
    finally:
        return res
################################################

def getCertificates(path_to_file):
    """
    Returns an dict with additional info about TTN
    cerificates numbers if any found in the TTN file
    """
    cert_array={}

    TTN_info={}

    try:
        with xlrd.open_workbook(path_to_file,
                                encoding_override='cp1251') as book:
            #Only one sheet is supposed to be inside
            sheet = book.sheets()[0]
            #reading other data from the file
            TTN_info["Agreement"] = sheet.cell(AGREEMENT_ROW, AGREEMENT_COLUMN)
            TTN_info["TTN_Number"] = sheet.cell(TTN_NUMBER_ROW, TTN_NUMBER_COLUMN)
            TTN_info["Consignee"] = sheet.cell(CONSIGNEE_ROW, CONSIGNEE_COLUMN)
            TTN_info["Certificates"] = {}
            #reading addtional info containg informatino about certificates
            delta = 0
            while True:
                cell_value = sheet.cell(INFO_ROW + delta, INFO_COLUMN).value
                if cell_value[0] in 'Сс':
                    cert_search_res = re.findall(MATCH_PATTERN, cell_value, re.M)
                    try:
                        #only unique cert numbers needed for the file
                        for cer in cert_search_res:
                            cert_array[cer] = ''
                    except:
                        pass
                    delta += 1
                else:
                    break
    except:
        pass
    finally:
        TTN_info["Certificates"] = cert_array
        return TTN_info

################################################

file_path = filedialog.askopenfilename(
    filetypes=[("Excel files","*.xls")], title="Файлы ТТН", multiple=True)

#Do necessary procedures to every selected file_path
for i in range(len(file_path)):
    #Check if it is a TTN file
    if CheckCorrectExcelFile(file_path[i]):
        res  = getCertificates(file_path[i])
        for j in res["Certificates"].keys():
            print(j)
