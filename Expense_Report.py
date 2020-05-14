import os
import openpyxl
import subprocess
from win32com import client
from datetime import datetime


'''
작성 데이터 가이드 표 
B10 '19/09/10 - 19/09/12' 프로젝트 날짜
B14 '홍성규' 투자자
B22 ~ B28 지출 날짜 2019/09/10
C22 ~ C28 카테고리 DANCE (대문자)
D22 ~ D28 내역 dancer(10%)
F22 ~ F28 노트
G22 ~ G28 금액 10000
F30 합계 금액
F34 작성날짜 2019 - 09 - 10
'''


THIS_YEAR = int(datetime.today().year)
THIS_MONTH = int(datetime.today().month)
THIS_DAY = int(datetime.today().day)
if THIS_MONTH < 10:
    THIS_MONTH = '0'+str(THIS_MONTH)
if THIS_DAY < 10:
    THIS_DAY = '0'+str(THIS_DAY)
PROJECT_DATE = '19/08/10 - 19/09/15'
WRITE_DATE = '%d - %s - %s' % (THIS_YEAR, str(THIS_MONTH), str(THIS_DAY))

data_wb = openpyxl.load_workbook('./data/expense_book.xlsx')
data_ws = data_wb.active
table = dict()
# {name: [ [cd, cost, date, category, description, notes, name], [cd, cost, date, category, description, notes, name] ]}
r_i = True
for r in data_ws.rows:
    if r_i is True:
        r_i = False
        continue
    if r[1].value.strip() not in table:
        table[r[1].value.strip()] = [[r[0].value, r[2].value, r[3].value, r[4].value, r[5].value, r[6].value, r[1].value]]
        continue
    if r[1].value.strip() in table:
        table[r[1].value.strip()].append([r[0].value, r[2].value, r[3].value, r[4].value, r[5].value, r[6].value, r[1].value])
        continue
data_wb.close()
# 데이터 읽어 오기

xlApp = client.Dispatch("Excel.Application")
for data_t in table:
    print_wb = openpyxl.load_workbook('./data/print_style.xlsx')
    print_ws = print_wb.active
    print_ws['B10'] = PROJECT_DATE
    print_ws['B14'] = data_t
    cost = 0

    sheet_set = [['B22', 'C22', 'D22', 'F22', 'G22'], ['B23', 'C23', 'D23', 'F23', 'G23'], ['B24', 'C24', 'D24', 'F24', 'G24'],
                 ['B25', 'C25', 'D25', 'F25', 'G25'], ['B26', 'C26', 'D26', 'F26', 'G26'], ['B27', 'C27', 'D27', 'F27', 'G27'],
                 ['B28', 'C28', 'D28', 'F28', 'G28']]
    idt = 0
    for dt in table[data_t]:
        print_ws[sheet_set[idt][0]] = '%d/%d/%d' % (dt[2].year, dt[2].month, dt[2].day)
        print_ws[sheet_set[idt][1]] = dt[3]
        print_ws[sheet_set[idt][2]] = dt[4]
        print_ws[sheet_set[idt][3]] = dt[5]
        print_ws[sheet_set[idt][4]] = dt[1]
        cost = dt[1] + cost
        idt += 1
    print_ws['F30'] = cost
    print_ws['F34'] = WRITE_DATE

    print_wb.save('./result_xlsx/%s.xlsx' % data_t)
    print_wb.close()
    books = xlApp.Workbooks.Open(os.path.abspath(r'C:\Users\JIN\PycharmProjects\YEH\result_xlsx\%s.xlsx' % data_t))
    ws = books.Worksheets[0]
    ws.Visible = 1
    ws.ExportAsFixedFormat(0, os.path.abspath(r'C:\Users\JIN\PycharmProjects\YEH\result_pdf\%s.pdf' % data_t), 0, True)
    books.Close(True)

    passwd = table[data_t][0][0]
    pdf_path = r'C:\Users\JIN\PycharmProjects\YEH\result_pdf\%s.pdf' % data_t
    zip_path = r'C:\Users\JIN\PycharmProjects\YEH\result_zip\%s.zip' % data_t
    cmd = ['7z', 'a', '-p%s' % passwd, zip_path, pdf_path]
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, executable=r'C:\Program Files\7-Zip\7z.exe')
    p.wait()
