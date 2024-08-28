import os
import re
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = '时差表格'

def replace_multiple_spaces(t):
    return re.sub(r'\s+', ' ', t).strip()

path = '0025-6153-MAYDICA-1-1000.txt'
CR_year = list()
PY_year = 0
TI_text = ''
TI_text_combine = ''
PT_text = ''
mean = 'NULL'

difference = list()
circles = 0

data = [
    ["TI", "PT", "PY", "MEAN"]
]



with open(path, 'r') as f:
    file = f.readlines()

length = len(file)
i = 0

while (i >= 0) and (i < length):
    text = file[i]
    stext = text.split()
    if len(stext) != 0:
        texthead = stext[0]
    if texthead == 'TI':
        for n in range(i, length):
            TI_text = file[n]
            TI_text = TI_text.split(" ", 1)
            print(TI_text[0])
            print('*')
            if (TI_text[0] != 'TI') and (TI_text[0] != ''):
                break
            addtext = replace_multiple_spaces(TI_text[1])
            TI_text_combine = TI_text_combine + " " +addtext
    if texthead == 'PT':
        text = text.split(" ", 1)
        PT_text = text[1]
    if texthead == 'CR':
        for k in range(i, length):
            CR_text = file[k]
            CR_text = CR_text.split()
            if CR_text[0] == 'NR':
                break
            for j in CR_text:
                j = j.replace(',', '')
                if j.isdigit():
                    CR_year.append(int(j))
                    break
    if texthead == 'PY':
        PY_year = int(stext[1])
    if len(stext) == 0:
        circles = circles + 1
        for m in CR_year:
            d = PY_year - m
            difference.append(d)
        if len(difference) != 0:
            mean = sum(difference) / len(difference)
        print(mean)
        data.append([TI_text_combine, PT_text, PY_year, mean])


        CR_year = list()
        PY_year = 0
        TI_text = ''
        TI_text_combine = ''
        PT_text = ''
        difference = list()
        mean = "NULL"
    i = i + 1


for row in data:
    ws.append(row)

wb.save("output.xlsx")
"""
print(CR_year)
print(circles)
print(file)
print(type(file))
print(data)
"""
print(circles)