import pandas as pd
import openpyxl
legaltags = ('health', 'finance', 'family loss', 'relationships',
             'personal', 'pets', 'assault', 'harassment', 'anxiety')


def checkIf(tag):
    if tag == 'NaN':
        return False
    flag = 0
    for checkForTag in legaltags:
        checkForTag = str(checkForTag).lower().strip()
        tag = str(tag).lower().strip()
        if checkForTag == tag:
            flag = 1
    if flag is 0:
        return False
    return True


data = pd.read_excel(
    r'/media/mayank/10AD068510AD0685/DepressionTagging/analyse.xlsx')
srcfile = openpyxl.load_workbook(
    '/media/mayank/10AD068510AD0685/DepressionTagging/analyse.xlsx', read_only=False, keep_vba=True)
sheet = srcfile.active
tags = pd.DataFrame(data, columns=['Tags'])
taglist = list(tags['Tags'])
c = 2
gender = ""
answer = ""
for x in taglist:
    if not checkIf(x):
        print(x," ",sheet.cell(row=c, column=3).value)
        print(sheet.cell(row=c, column=1).value)
        print("Enter Gender.")
        gender = input()
        gender = gender.strip()
        if gender == 'exit':
            break
        print("Enter tag.")
        answer = input()
        sheet.cell(row=c, column=2).value = gender
        sheet.cell(row=c, column=3).value = answer
    c += 1
srcfile.save('/media/mayank/10AD068510AD0685/DepressionTagging/analyse.xlsx')