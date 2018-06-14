# -*- coding: utf-8 -*-

import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import sys


def main():

    scope = ['https://spreadsheets.google.com/feeds']
    doc_id = '1aQmUX0dYnuxHRgHndEvW_6uHp2ZcqX7b6jbApL8zSDc'
    path = os.path.expanduser("secret.json")
    ID = "011203125A198D0D"

    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    client = gspread.authorize(credentials)
    gfile   = client.open_by_key(doc_id)
    datasheet  = gfile.worksheet('data')
    workingsheet = gfile.worksheet('working')

    workingsheet_list_date = workingsheet.col_values(1)
    workingsheet_list_name = workingsheet.col_values(2)
    global lastDate
    lastDate = ""
    if None in workingsheet_list_date:
        lastDate = workingsheet_list_date.index("")
    else:
        print("WorkingSheetがおかしいです")
    try:
        cell = datasheet.find(ID)
        print("Found something at R%sC%s" % (cell.row, cell.col))
    except:
        print("そのICカードは登録されていません。")


    global name
    name = datasheet.cell(cell.row, cell.col-1).value #unicode

    print(str(name.encode('utf-8'))+"です") #utf-8で表示
    print("書き込み中")

    workingsheet.add_rows(lastDate)
    zikan = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
    workingsheet.update_cell(lastDate+1, 1, zikan)
    workingsheet.update_cell(lastDate+1, 2, name) #unicodeで書き込み

    nameCount = workingsheet_list_name.count(name) #unicodeで検索
    print("要素数は "+str(nameCount))
    if (nameCount % 2) == 0:
        workingsheet.update_cell(lastDate+1, 3, '出勤'.decode('utf-8'))
    else:
        workingsheet.update_cell(lastDate+1, 3, '退勤'.decode('utf-8'))

    print("プログラム終了")


if __name__ == '__main__':
    main()
