# -*- coding: utf-8 -*-

import nfc
import binascii
import datetime
import sys
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

def process(idm):
    print(" 処理開始")
    scope = ['https://spreadsheets.google.com/feeds']
    #doc_id = '1aQmUX0dYnuxHRgHndEvW_6uHp2ZcqX7b6jbApL8zSDc'
    doc_id = '1X0TbbGm9s6_rQUXAqWxPCZESCqPUo_n6jeJpSK6wClI'
    path = os.path.expanduser("secret.json")

    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    client = gspread.authorize(credentials)
    gfile = client.open_by_key(doc_id)
    datasheet = gfile.worksheet('data')
    workingsheet = gfile.worksheet('working')
    print(" 列取得開始")
    workingsheet_list_date = workingsheet.col_values(1)
    workingsheet_list_name = workingsheet.col_values(2)
    #print(workingsheet_list_date)
    global lastDate
    lastDate = len(workingsheet_list_date)

    '''
    if "" in workingsheet_list_date:
        global lastDate
        lastDate = workingsheet_list_date.index("")#lastDateにはNoneの縦番号が入っている
    else:
        print("WorkingSheetがおかしいです")
    '''

    print(" 検索開始")
    dataMatch = False
    try:
        cell = datasheet.find(idm)
        print("Found something at R%sC%s" % (cell.row, cell.col))
        global name
        name = datasheet.cell(cell.row, cell.col-1).value  # unicode
        dataMatch = True
    except:
        print(" そのICカードは登録されていません。")
        dataMatch = False

    if dataMatch == True:
        print(" "+name.encode('utf-8')+"です")  # utf-8で表示
        print(" 書き込み中")
        workingsheet.add_rows(lastDate)
        zikan = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
        workingsheet.update_cell(lastDate+1, 1, zikan)
        workingsheet.update_cell(lastDate+1, 2, name)  # unicodeで書き込み

        nameCount = workingsheet_list_name.count(name)  # unicodeで検索
        print(" 要素数は "+str(nameCount))
        if (nameCount % 2) == 0:
            workingsheet.update_cell(lastDate+1, 3, '出勤'.decode('utf-8'))
        else:
            workingsheet.update_cell(lastDate+1, 3, '退勤'.decode('utf-8'))        

    print(' プログラム終了')



def connected(tag):
    print('ICカード検知 Type = {}'.format(tag.type))
    now = datetime.datetime.now()
    idm = binascii.hexlify(tag.identifier).upper()
    print("Process開始")
    process(idm)
    return idm


while True:
    print('****************************')
    print('ICカードをタッチしてください。')
    print('****************************')
    clf = nfc.ContactlessFrontend('usb')
    clf.connect(rdwr={'on-connect': connected})  # now touch a tag
    clf.close()
    #sys.exit()
