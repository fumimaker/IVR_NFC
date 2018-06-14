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
    doc_id = '1aQmUX0dYnuxHRgHndEvW_6uHp2ZcqX7b6jbApL8zSDc'
    path = os.path.expanduser("secret.json")

    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    client = gspread.authorize(credentials)
    gfile = client.open_by_key(doc_id)
    datasheet = gfile.worksheet('data')
    workingsheet = gfile.worksheet('working')

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

    dataMatch = True

    if dataMatch == True:
        print(" "+name.encode('utf-8')+"です")  # utf-8で表示

        workingsheet_list_date = workingsheet.col_values(1)
        workingsheet_list_name = workingsheet.col_values(2)
        global lastDate
        lastDate = len(workingsheet_list_date)
        print("lastDate = "+str(lastDate))
        print(workingsheet_list_name)
        workingsheet_list_name.reverse()
        try:
            namecolnum = workingsheet_list_name.index(name) #unicodeとUnicodeで比較して行番号を入れる
            print(" すでに名前があります。 "+ "下から"+str(namecolnum))
            upnum = lastDate - namecolnum
            if workingsheet.cell(upnum, 4).value == "": #退勤時処理
                print(" 退勤処理")
                dt = datetime.datetime.now()
                writeEnd = dt.strftime('%H:%M:%S')
                t1 = workingsheet.cell(upnum, 3).value
                tmpDate = workingsheet.cell(upnum, 1).value + " " + t1
                time_t1 = datetime.datetime.strptime(tmpDate, "%Y/%m/%d %H:%M:%S")
                time_t2 = dt
                time_t3 = time_t2 - time_t1
                writeDuty = time_t3
                print(writeDuty)
                workingsheet.update_cell(upnum, 4, writeEnd)
                workingsheet.update_cell(upnum, 5, str(writeDuty))
            else: #出勤時処理
                print(" 出勤処理")
                writeDate = datetime.datetime.now().strftime('%Y/%m/%d')
                writeName = name #unicodeのまま
                writeStart = datetime.datetime.now().strftime('%H:%M:%S')
                workingsheet.update_cell(lastDate+1, 1, writeDate)
                workingsheet.update_cell(lastDate+1, 2, writeName)
                workingsheet.update_cell(lastDate+1, 3, writeStart)

        except ValueError:
            print(" 名前がまだありませんでした。登録します。")
            writeDate = datetime.datetime.now().strftime('%Y/%m/%d')
            writeName = name #unicodeのまま
            writeStart = datetime.datetime.now().strftime('%H:%M:%S')
            workingsheet.update_cell(lastDate+1, 1, writeDate)
            workingsheet.update_cell(lastDate+1, 2, writeName)
            workingsheet.update_cell(lastDate+1, 3, writeStart)
'''
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
'''



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