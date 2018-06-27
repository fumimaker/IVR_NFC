# -*- coding: utf-8 -*-

import nfc
import binascii
import datetime
import sys
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import time
import pygame.mixer

def process(idm): 
    
    ''' 
    scope = ['https://spreadsheets.google.com/feeds']
    doc_id = '1X0TbbGm9s6_rQUXAqWxPCZESCqPUo_n6jeJpSK6wClI'
    path = os.path.expanduser("secret2.json")

    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    client = gspread.authorize(credentials)
    gfile = client.open_by_key(doc_id)
    datasheet = gfile.worksheet('data')
    workingsheet = gfile.worksheet('working')
    '''
    print("検索開始")

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
                playNameText(idm)
                playEvening()
                print("退勤処理")
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
                playNameText(idm)
                playMorning()
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

def connected(tag):
    print('ICカード検知 Type = {}'.format(tag.type))
    now = datetime.datetime.now()
    idm = binascii.hexlify(tag.identifier).upper()
    print("Process開始")
    process(idm)    
    return idm


def playCheck():
    pygame.mixer.init()
    pygame.mixer.music.load('./check.mp3')
    pygame.mixer.music.play(1)
    time.sleep(1)

def playMorning():
    pygame.mixer.init()
    pygame.mixer.music.load('./morning.mp3')
    pygame.mixer.music.play(1)
    time.sleep(1)

def playEvening():
    pygame.mixer.init()
    pygame.mixer.music.load('./evening.mp3')
    pygame.mixer.music.play(1)
    time.sleep(1)

def playNameText(idm):
    pygame.mixer.init()
    name = "none"
    if idm == "C3D78B5A": #otsuru
        pygame.mixer.music.load('./otsuru.mp3')
        pygame.mixer.music.play(1)
        name = "otsuru"

    elif idm == "01010312D718B809": #鈴木 謙太
        pygame.mixer.music.load('./suzuki.mp3')
        pygame.mixer.music.play(1)
        name = "suzuki"

    elif idm == "E3E88B5A": #斉藤 弘恭
        pygame.mixer.music.load('./saitou.mp3')
        pygame.mixer.music.play(1)
        name = "saitou"

    elif idm == "93D78B5A": #平井 雄一
        pygame.mixer.music.load('./yuukun.mp3')
        pygame.mixer.music.play(1)
        name = "hirai"

    elif idm == "33C78B5A": #田中 修一
        pygame.mixer.music.load('./tanaka.mp3')
        pygame.mixer.music.play(1)
        name = "tanaka"

    elif idm == "B3E8F65B": #前田 じゅんき
        pygame.mixer.music.load('./maeda.mp3')
        pygame.mixer.music.play(1)
        name = "maeda"

    elif idm == "53C78B5A": #小林 隆貴
        pygame.mixer.music.load('./kobayashi.mp3')
        pygame.mixer.music.play(1)
        name = "kobayashi"

    elif idm == "33D5F65B": #堤 修平
        pygame.mixer.music.load('./tsutsumi.mp3')
        pygame.mixer.music.play(1)
        name = "tsutsumi"

    elif idm == "A3D78B5A": #赤木 豪
        pygame.mixer.music.load('./gou.mp3')
        pygame.mixer.music.play(1)
        name = "akagi"

    elif idm == "13E9F65B": #浅川 雄基
        pygame.mixer.music.load('./asakawa.mp3')
        pygame.mixer.music.play(1)
        name = "asakawa"

    elif idm == "011203125A198D0D": #デバッグ用水野 史暁
        pygame.mixer.music.load('./mizuno.mp3')
        pygame.mixer.music.play(1)
        name = "mizuno"

    elif idm == "C36F395C": #デバッグ用水野 史暁
        pygame.mixer.music.load('./mizuno.mp3')
        pygame.mixer.music.play(1)
        name = "mizuno"

    elif idm == "0114D6AB78181719": #デバッグ用水野 史暁デバッグ用
        pygame.mixer.music.load('./mizuno.mp3')
        pygame.mixer.music.play(1)
        name = "mizuno"

    time.sleep(1)
    return name
while True:
    print("処理開始")   
    global scope
    global doc_id
    global path
    global credentials
    global client
    global gfile
    global datasheet
    global workingsheet

    scope = ['https://spreadsheets.google.com/feeds']
    doc_id = '1X0TbbGm9s6_rQUXAqWxPCZESCqPUo_n6jeJpSK6wClI'
    path = os.path.expanduser("secret2.json")

    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    client = gspread.authorize(credentials)
    gfile = client.open_by_key(doc_id)
    datasheet = gfile.worksheet('data')
    workingsheet = gfile.worksheet('working')
    
    print('****************************')
    print('ICカードをタッチしてください。')
    print('****************************')
    clf = nfc.ContactlessFrontend('usb')
    clf.connect(rdwr={'on-connect': connected})  # now touch a tag
    clf.close()
    #sys.exit()