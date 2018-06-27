# -*- coding: utf-8 -*-
import subprocess
import nfc
import binascii
import datetime
import sys
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import pygame.mixer
import traceback
from time import sleep
import signal
import random
import pywapi
import threading

global delta_t
global PATH_VRX
global PATH_VOICEROID
global scope
global doc_id
global path
global first_flag
first_flag = True
PATH_VRX = r"C:\Voiceroid2\tamiyasu_talk\vrx.exe"
PATH_VOICEROID = r"C:\Voiceroid2\VoiceroidEditor.exe"
scope = ['https://spreadsheets.google.com/feeds']
doc_id = '1X0TbbGm9s6_rQUXAqWxPCZESCqPUo_n6jeJpSK6wClI'
path = os.path.expanduser("secret2.json")#大津留さん仕様

def process(idm): 

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
        startVoiceroid()
        cmd = PATH_VRX + u" ".encode('cp932') +  u"そのカードは登録されていません。再度ご確認ください。".encode('cp932')
        endVoiceroid(cmd)
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
                thread_2 = threading.Thread(target=playEnd())
                thread_2.start()
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
                thread_2 = threading.Thread(target=playStart())
                thread_2.start()
                print("出勤処理")
                writeDate = datetime.datetime.now().strftime('%Y/%m/%d')
                writeName = name #unicodeのまま
                writeStart = datetime.datetime.now().strftime('%H:%M:%S')
                workingsheet.update_cell(lastDate+1, 1, writeDate)
                workingsheet.update_cell(lastDate+1, 2, writeName)
                workingsheet.update_cell(lastDate+1, 3, writeStart)
                

        except ValueError:
            thread_2 = threading.Thread(target=playStart())
            thread_2.start()
            print(" 名前がまだありませんでした。登録します。")
            writeDate = datetime.datetime.now().strftime('%Y/%m/%d')
            writeName = name #unicodeのまま
            writeStart = datetime.datetime.now().strftime('%H:%M:%S')
            workingsheet.update_cell(lastDate+1, 1, writeDate)
            workingsheet.update_cell(lastDate+1, 2, writeName)
            workingsheet.update_cell(lastDate+1, 3, writeStart)

def connected(tag):
    print("受付")
    thread_2 = threading.Thread(target=playCheck())
    thread_2.start()
    print('ICカード検知 Type = {}'.format(tag.type))
    idm = binascii.hexlify(tag.identifier).upper()
    print("Process開始")
    process(idm)    
    return idm

def playStart():
    
    startVoiceroid()
    cmd = PATH_VRX + u" ".encode('cp932') +u"おはようございます。".encode('cp932') +  u" ".encode('cp932') + u"今日の出勤登録が完了しました。今日も一日頑張りましょう！".encode('cp932')
    endVoiceroid(cmd)

def playEnd():
    startVoiceroid()
    cmd = PATH_VRX + u" ".encode('cp932') + u"退勤登録が完了しました。今日も一日お疲れ様でした。".encode('cp932')
    endVoiceroid(cmd)

def playCheck():
    pygame.mixer.init()
    pygame.mixer.music.load('./check.mp3')
    pygame.mixer.music.play(1)
    time.sleep(0)

def playMorning():
    pygame.mixer.init()
    pygame.mixer.music.load('./morning.mp3')
    pygame.mixer.music.play(1)
    time.sleep(0)

def playEvening():
    pygame.mixer.init()
    pygame.mixer.music.load('./evening.mp3')
    pygame.mixer.music.play(1)
    time.sleep(0)

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

def startVoiceroid():
    cmd = r"cmd /c start " + PATH_VOICEROID
    if subprocess.call(cmd)==0:
        print("Voiceroid staring")
    else:
        print("Voiceroid failed")
    
    cmd = r"cmd /c start " + PATH_VRX + ""
    if subprocess.call(cmd)==0:
        print("VRX starting")
    else:
        print("VRX failed")
    sleep(1)

def endVoiceroid(cmd):
    if subprocess.call(cmd)==0:
        print("VRX text speaching successed!")
        print(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S".encode('cp932')))
    else:
        print("VRX text speaching failed...")
    sleep(3)
    os.system("taskkill /im vrx.exe /f")

def voiceroidLunch(Wlist):
    startVoiceroid()
    hour = datetime.datetime.now().strftime(u"%H".encode('cp932'))
    minute = datetime.datetime.now().strftime(u"%M".encode('cp932'))
    LUNCH_WORD = u" みなさん、お昼ですよーー！".encode('cp932')
    times = len(Wlist)
    for i in range(times):
        Wlist[i] = PATH_VRX + u" ".encode('cp932') + Wlist[i].encode('cp932')

    cmdlist =   [PATH_VRX + u" ".encode('cp932') + hour + u"時".encode('cp932') + minute + u"分になりました。".encode('cp932') + LUNCH_WORD,
                PATH_VRX + u" ".encode('cp932') + u"はあーーおなかすきました。もうこんな時間！お昼休憩にしましょうか！".encode('cp932'),
                PATH_VRX + u" ".encode('cp932') + u"お昼です！お昼になりました。ご飯を食べに行きましょう。".encode('cp932'),
                PATH_VRX + u" ".encode('cp932') + u"皆さんお疲れ様です。お昼になりました。ちょっと休憩しませんか？".encode('cp932')]
    newcmd = cmdlist + Wlist
    print(newcmd)   
    newcmd = random.choice(newcmd)
    endVoiceroid(newcmd)

def voiceroidBreaktime(Wlist):
    startVoiceroid()
    hour = datetime.datetime.now().strftime(u"%H".encode('cp932'))
    minute = datetime.datetime.now().strftime(u"%M".encode('cp932'))
    times = len(Wlist)
    for i in range(times):
        Wlist[i] = PATH_VRX + u" ".encode('cp932') + Wlist[i].encode('cp932')
    cmdlist =   [PATH_VRX + u" ".encode('cp932') + hour + u"時".encode('cp932')  + minute + u"分になりました。".encode('cp932') + u"そろそろ休憩しませんか？".encode('cp932') + u"コーヒーでも飲んで休みましょう！".encode('cp932'),
                PATH_VRX + u" ".encode('cp932') + u"みなさんお疲れ様です！休憩の時間になりました。少し立って体を動かしてリフレッシュしましょう！".encode('cp932'),
                PATH_VRX + u" ".encode('cp932') + u"みなさーーん、休憩ですよ！".encode('cp932') + u"コーヒーでもどうですか？".encode('cp932'),
                PATH_VRX + u" ".encode('cp932') + u"はあーー疲れたもん。もうこんな時間！休憩にしましょう！".encode('cp932') + u"コーヒーでもどうですか？".encode('cp932'),
                PATH_VRX + u" ".encode('cp932') + u"休憩にしましょう！休憩！...みなさん休んでもいいんですよ！".encode('cp932'),
                PATH_VRX + u" ".encode('cp932') + u"お疲れ様ーー！そろそろ休憩時間になります。ちょっと休憩しませんか？".encode('cp932')]
    newcmd = cmdlist + Wlist
    print(newcmd)   
    newcmd = random.choice(newcmd)
    endVoiceroid(newcmd)

def voiceroidEndtime(Wlist):
    startVoiceroid()
    hour = datetime.datetime.now().strftime(u"%H".encode('cp932'))
    minute = datetime.datetime.now().strftime(u"%M".encode('cp932'))
    times = len(Wlist)
    for i in range(times):
        Wlist[i] = PATH_VRX + u" ".encode('cp932') + Wlist[i].encode('cp932')
    cmdlist =   [PATH_VRX + u" ".encode('cp932') + hour + u"時".encode('cp932')  + minute + u"分になりました。".encode('cp932') + u"まもなく終業時間になります。".encode('cp932') + u"お疲れさまでした。".encode('cp932'),
                PATH_VRX + u" ".encode('cp932') + u"まもなく".encode('cp932')+ hour + u"時".encode('cp932')  + minute + u"分になります。".encode('cp932') + u"今日も一日お疲れ様でした。".encode('cp932'),
                PATH_VRX + u" ".encode('cp932')  + u"みなさんお疲れ様です。もうすぐお仕事が終わりの時間になります。毎日お仕事お疲れ様です！".encode('cp932'),
                PATH_VRX + u" ".encode('cp932')  + u"ふうーー疲れました。お仕事の終わりの時間です。帰るかたはお気をつけてお帰りください。それではまた明日！".encode('cp932')]
                
    newcmd = cmdlist + Wlist
    print(newcmd)   
    newcmd = random.choice(newcmd)
    endVoiceroid(newcmd)

def voiceroidOthertime(Wlist):
    startVoiceroid()
    times = len(Wlist)
    for i in range(times):
        Wlist[i] = PATH_VRX + u" ".encode('cp932') + Wlist[i].encode('cp932')
    newcmd = Wlist
    print(newcmd)   
    newcmd = random.choice(newcmd)
    endVoiceroid(newcmd)

def checkNFC():
    while True:
        global credentials
        global client
        global gfile
        global datasheet
        global workingsheet
        print("処理開始")   
        
        scope = ['https://spreadsheets.google.com/feeds']
        doc_id = '1X0TbbGm9s6_rQUXAqWxPCZESCqPUo_n6jeJpSK6wClI'
        path = os.path.expanduser("secret2.json")#大津留さん仕様

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

def readConfigs():
    '''
    scope = ['https://spreadsheets.google.com/feeds']
    doc_id = '1X0TbbGm9s6_rQUXAqWxPCZESCqPUo_n6jeJpSK6wClI'
    path = os.path.expanduser("secret2.json")#大津留さん仕様
    '''
    print("Readconfig開始")
    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    client = gspread.authorize(credentials)
    gfile = client.open_by_key(doc_id)
    wordsheet = gfile.worksheet('word')
    print("読み込み終わり")
    time_list=[]
    word_list_lunch=[]
    word_list_break=[]
    word_list_end=[]
    word_list_other=[]
    for i in range(8):
        time_list.append(wordsheet.cell(3, i+1).value)
    for i in range(10):
        word_list_lunch.append(wordsheet.cell(5+i, 1).value)
        while "" in word_list_lunch:
            word_list_lunch.remove("")
    for i in range(10):
        word_list_break.append(wordsheet.cell(5+i, 3).value)
        while "" in word_list_break:
            word_list_break.remove("")
    for i in range(10):
        word_list_end.append(wordsheet.cell(5+i, 5).value)
        while "" in word_list_end:
            word_list_end.remove("")
    for i in range(10):
        word_list_other.append(wordsheet.cell(5+i, 7).value)
        while "" in word_list_other:
            word_list_other.remove("")
    print("Readconfig終わり")
    return time_list, word_list_lunch, word_list_break, word_list_end, word_list_other

def checkTime(send):
    LIST = []
    TIMELIST = []
    global first_flag
    print("Checkしてる")
    LIST = send
    print("Checkおわり")
    TIMELIST = LIST[0]
    print(LIST[0])
    print(TIMELIST[0])
    if first_flag==True:
        first_flag = False
    nowHour = datetime.datetime.now().hour
    nowMinute = datetime.datetime.now().minute
    nowSecond = datetime.datetime.now().second
    if nowMinute % 20 == 0 and nowSecond<10:
        a=1
        
    if nowHour==TIMELIST[0] and nowMinute==TIMELIST[1]:
        
        print("昼食")
        voiceroidLunch(LIST[1])
       
        
    if nowHour==TIMELIST[2] and nowMinute==TIMELIST[3]:

        print("休憩")
        voiceroidBreaktime(LIST[2])
    elif nowHour==TIMELIST[4] and nowMinute==TIMELIST[5]:

        print("退勤時間")
        voiceroidEndtime(LIST[3])
    elif nowHour==TIMELIST[6] and nowMinute==TIMELIST[7]:

        print("お知らせ")
        voiceroidOthertime(LIST[4])

if __name__ == "__main__":
    while True:
        checkNFC()
        #sys.exit()