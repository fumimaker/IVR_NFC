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
path = os.path.expanduser("secret2.json")  # 大津留さん仕様


def startVoiceroid():
    cmd = r"cmd /c start " + PATH_VOICEROID
    if subprocess.call(cmd) == 0:
        print("Voiceroid staring")
    else:
        print("Voiceroid failed")

    cmd = r"cmd /c start " + PATH_VRX + ""
    if subprocess.call(cmd) == 0:
        print("VRX starting")
    else:
        print("VRX failed")
    sleep(1)


def endVoiceroid(cmd):
    if subprocess.call(cmd) == 0:
        print("VRX text speaching successed!")
        print(datetime.datetime.now().strftime(
            "%Y/%m/%d %H:%M:%S".encode('cp932')))
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

    cmdlist = [PATH_VRX + u" ".encode('cp932') + hour + u"時".encode('cp932') + minute + u"分になりました。".encode('cp932') + LUNCH_WORD,
               PATH_VRX +
               u" ".encode('cp932') +
               u"はあーーおなかすきました。もうこんな時間！お昼休憩にしましょうか！".encode('cp932'),
               PATH_VRX + u" ".encode('cp932') +
               u"お昼です！お昼になりました。ご飯を食べに行きましょう。".encode('cp932'),
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
    cmdlist = [PATH_VRX + u" ".encode('cp932') + hour + u"時".encode('cp932') + minute + u"分になりました。".encode('cp932') + u"そろそろ休憩しませんか？".encode('cp932') + u"コーヒーでも飲んで休みましょう！".encode('cp932'),
               PATH_VRX +
               u" ".encode(
                   'cp932') + u"みなさんお疲れ様です！休憩の時間になりました。少し立って体を動かしてリフレッシュしましょう！".encode('cp932'),
               PATH_VRX + u" ".encode('cp932') + u"みなさーーん、休憩ですよ！".encode(
                   'cp932') + u"コーヒーでもどうですか？".encode('cp932'),
               PATH_VRX + u" ".encode('cp932') + u"はあーー疲れたもん。もうこんな時間！休憩にしましょう！".encode(
                   'cp932') + u"コーヒーでもどうですか？".encode('cp932'),
               PATH_VRX +
               u" ".encode('cp932') +
               u"休憩にしましょう！休憩！...みなさん休んでもいいんですよ！".encode('cp932'),
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
    cmdlist = [PATH_VRX + u" ".encode('cp932') + hour + u"時".encode('cp932') + minute + u"分になりました。".encode('cp932') + u"まもなく終業時間になります。".encode('cp932') + u"お疲れさまでした。".encode('cp932'),
               PATH_VRX + u" ".encode('cp932') + u"まもなく".encode('cp932') + hour + u"時".encode(
                   'cp932') + minute + u"分になります。".encode('cp932') + u"今日も一日お疲れ様でした。".encode('cp932'),
               PATH_VRX +
               u" ".encode(
                   'cp932') + u"みなさんお疲れ様です。もうすぐお仕事が終わりの時間になります。毎日お仕事お疲れ様です！".encode('cp932'),
               PATH_VRX + u" ".encode('cp932') + u"ふうーー疲れました。お仕事の終わりの時間です。帰るかたはお気をつけてお帰りください。それではまた明日！".encode('cp932')]

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


def readConfigs():
    scope = ['https://spreadsheets.google.com/feeds']
    doc_id = '1X0TbbGm9s6_rQUXAqWxPCZESCqPUo_n6jeJpSK6wClI'
    path = os.path.expanduser("secret2.json")  # 大津留さん仕様
    # print("Readconfig開始")
    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    client = gspread.authorize(credentials)
    gfile = client.open_by_key(doc_id)
    wordsheet = gfile.worksheet('word')
    # print("Wordsheetインスタンス読み込み終わり")
    time_list = []
    word_list_lunch = []
    word_list_break = []
    word_list_end = []
    word_list_other = []
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
    # print("Readconfig終わり")
    return time_list, word_list_lunch, word_list_break, word_list_end, word_list_other


def checkTime(send):
    global sayFlg
    LIST = []
    TIMELIST = []
    LIST = send
    TIMELIST = LIST[0]
    nowHour = datetime.datetime.now().hour
    nowMinute = datetime.datetime.now().minute
    nowSecond = datetime.datetime.now().second

    if nowHour == int(TIMELIST[0]) and nowMinute == int(TIMELIST[1]) and sayFlg:
        print("昼食")
        voiceroidLunch(LIST[1])
        sayFlg = False

    elif nowHour == int(TIMELIST[2]) and nowMinute == int(TIMELIST[3]) and sayFlg:
        print("休憩")
        voiceroidBreaktime(LIST[2])
        sayFlg = False

    elif nowHour == int(TIMELIST[4]) and nowMinute == int(TIMELIST[5]) and sayFlg:
        print("退勤時間")
        voiceroidEndtime(LIST[3])
        sayFlg = False

    elif nowHour == int(TIMELIST[6]) and nowMinute == int(TIMELIST[7]) and sayFlg:
        print("お知らせ")
        voiceroidOthertime(LIST[4])
        sayFlg = False


if __name__ == "__main__":
    sayFlg = True
    print("timeReport開始")
    sendtuple = []
    print("GoogleSpreadSheetのWordデータを読み込みます")
    sendtuple = readConfigs()
    print("時間監視開始")
    counter = 0
    while True:
        checkTime(sendtuple)
        # counter+=1
        # print(counter)
        time.sleep(1)
        if(counter > 1800):
            # print("Wordデータ定期読み込み時間")
            counter = 0
            sendtuple = readConfigs()
            # print("時間監視開始")
            sayFlg = True
