# -*- coding: utf-8 -*-
import nfc
import binascii
import csv
import sys
import time
import pygame.mixer
import json
import codecs
import PlayVoice
import datetime
userData = {}
nowDataList = []


CSV_PATH1 = "./csv/"
CSV_PATH2 = ".csv"
CSV_NAMELIST_PATH = "./csv/nameList.csv"
VOICE_PATH1 = "./voice/"
VOICE_PATH2 = ".wav"


def clockIn():
    now = datetime.datetime.now()
    writeList = []
    writeList.append(now.strftime("%Y-%m-%d %A"))
    writeList.append(now.strftime("%X"))
    writeList.append(now.strftime("%s"))
    print("書き込み予定データ" + str(writeList))
    try:
        with open(CSV_PATH1 + nowDataList[2] + CSV_PATH2, 'a') as f:
            writer = csv.writer(f, lineterminator='\n')  # 改行コード（\n）を指定しておく
            writer.writerow(writeList)  # list（1次元配列）の場合
            print("成功")
        PlayVoice.playPathById(nowDataList[2])
        PlayVoice.playMorning()
    except:
        PlayVoice.playError_DataError()
        print("書き込み失敗")


def clockOut(list):
    now = datetime.datetime.now()
    delta = long(now.strftime("%s"))
    tmp = list[len(list) - 1]
    before = long(tmp[2])
    tmp[2] = now.strftime("%X")
    tmp.append(str((delta - before) / 60.0))
    print("書き込み予定データ" + str(tmp))
    try:
        with open(CSV_PATH1 + nowDataList[2] + CSV_PATH2, 'w') as f:
            writer = csv.writer(f, lineterminator='\n')  # 改行コード（\n）を指定しておく
            for i in range(len(list)):
                writer.writerow(list[i-1])
        print("書き込み正常終了")
        print(nowDataList[2])
        PlayVoice.playPathById(nowDataList[2])
        PlayVoice.playEvening()
    except:
        print("書き込み失敗")
        PlayVoice.playError_DataError()


def operateCsv():
    print("")
    print("タイムカードを読み込んでいます。")
    path = CSV_PATH1 + nowDataList[2] + CSV_PATH2
    print("書き込みパス: " + str(path))
    workingList = []
    factor_num = 0
    try:
        with open(path, 'r') as f:
            reader = csv.reader(f)
            # header = reader.next()
            for line in reader:
                workingList.append(line)

        factor_num = len(workingList[len(workingList) - 1])
    except:
        print("何もなし")
        factor_num = 0
    print("データ数: " + str(factor_num))
    if factor_num == 3:  # 退勤処理
        print("")
        print("**************")
        print("   退勤処理")
        print(nowDataList[0])
        print("**************")
        print("")
        clockOut(workingList)
    elif factor_num == 4 or factor_num == 0:
        print("")
        print("**************")
        print("   出勤処理")
        print(nowDataList[0])
        print("**************")
        print("")
        clockIn()
    else:
        PlayVoice.playError_DataError()
        print("データがおかしい")


def get_keys_from_value(d, val):  # (dictinary, search value)
    return [k for k, v in d.items() if v == val]


def CheckNFC(_idm):
    PlayVoice.playCheck()
    checkFlg = False
    for value in userData.values():
        if _idm in value:
            nowDataList.append(get_keys_from_value(userData, value))  # name
            nowDataList.append(_idm)  # idm
            nowDataList.append(value[_idm])  # path
            checkFlg = True
            break
    if checkFlg:
        print(str(nowDataList[0])+"ありますねえ！")
    else:
        print("ないです")
        PlayVoice.playError_NoName()

    return checkFlg


def readNameLists():
    print("")
    print("")
    print("nameListを読み込んでいます。")
    with open(CSV_NAMELIST_PATH, 'r') as f:
        reader = csv.reader(f)
        header = reader.next()
        for line in reader:
            tmpDict = {}
            print(str(line))
            tmpDict[line[1]] = line[2]
            userData[line[0]] = tmpDict
    print("nameList読み込み終了")
    print("データ数: "+str(len(userData)))
    print("")
    print("")


def connected(tag):

    print('ICカード検知 Type = {}'.format(tag.type))
    idm = binascii.hexlify(tag.identifier).upper()
    print("idm = %s" % idm)
    if CheckNFC(idm):
        operateCsv()


if __name__ == "__main__":
    readNameLists()

    print("")
    print("")
    print("**********************************")
    print("      セキュリティカード待機中")
    print("**********************************")
    print("")
    print("")

    clf = nfc.ContactlessFrontend('usb')
    clf.connect(rdwr={'on-connect': connected})
    clf.close()
    print("待機")
