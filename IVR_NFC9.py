# -*- coding: utf-8 -*-
import nfc
import binascii
import csv
import sys
import time
import json
import codecs
import PlayVoice
import datetime
from slacker import Slacker

userData = {}
nowDataList = ["", "", ""]


CSV_PATH1 = "./csv/"
CSV_PATH2 = ".csv"
CSV_NAMELIST_PATH = "./csv/nameList.csv"
VOICE_PATH1 = "./voice/"
VOICE_PATH2 = ".wav"

channelName = "working"


class Slack(object):
    __slacker = None

    def __init__(self, token):
        self.__slacker = Slacker(token)

    def get_channnel_list(self):
        """
        Slackチーム内のチャンネルID、チャンネル名一覧を取得する。
        """
        # bodyで取得することで、[{チャンネル1},{チャンネル2},...,]の形式で取得できる。
        raw_data = self.__slacker.channels.list().body
        result = []
        for data in raw_data["channels"]:
            result.append(
                dict(channel_id=data["id"], channel_name=data["name"]))
        return result

    def post_message_to_channel(self, channel, message):
        """
        Slackチームの任意のチャンネルにメッセージを投稿する。
        """
        channel_name = "#" + channel
        self.__slacker.chat.post_message(channel_name, message)


def clockIn():
    now = datetime.datetime.now()
    writeList = []
    writeList.append(now.strftime("%Y-%m-%d %A"))
    writeList.append(now.strftime("%X"))
    writeList.append(now.strftime("%s"))
    print("[log] "+"書き込み予定データ" + str(writeList))
    try:
        with open(CSV_PATH1 + nowDataList[2] + CSV_PATH2, 'a') as f:
            writer = csv.writer(f, lineterminator='\n')  # 改行コード（\n）を指定しておく
            writer.writerow(writeList)  # list（1次元配列)
            print("[log] " + "成功")
        mozi = str(nowDataList[0])+"さん、おはようございます。正常に出勤処理を行いました。"
        postSlack(mozi)
        PlayVoice.playPathById(nowDataList[2])
        PlayVoice.playMorning()

    except:
        PlayVoice.playError_DataError()
        print("[error] " + "書き込み失敗")
        mozi = "ごめんなさい、書き込みに失敗しました。"
        postSlack(mozi)


def clockOut(list):
    now = datetime.datetime.now()
    delta = long(now.strftime("%s"))
    tmp = list[len(list) - 1]
    before = long(tmp[2])
    tmp[2] = now.strftime("%X")
    tmp.append(str((delta - before) / 60.0))
    print("[log] "+"書き込み予定データ" + str(tmp))
    try:
        with open(CSV_PATH1 + nowDataList[2] + CSV_PATH2, 'w') as f:
            writer = csv.writer(f, lineterminator='\n')  # 改行コード（\n）を指定しておく
            for i in range(len(list)):
                writer.writerow(list[i-1])
        print("[log] " + "書き込み正常終了")
        mozi = str(nowDataList[0])+"さん、お疲れ様でした。正常に退勤処理を行いました。"
        postSlack(mozi)
        print(nowDataList[2])
        PlayVoice.playPathById(nowDataList[2])
        PlayVoice.playEvening()

    except:
        print("[error] " + "書き込み失敗")
        PlayVoice.playError_DataError()
        mozi = "ごめんなさい、書き込みに失敗しました。"
        postSlack(mozi)


def operateCsv():
    print("")
    print("[log] "+"タイムカードを読み込んでいます。")
    path = CSV_PATH1 + nowDataList[2] + CSV_PATH2
    print("[log] "+"書き込みパス: " + str(path))
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
        print("[error] " + "何もなし")
        factor_num = 0
    print("[log] "+"データ数: " + str(factor_num))
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
        print("[error] " + "データがおかしい")
        mozi = "データベースが何かおかしいみたいです。ごめんなさい、記録できませんでした。"
        postSlack(mozi)


def get_keys_from_value(d, val):  # (dictinary, search value)
    return [k for k, v in d.items() if v == val]


def CheckNFC(_idm):
    PlayVoice.playCheck()
    checkFlg = False
    for value in userData.values():
        if _idm in value:
            nowDataList[0] = get_keys_from_value(userData, value)  # name
            nowDataList[1] = _idm  # idm
            nowDataList[2] = value[_idm]  # path
            checkFlg = True
            break
    if checkFlg:
        print("[log] "+str(nowDataList[0])+"名前発見")
    else:
        print("[error] " + "CSV名前なし")
        PlayVoice.playError_NoName()
        mozi = "カードがうまく読み取れませんでした。カードが重なってませんか？ご確認ください。"
        postSlack(mozi)

    return checkFlg


def readNameLists():
    print("")
    print("")
    print("[log] "+"nameListを読み込んでいます。")
    with open(CSV_NAMELIST_PATH, 'r') as f:
        reader = csv.reader(f)
        header = reader.next()
        for line in reader:
            tmpDict = {}
            print(str(line))
            tmpDict[line[1]] = line[2]
            userData[line[0]] = tmpDict
    print("[log] "+"nameList読み込み終了")
    print("[log] "+"データ数: "+str(len(userData)))
    print("")
    print("")


def postSlack(mozi):
    slack.post_message_to_channel(channelName, mozi)


def connected(tag):

    print("[log] "+'ICカード検知 Type = {}'.format(tag.type))
    idm = binascii.hexlify(tag.identifier).upper()
    print("[log] "+"idm = %s" % idm)
    if CheckNFC(idm):
        operateCsv()


if __name__ == "__main__":

    slack = Slack("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
    postSlack("勤怠管理を始めるね。")
    readNameLists()
    while True:
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
        print("[log] "+"待機")
