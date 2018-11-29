# coding:utf-8
import time
import pyaudio
import wave
import datetime
import os
import random
from slacker import Slacker

PATH1 = "./voice_time/"

morningTime = datetime.time(10, 0)
lunchTime = datetime.time(13, 10)
breakTime = datetime.time(16, 0)
eveningTime = datetime.time(19, 0)
testTime = datetime.time(0, 0)

morningPath = "morning/"
lunchPath = "lunch/"
breakPath = "break/"
eveningPath = "evening/"
testPath = "test/"
time_list = {morningTime: morningPath, lunchTime: lunchPath,
             breakTime: breakPath, eveningTime: eveningPath,
             testTime: testPath}

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


def playSound(path):
    cmd = "aplay " + path
    os.system(cmd)


def selectRandom(name):
    path = PATH1 + name
    files = os.listdir(path)
    files_file = [f for f in files if os.path.isfile(os.path.join(path, f))]
    return random.choice(files_file)


def checkTime(tmp):
    flg = tmp
    now = datetime.datetime.now()
    for i in time_list.keys():
        if now.hour == i.hour and now.minute == i.minute \
                and now.second == i.second:
            if flg:
                playSound(PATH1 + time_list[i] + selectRandom(time_list[i]))
                mozi = "[log] " + str(now) + ": " + PATH1 + time_list[i] + \
                    selectRandom(time_list[i]) + "を再生しました。"
                print(mozi)
                slack.post_message_to_channel(channelName, mozi)
                flg = False
        else:
            flg = True
    return flg


def postSlack(mozi):
    slack.post_message_to_channel(channelName, mozi)


if __name__ == "__main__":
    slack = Slack("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
    postSlack("[log]時報システムを起動するよ！")
    print("****************************")
    print("   IVR_TimeReport-System")
    print("****************************")

    print("")
    print("時報予定時刻")
    mozi = " ``` "+"時報予定時刻\n"
    for k in time_list.keys():
        print(k)
        mozi = mozi + str(k) + "\n"
    mozi = mozi + "```"
    postSlack(mozi)
    print("")
    print("時間監視開始")
    print("")
    postSlack("[log]時間監視開始します。")
    played = True
    while True:
        played = checkTime(played)
        time.sleep(0.5)
