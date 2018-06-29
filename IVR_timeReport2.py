# coding:utf-8
import time
import pyaudio
import wave
import datetime
import os
import random

PATH1 = "./voice_time/"

morningTime = datetime.time(9, 0)
lunchTime = datetime.time(13, 10)
breakTime = datetime.time(16, 0)
eveningTime = datetime.time(19,0)
testTime = datetime.time(0, 0)

morningPath = "morning/"
lunchPath = "lunch/"
breakPath = "break/"
eveningPath = "evening/"
testPath = "test/"
time_list = {morningTime: morningPath, lunchTime: lunchPath,
             breakTime: breakPath, eveningTime: eveningPath,
             testTime: testPath}


def playSound(path):
    '''
    CHUNK = 1024
    filename = path
    wf = wave.open(filename, 'rb')
    p = pyaudio.PyAudio()
    stream = p.open(format=p.get_format_from_width(wf.getsampwidth()),
                    channels=wf.getnchannels(),
                    rate=wf.getframerate(),
                    output=True)
    data = wf.readframes(CHUNK)
    while data != '':
        stream.write(data)
        data = wf.readframes(CHUNK)
    stream.stop_stream()
    stream.close()
    p.terminate()
    '''
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
                print("[log] " + str(now) + ": " + PATH1 + time_list[i] +
                      selectRandom(time_list[i]) + "を再生しました。")
                flg = False
        else:
            flg = True
    return flg


if __name__ == "__main__":
    print("****************************")
    print("   IVR_TimeReport-System")
    print("****************************")

    print("")
    print("時報予定時刻")
    for k in time_list.keys():
        print(k)
    print("")
    print("時間監視開始")
    print("")
    played = True
    while True:
        played = checkTime(played)
        time.sleep(0.5)
