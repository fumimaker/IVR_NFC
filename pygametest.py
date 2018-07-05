import time
import pyaudio
import wave

VOICE_PATH1 = "./voice/"
VOICE_PATH2 = ".wav"


def playSound(path):
    CHUNK = 1024
    filename = path
    wf = wave.open(filename, 'rb')
    p = pyaudio.PyAudio()
    stream = p.open(format=p.get_format_from_width(wf.getsampwidth()),
                    channels=wf.getnchannels(),
                    rate=wf.getframerate(),
                    output=True)
    # 1024個読み取り
    data = wf.readframes(CHUNK)
    while data != '':
        stream.write(data)          # ストリームへの書き込み(バイナリ)
        data = wf.readframes(CHUNK)  # ファイルから1024個*2個の
    stream.stop_stream()
    stream.close()
    p.terminate()


def playMorning():
    playSound("./voice/morning.wav")


def playCheck():
    playSound("./voice/check.wav")


def playEvening():
    playSound("./voice/evening.wav")


def playError_NoName():
    playSound("./voice/error_noname.wav")


def playError_DataError():
    playSound("./voice/error_dataerror.wav")


def playNameById(_path):
    path = VOICE_PATH1 + _path + VOICE_PATH2
    print(path)
    playSound(path)


if __name__ == "__main__":
    playCheck()
    playNameById("mizuno")
    playEvening()
    playNameById("mizuno")
    playMorning()
    playError_NoName()
    playError_DataError()