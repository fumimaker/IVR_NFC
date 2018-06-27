# -*- coding: utf-8 -*-

import nfc
import binascii
import datetime
import sys
sys.path.append('C:\Users\LattePanda\AppData\Local\Programs\Python\Python36\Lib\site-packages')
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

##################################################################
import csv
from playsound import playsound
import time

def playCheck():
    playsound('./check.mp3')

def playMorning():
    playsound('./morning.mp3')

def playEvening():
    playsound('./evening.mp3')


def playNameText(idm):
    name = "none"
    if idm == "C3D78B5A": #otsuru
        playsound('./otsuru.mp3')
        name = "otsuru"
    if idm == "01010312D718B809": #鈴木 謙太
        playsound('./suzuki.mp3')
        name = "suzuki"
    if idm == "E3E88B5A": #斉藤 弘恭
        playsound('./saitou.mp3')
        name = "saitou"
    if idm == "93D78B5A": #平井 雄一
        playsound('./yuukun.mp3')
        name = "hirai"
    if idm == "33C78B5A": #田中 修一
        playsound('./tanaka.mp3')
        name = "tanaka"
    if idm == "B3E8F65B": #前田 じゅんき
        playsound('./maeda.mp3')
        name = "maeda"
    if idm == "53C78B5A": #小林 隆貴
        playsound('./kobayashi.mp3')
        name = "kobayashi"
    if idm == "33D5F65B": #堤 修平
        playsound('./tsutsumi.mp3')
        name = "tsutsumi"
    if idm == "A3D78B5A": #赤木 豪
        playsound('./gou.mp3')
        name = "akagi"
    if idm == "13E9F65B": #浅川 雄基
        playsound('./asakawa.mp3')
        name = "asakawa"
    if idm == "011203125A198D0D": #水野 史暁
        playsound('./mizuno.mp3')
        name = "mizuno"
    if idm == "C36F395C": #水野 史暁
        playsound('./mizuno.mp3')
        name = "mizuno"
    if idm == "0114D6AB78181719": #水野 史暁
        playsound('./mizuno.mp3')
        name = "mizuno"
    return name

##################################################################

def process(idm):
    #print(" 検索開始")
    name = playNameText(idm)
    #print(" "+name.encode('utf-8')+"です")  # utf-8で表示    

    time.sleep(0.6)  

    StartOrEnd = False

    dataMatch = True

    if dataMatch == True:

        nowdate = datetime.datetime.now().strftime('%Y%m')
        csvfile = './csv/'+nowdate+'_IVRNFC_'+name+'.csv'

        if os.path.exists(csvfile) == False:
            f = open(csvfile, "w")
            f.write("date,name,now,status,time\n")
            f.close()
        
        f = open(csvfile, "r")
        reader = csv.reader(f)
        # header = next(reader)

        print(" check csv")        
 
        global lastDate
        global nowDate
        global endDate
        global timeDate
        global statusDate
        global namecolnum
        namecolnum =0

        global writeDate
        global writeStart
        global writeEnd
        global writeDuty

        csv_list = {}
        #csv_list = [ e for e in reader]

        #for i in range(len(csv_list)):
         #   print(csv_list[i][0])


        breakFlag = False

        for row in reader:
            rowCount = len(row)
            #print(rowCount)

            if rowCount < 2:                
                break

            if namecolnum == 0:
                csv_list[namecolnum] = row
            else:

                nowDate = ""
                endDate = ""
                statusDate = ""
                timeDate = ""

                csv_list[namecolnum] = row

                workingsheet_list_date = row[0]
                workingsheet_list_name = row[1]
                nowDate = row[2]
                statusDate = row[3]
                timeDate = row[4]

                lastDate = len(workingsheet_list_date)

                if workingsheet_list_date == datetime.datetime.now().strftime('%Y/%m/%d'):
                    if statusDate == "start": #start
                        StartOrEnd = True                        
                    else: #end
                        StartOrEnd = False

                    #print(" すでに名前があります。 "+ "下から"+str(namecolnum))

            namecolnum += 1
        
        f.close()

        f = open(csvfile, "a")

        if StartOrEnd == False:
            print(" 出勤処理")
            playMorning()
            writeDate = datetime.datetime.now().strftime('%Y/%m/%d')
            writeStart = datetime.datetime.now().strftime('%H:%M:%S')
            f.write(writeDate+","+name+","+writeStart+",start,0\n")
        else:
            playEvening()
            print(" 退勤処理")

            #writeDate = datetime.datetime.now().strftime('%Y/%m/%d')
            dt = datetime.datetime.now()
            writeEnd = dt.strftime('%H:%M:%S')
            t1 = nowDate
            tmpDate = workingsheet_list_date + " " + t1
            time_t1 = datetime.datetime.strptime(tmpDate, "%Y/%m/%d %H:%M:%S")
            time_t2 = dt
            time_t3 = time_t2 - time_t1

            (dtt, micro) = str(time_t3).split('.')
            writeDuty = dtt
            f.write(workingsheet_list_date+","+name+","+writeEnd+",end,"+writeDuty+"\n")

        f.close()
'''
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

    
    #wf.close()
    # close PyAudio (7)
    #p.terminate()
