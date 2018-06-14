# coding:utf-8

import nfc
import binascii
import datetime
import sys

dict = {"fumi(Suica)":'011203125A198D0D', "fumi(studentCard)":'0114D6AB78181719', "suzuki":'C36F395C'}

def connected(tag):
    print('ICカード検知 Type = {}'.format(tag.type))
    now = datetime.datetime.now()
    idm = binascii.hexlify(tag.identifier).upper() 

    if idm in dict.values() == False:
        print('登録されているICカードではありません。再度やり直してください。')
    else:
        print( 'ICカードIDm読み取り成功  IDm = '+idm  )
    for name, id_val in dict.items():
        if id_val == idm:
            print(str(now) + ' {}さんが出勤しました。'.format(name))
    return idm

while True:
    print('****************************')
    print('ICカードをタッチしてください。')
    print('****************************')
    clf = nfc.ContactlessFrontend('usb')
    clf.connect(rdwr={'on-connect': connected}) # now touch a tag    
    clf.close()
    sys.exit()

