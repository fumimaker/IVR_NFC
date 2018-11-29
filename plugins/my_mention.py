# coding: utf-8

from slackbot.bot import respond_to     # @botname: で反応するデコーダ
from slackbot.bot import listen_to      # チャネル内発言で反応するデコーダ
from slackbot.bot import default_reply  # 該当する応答がない場合に反応するデコーダ
import PlayVoice
# @respond_to('string')     bot宛のメッセージ
#                           stringは正規表現が可能 「r'string'」
# @listen_to('string')      チャンネル内のbot宛以外の投稿
#                           @botname: では反応しないことに注意
#                           他の人へのメンションでは反応する
#                           正規表現可能
# @default_reply()          DEFAULT_REPLY と同じ働き
#                           正規表現を指定すると、他のデコーダにヒットせず、
#                           正規表現にマッチするときに反応
#                           ・・・なのだが、正規表現を指定するとエラーになる？

# message.reply('string')   @発言者名: string でメッセージを送信
# message.send('string')    string を送信
# message.react('icon_emoji')  発言者のメッセージにリアクション(スタンプ)する
#                               文字列中に':'はいらない


@respond_to(u'メンション')
def mention_func(message):
    message.reply('どうかしましたか？')  # メンション


@listen_to(u'あかりちゃん')
def listen_func(message):
    message.send('誰か私を呼んだような気がします。')      # ただの投稿
    message.reply('私を呼んだのはあなたですね？')


@respond_to(u'かわいい')
def cool_func(message):
    message.reply('ありがとうございます!')     # メンション
    message.react('heart')     # リアクション


@listen_to(u'かわいい')
def listen_func(message):
    message.reply('もしかして私を呼びましたか？')


@default_reply()
def default_func(message):
    text = message.body[u'text']     # メッセージを取り出す
    # 送信メッセージを作る。改行やトリプルバッククォートで囲む表現も可能
    msg = u'あなたの送ったメッセージは\n```' + text + '```'
    message.reply(msg)      # メンション


@respond_to(u'ビッチ')
def mention_func(message):
    message.reply('あへぇ...')  # メンション


@listen_to(u'IVRアキハバラスタジオ受付に、「配達業者さま専用」ボタンから')
def listen_func(message):
    message.send('<!here> 何か荷物がきたみたいですよ！どなたか出ていただけませんか？')
    PlayVoice.playSound("./voice/nimotu.wav")
    print("荷物が来ました。")


@listen_to(u'IVRアキハバラスタジオ受付に、「総合受付」ボタンから')
def listen_func(message):
    message.send('<!here> 外に誰か来たみたいですよ。どなたか出ていただけませんか？')
    PlayVoice.playSound("./voice/guest.wav")
    print("お客さんが来ました。")
