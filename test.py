from pywechat.WechatAuto import AutoReply
from pywechat import check_new_message
from pywechat.WechatAutoReply import WechatAutoReply

if __name__ == '__main__':
    # 处理消息的函数
    def deal_message(messages):
        print(messages)
        return '测试'


    WechatAutoReply.auto_reply_messages_to_friends(deal_message, '10s', ['Snaill'], is_maximize=False)
