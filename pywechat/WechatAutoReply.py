import re
from collections import deque
from concurrent.futures import ThreadPoolExecutor
import time
from typing import Deque, Any, Callable
import pyautogui
from pywinauto import mouse
from pywechat import match_duration, TimeNotCorrectError, Systemsettings, Tools
from .Uielements import (Main_window, SideBar, Independent_window, Buttons, SpecialMessages,
                         Edits, Texts, TabItems, Lists, Panes, Windows, CheckBoxes, MenuItems, Menus, ListItems)

# 性能监控装饰器
def performance_monitor(func_name):
    def decorator(func):
        def wrapper(*args, **kwargs):
            start_time = time.time()
            result = func(*args, **kwargs)
            end_time = time.time()
            print(f"[耗时统计] {func_name} 耗时: {end_time - start_time:.3f}秒")
            return result
        return wrapper
    return decorator


class Message:
    def __init__(self, nickname: str):
        self._nickname: str = nickname
        self._pending_message_queue: Deque[Any] = deque()
        self._processed_message_queue: Deque[Any] = deque()

    # Getter methods
    def get_nickname(self) -> str:
        """Get the nickname"""
        return self._nickname

    def get_pending_message_queue(self) -> Deque[Any]:
        """Get the pending message queue"""
        return self._pending_message_queue

    def get_processed_message_queue(self) -> Deque[Any]:
        """Get the processed message queue"""
        return self._processed_message_queue

    # Setter methods
    def set_nickname(self, nickname: str) -> None:
        """Set the nickname"""
        self._nickname = nickname

    def add_pending_message(self, message: Any) -> None:
        """Add a message to the pending queue"""
        self._pending_message_queue.append(message)

    def add_processed_message(self, message: Any) -> None:
        """Add a message to the processed queue"""
        self._processed_message_queue.append(message)


language = Tools.language_detector()  # 有些功能需要判断语言版本
Main_window = Main_window()  # 主界面UI
SideBar = SideBar()  # 侧边栏UI
Independent_window = Independent_window()  # 独立主界面
Buttons = Buttons()  # 所有Button类型UI
Edits = Edits()  # 所有Edit类型UI
Texts = Texts()  # 所有Text类型UI
TabItems = TabItems()  # 所有TabIem类型UI
Lists = Lists()  # 所有列表类型UI
Panes = Panes()  # 所有Pane类型UI
Windows = Windows()  # 所有Window类型UI
CheckBoxes = CheckBoxes()  # 所有CheckBox类型UI
MenuItems = MenuItems()  # 所有MenuItem类型UI
Menus = Menus()  # 所有Menu类型UI
ListItems = ListItems()  # 所有ListItem类型UI
SpecialMessages = SpecialMessages()  # 特殊消息
pyautogui.FAILSAFE = False  # 防止鼠标在屏幕边缘处造成的误触


class WechatAutoReply:

    @staticmethod
    def auto_reply_messages_to_friends(content_func: Callable, reply_duration: str = '10s', friends: list = [], max_pages: int = 5, wechat_path: str = None,
                                       is_maximize: bool = True, close_wechat: bool = True) -> None:
        '''
        该方法用来遍历会话列表查找新消息自动回复,最大回复数量=max_pages*(8~10)\n
        只会回复指定好友和文本类信息
        Args:
            content_func:自动回复内容函数
            reply_duration:自动回复持续时长,格式:'s','min','h'单位:s/秒,min/分,h/小时. 默认10s
            friends: 指定的好友才会回复(有备注用备注，没备注用昵称)
            max_pages:遍历会话列表页数,一页为8~10人,设定持续时间后,将持续在max_pages内循环遍历查找是否有新消息
            wechat_path:微信的WeChat.exe文件地址,主要针对未登录情况而言,一般而言不需要传入该参数,因为pywechat会通过查询环境变量,注册表等一些方法
                尽可能地自动找到微信路径,然后实现无论PC微信是否启动都可以实现自动化操作,除非你的微信路径手动修改过,发生了变动的话可能需要
                传入该参数。最后,还是建议加入到环境变量里吧,这样方便一些。加入环境变量可调用set_wechat_as_environ_path函数
            is_maximize:微信界面是否全屏,默认全屏。
            close_wechat:任务结束后是否关闭微信,默认关闭
        '''
        friends_map = dict()
        reply_duration = match_duration(reply_duration)
        if not reply_duration:
            raise TimeNotCorrectError
        Systemsettings.open_listening_mode(full_volume=False)

        if language == '简体中文':
            filetransfer = '文件传输助手'
        if language == '英文':
            filetransfer = 'File Transfer'
        if language == '繁体中文':
            filetransfer = '檔案傳輸'

        @performance_monitor("record方法")
        def record():
            # 遍历当前会话列表内可见的所有成员，获取他们的名称和新消息条数，没有新消息的话返回[]
            # newMessage friends为会话列表(List)中所有含有新消息的ListItem
            newMessagefriends = [friend for friend in messageList.items() if '条新消息' in friend.window_text()]
            if newMessagefriends:
                # newMessageTips为newMessage friends中每个元素的文本:['测试365 5条新消息','一家人已置顶20条新消息']这样的字符串列表
                newMessageTips = [friend.window_text() for friend in newMessagefriends]
                # 会话列表中的好友具有Text属性，Text内容为备注名，通过这个按钮的名称获取好友名字
                names = [friend.descendants(control_type='Text')[0].window_text() for friend in newMessagefriends]
                # 此时filtered_Tips变为：['5条新消息','20条新消息']直接正则匹配就不会出问题了
                filtered_Tips = [friend.replace(name, '') for name, friend in zip(names, newMessageTips)]
                nums = [int(re.findall(r'\d+', tip)[0]) for tip in filtered_Tips]
                # 过滤出满足条件的朋友
                return [(name, num) for name, num in list(zip(names, nums)) if name in friends]
            return []

        @performance_monitor("get_messages方法")
        def get_messages(filtered_messages):
            # 获取文本类聊天信息
            if filtered_messages:
                for name, num in filtered_messages:
                    if not friends_map.get(name):
                        friends_map[name] = Message(name)

                    Tools.find_friend_in_MessageList(friend=name, is_maximize=is_maximize)[1]
                    check_more_messages_button = main_window.child_window(**Buttons.CheckMoreMessagesButton)
                    voice_call_button = main_window.child_window(**Buttons.VoiceCallButton)  # 语音聊天按钮
                    video_call_button = main_window.child_window(**Buttons.VideoCallButton)  # 视频聊天按钮
                    # 只处理好友,不处理群聊和公众号
                    if voice_call_button.exists() and video_call_button.exists():
                        friendtype = '好友'
                        chatList = main_window.child_window(**Main_window.FriendChatList)
                        x, y = chatList.rectangle().left + 10, (
                                main_window.rectangle().top + main_window.rectangle().bottom) // 2
                        ListItems = [message for message in chatList.children(control_type='ListItem') if
                                     message.descendants(control_type='Button')]
                        # 点击聊天区域侧边栏靠里一些的位置,依次来激活滑块,不直接main_window.click_input()是为了防止点到消息
                        mouse.click(coords=(x, y))
                        # 按一下pagedown到最下边
                        pyautogui.press('pagedown')
                        ########################
                        # 需要先提前向上遍历一遍,防止语音消息没有转换完毕
                        if num <= 10:  # 10条消息最多先向上遍历number//3页
                            pages = num // 3
                        else:
                            pages = num // 2  # 超过10条
                        for _ in range(pages):
                            if check_more_messages_button.exists():
                                check_more_messages_button.click_input()
                                mouse.click(coords=(x, y))
                            pyautogui.press('pageup', _pause=False)
                        pyautogui.press('End')
                        mouse.click(coords=(x, y))
                        pyautogui.press('pagedown')
                        # 开始记录消息
                        while len(list(set(ListItems))) < num:
                            if check_more_messages_button.exists():
                                check_more_messages_button.click_input()
                                mouse.click(coords=(x, y))
                            pyautogui.press('pageup', _pause=False)
                            ListItems.extend([message for message in chatList.children(control_type='ListItem') if
                                              message.descendants(control_type='Button')])
                        pyautogui.press('End')
                        #######################################################
                        ListItems = ListItems[-num:]
                        message_contents = []
                        for ListItem in ListItems:
                            message_sender, message_content, message_type = Tools.parse_message_content(ListItem,
                                                                                                        friendtype)
                            if message_type == '文本':
                                message_contents.append(message_content)
                        friend_message = friends_map[name]
                        friend_message.add_pending_message(message_contents)

        @performance_monitor("reply方法")
        def reply():
            for key, value in friends_map.items():
                processed_message_queue = value.get_processed_message_queue()
                while processed_message_queue:
                    message = processed_message_queue.popleft()  # 取出最早的消息
                    print(f"回复 {key} 的消息: {message}")
                    Tools.find_friend_in_MessageList(friend=key, is_maximize=is_maximize)
                    current_chat = main_window.child_window(**Main_window.CurrentChatWindow)
                    current_chat.click_input()
                    Systemsettings.copy_text_to_windowsclipboard(message)
                    pyautogui.hotkey('ctrl', 'v', _pause=False)
                    pyautogui.hotkey('alt', 's', _pause=False)

            if scrollable:
                mouse.click(coords=(x, y))  # 回复完成后点击右上方,激活滑块，继续遍历会话列表

        def process_friends_map():
            """单独的工作线程，专门处理 friends_map 中的数据"""
            while True:
                try:
                    # 遍历 friends_map 中的所有好友
                    for name, friend_message in friends_map.items():
                        # 获取待处理的消息队列
                        pending_queue = friend_message.get_pending_message_queue()

                        # 处理待处理的消息
                        while pending_queue:
                            messages = pending_queue.popleft()
                            print(f"工作线程处理 {name} 的待处理消息: {messages}")

                            # 调用内容生成函数
                            if callable(content_func):
                                reply_content = content_func(messages)
                            else:
                                raise Exception('请传入处理消息的函数')

                            # 将处理后的消息添加到已处理队列
                            friend_message.add_processed_message(reply_content)
                            print(f"已将回复内容添加到 {name} 的已处理队列: {reply_content}")

                except Exception as e:
                    print(f"工作线程处理 friends_map 时出错: {e}")
                    time.sleep(1)

        executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="FriendsMapProcessor")
        try:
            # 使用线程池
            executor.submit(process_friends_map)
            print("工作线程已启动")
            # 打开文件传输助手是为了防止当前窗口有好友给自己发消息无法检测到,因为当前窗口即使有新消息也不会在会话列表中好友头像上显示数字,
            main_window = Tools.open_dialog_window(friend=filetransfer, wechat_path=wechat_path, is_maximize=is_maximize)[1]
            myname = main_window.child_window(control_type='Button', found_index=0).window_text()
            messageList = main_window.child_window(**Main_window.ConversationList)
            scrollable = Tools.is_VerticalScrollable(messageList)
            x, y = messageList.rectangle().right - 5, messageList.rectangle().top + 8  # 右上方滑块的位置
            if scrollable:
                mouse.click(coords=(x, y))  # 点击右上方激活滑块
                pyautogui.press('Home')  # 按下Home健确保从顶部开始
            search_pages = 1

            chatsButton = main_window.child_window(**SideBar.Chats)
            while True:
                if chatsButton.legacy_properties().get('Value'):  # 如果左侧的聊天按钮式红色的就遍历,否则原地等待
                    if scrollable:
                        for _ in range(max_pages + 1):
                            filtered_messages = record()
                            if filtered_messages:
                                get_messages(filtered_messages)
                                pyautogui.press('pagedown', _pause=False)
                            search_pages += 1
                        pyautogui.press('Home')

                    else:
                        filtered_messages = record()
                        if filtered_messages:
                            get_messages(filtered_messages)
                reply()

                Tools.open_dialog_window(friend=filetransfer, wechat_path=wechat_path,
                                         is_maximize=is_maximize)[1]
                time.sleep(reply_duration)
        except Exception as e:
            print('微信自动回复异常:', e)
        finally:
            # 关闭线程池
            if executor:
                executor.shutdown(wait=False)  # 不等待线程完成，直接关闭
            Systemsettings.close_listening_mode()
            if close_wechat:
                main_window.close()
