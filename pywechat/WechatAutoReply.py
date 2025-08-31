import re
from collections import deque
from concurrent.futures import ThreadPoolExecutor
import threading
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
        self._lock = threading.Lock()

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
        """Add a message to the pending queue (thread-safe)"""
        with self._lock:
            self._pending_message_queue.append(message)

    def get_pending_message(self) -> Any:
        """Get and remove a message from the pending queue (thread-safe)"""
        with self._lock:
            if self._pending_message_queue:
                return self._pending_message_queue.popleft()
            return None

    def add_processed_message(self, message: Any) -> None:
        """Add a message to the processed queue (thread-safe)"""
        with self._lock:
            self._processed_message_queue.append(message)

    def get_processed_message(self) -> Any:
        """Get and remove a message from the processed queue (thread-safe)"""
        with self._lock:
            if self._processed_message_queue:
                return self._processed_message_queue.popleft()
            return None

    def has_pending_messages(self) -> bool:
        """Check if there are pending messages"""
        with self._lock:
            return len(self._pending_message_queue) > 0

    def has_processed_messages(self) -> bool:
        """Check if there are processed messages"""
        with self._lock:
            return len(self._processed_message_queue) > 0


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
            content_func:自动回复内容函数.入参：{'name': name, 'messages': messages}, 其中name是指定的好友(有备注用备注，没备注用昵称)，messages是获取的消息
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
        friends_map_lock = threading.RLock()
        message_available = threading.Event()  # 用于通知有新消息需要处理
        shutdown_event = threading.Event()  # 用于优雅关闭线程
        
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

        def get_current_chat_friend():
            """获取当前聊天窗口的好友名称，用于避免重复切换"""
            try:
                # 检查当前聊天窗口的标题或其他标识
                current_window = main_window.child_window(**Main_window.CurrentChatWindow)
                if current_window.exists():
                    return current_window.window_text()
                return None
            except:
                return None

        def smart_switch_to_friend(friend_name):
            """智能切换到好友聊天窗口，避免重复操作"""
            current_friend = get_current_chat_friend()
            if current_friend == friend_name:
                print(f"已在 {friend_name} 的聊天窗口，跳过切换操作")
                return False  # 未发生切换
            else:
                Tools.find_friend_in_MessageList(friend=friend_name, is_maximize=is_maximize)
                print(f"从 {current_friend} 切换到 {friend_name}")
                return True  # 发生了切换

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
            if not filtered_messages:
                return
            
            # 批量处理所有需要获取消息的好友，减少窗口切换次数
            for name, num in filtered_messages:
                # 线程安全地创建消息对象
                with friends_map_lock:
                    if not friends_map.get(name):
                        friends_map[name] = Message(name)
                
                # 智能切换到好友聊天窗口（避免重复切换）
                smart_switch_to_friend(name)
                
                # 快速检查是否为好友聊天
                voice_call_button = main_window.child_window(**Buttons.VoiceCallButton)
                video_call_button = main_window.child_window(**Buttons.VideoCallButton)
                
                if not (voice_call_button.exists() and video_call_button.exists()):
                    continue
                    
                # 极速消息获取策略（保持最小停留时间）
                collected_messages = quick_get_messages(name, num)

                
                # 立即添加到待处理队列
                if collected_messages:
                    with friends_map_lock:
                        friend_message = friends_map[name]
                        friend_message.add_pending_message(collected_messages)
                    # 通知异步处理线程有新消息
                    message_available.set()
        
        def quick_get_messages(friend_name, num):
            """极速消息获取：最小化在聊天窗口的停留时间"""
            try:
                friendtype = '好友'
                chatList = main_window.child_window(**Main_window.FriendChatList)
                
                # 一键跳转到聊天底部
                pyautogui.press('End')
                
                # 根据消息数量快速计算翻页策略
                if num <= 10:
                    # 10条以内，直接获取当前页面消息
                    messages = extract_current_page_messages(chatList, friendtype, num)
                else:
                    # 超过10条，快速翻页获取
                    pages_needed = min((num + 9) // 10, 3)  # 最多翻3页，防止过度停留
                    pyautogui.press('pageup', presses=pages_needed, interval=0.02, _pause=False)
                    
                    # 快速检查"查看更多消息"按钮
                    check_more_button = main_window.child_window(**Buttons.CheckMoreMessagesButton)
                    if check_more_button.exists():
                        check_more_button.click_input()
                    
                    # 获取消息
                    messages = extract_current_page_messages(chatList, friendtype, num)
                    
                    # 立即回到底部，减少停留时间
                    pyautogui.press('End')
                
                return messages
                
            except Exception as e:
                print(f"快速获取 {friend_name} 消息时出错: {e}")
                return []
        
        def extract_current_page_messages(chatList, friendtype, max_count):
            """快速提取当前页面的消息，无循环等待"""
            messages = []
            try:
                # 一次性获取所有消息项
                list_items = [msg for msg in chatList.children(control_type='ListItem') 
                             if msg.descendants(control_type='Button')]
                
                # 快速解析消息内容，限制处理数量
                for item in list_items[-max_count:]:  # 多取一些以防有非文本消息
                    try:
                        message_sender, message_content, message_type = Tools.parse_message_content(item, friendtype)
                        if message_type == '文本' and message_content.strip():
                            messages.append(message_content)
                    except:
                        continue  # 跳过解析失败的消息
                
                # 返回最新的消息
                return messages[-max_count:] if len(messages) > max_count else messages
                
            except Exception as e:
                print(f"提取消息时出错: {e}")
                return messages

        @performance_monitor("reply方法")
        def reply():
            """超快速回复方法：最小化在聊天窗口的停留时间"""
            reply_tasks = []  # 收集所有需要回复的任务
            
            # 快速收集所有需要回复的消息
            with friends_map_lock:
                friend_keys = list(friends_map.keys())
            
            for key in friend_keys:
                with friends_map_lock:
                    if key not in friends_map:
                        continue
                    friend_message = friends_map[key]
                
                # 快速收集该好友的所有待回复消息
                messages_to_reply = []
                while True:
                    message = friend_message.get_processed_message()
                    if message is None:
                        break
                    messages_to_reply.append(message)
                
                # 如果有消息需要回复，添加到任务列表
                if messages_to_reply:
                    reply_tasks.append((key, messages_to_reply))
            
            # 如果没有回复任务，直接返回
            if not reply_tasks:
                return
            
            # 超快速执行回复任务
            for friend_name, messages in reply_tasks:
                print(f"极速回复 {friend_name} 的 {len(messages)} 条消息")
                
                # 极速回复策略
                quick_reply_to_friend(friend_name, messages)
            
            # 回复完成后立即激活滑块
            if scrollable:
                mouse.click(coords=(x, y))  # 立即激活滑块，继续遍历会话列表
        
        def quick_reply_to_friend(friend_name, messages):
            """单个好友的极速回复策略"""
            try:
                start_time = time.time()
                
                # 智能切换到好友聊天窗口（避免重复切换）
                smart_switch_to_friend(friend_name)
                
                # 预先获取输入框引用
                current_chat = main_window.child_window(**Main_window.CurrentChatWindow)
                
                # 批量发送消息，减少每条消息之间的延迟
                for i, message in enumerate(messages):
                    # 只在第一条消息时点击输入框
                    if i == 0:
                        current_chat.click_input()
                    
                    # 超快速发送消息
                    Systemsettings.copy_text_to_windowsclipboard(message)
                    pyautogui.hotkey('ctrl', 'v', _pause=False)
                    pyautogui.hotkey('alt', 's', _pause=False)
                    
                    # 只在多条消息时添加最小延迟
                    if len(messages) > 1 and i < len(messages) - 1:
                        time.sleep(0.05)  # 最小延迟，确保消息发送完成
                
                end_time = time.time()
                print(f"极速回复 {friend_name} 耗时: {end_time - start_time:.3f}秒")
                
            except Exception as e:
                print(f"极速回复 {friend_name} 时出错: {e}")

        def initialize_wechat():
            """微信初始化方法，封装所有初始化逻辑"""
            print("正在初始化微信...")
            
            try:
                # 打开微信和初始化UI
                main_window = Tools.open_wechat(wechat_path=wechat_path, is_maximize=is_maximize)
                Tools.open_dialog_window(friend=filetransfer, wechat_path=wechat_path, is_maximize=is_maximize)
                
                # 初始化UI组件
                chat_button = main_window.child_window(**SideBar.Chats)
                chat_button.click_input()
                myname = main_window.child_window(control_type='Button', found_index=0).window_text()
                messageList = main_window.child_window(**Main_window.ConversationList)
                scrollable = Tools.is_VerticalScrollable(messageList)
                x, y = messageList.rectangle().right - 5, messageList.rectangle().top + 8  # 右上方滑块的位置
                
                if scrollable:
                    mouse.click(coords=(x, y))  # 点击右上方激活滑块
                    pyautogui.press('Home')  # 按下Home健确保从顶部开始
                
                print("微信初始化完成")
                return main_window, chat_button, messageList, scrollable, x, y, myname
                
            except Exception as e:
                print(f"微信初始化失败: {e}")
                raise

        def check_and_reconnect_wechat():
            """检查微信状态并在必要时重连"""
            if not Tools.is_wechat_running():
                print("检测到微信进程已退出，尝试重新启动微信...")
                return initialize_wechat()
            return None

        def process_friends_map():
            """事件驱动的异步消息处理线程"""
            print("异步消息处理线程已启动")
            
            while not shutdown_event.is_set():
                try:
                    # 等待新消息通知或超时
                    if message_available.wait(timeout=1.0):
                        message_available.clear()  # 清除事件标志
                        
                        # 批量处理所有待处理的消息
                        has_work = False
                        
                        # 获取当前所有好友的副本，避免长时间持有锁
                        with friends_map_lock:
                            current_friends = dict(friends_map)
                        
                        for name, friend_message in current_friends.items():
                            # 处理该好友的所有待处理消息
                            while True:
                                messages = friend_message.get_pending_message()
                                if messages is None:
                                    break
                                
                                has_work = True
                                print(f"异步处理 {name} 的待处理消息: {len(messages) if isinstance(messages, list) else 1}条")
                                
                                try:
                                    # 调用内容生成函数
                                    if callable(content_func):
                                        reply_content = content_func({'name': name, 'messages': messages})
                                        
                                        # 将处理后的消息添加到已处理队列
                                        friend_message.add_processed_message(reply_content)
                                        print(f"已生成回复内容给 {name}: {reply_content[:50]}...")
                                    else:
                                        print("错误：未提供有效的消息处理函数")
                                        break
                                        
                                except Exception as e:
                                    print(f"处理 {name} 的消息时出错: {e}")
                                    continue
                        
                        if has_work:
                            print("本轮异步处理完成")
                
                except Exception as e:
                    print(f"异步处理线程出错: {e}")
                    time.sleep(0.5)
            
            print("异步消息处理线程已停止")

        # 启动异步消息处理线程
        executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="AsyncMessageProcessor")
        processing_future = None
        
        try:
            # 启动异步处理线程
            processing_future = executor.submit(process_friends_map)

            # 初始化微信
            main_window, chat_button, messageList, scrollable, x, y, myname = initialize_wechat()

            # 主循环：快速获取新消息和发送回复
            while True:
                try:
                    # 检查微信状态并在必要时重连
                    reconnect_result = check_and_reconnect_wechat()
                    if reconnect_result:
                        main_window, chat_button, messageList, scrollable, x, y, myname = reconnect_result
                        print("微信已重新启动并初始化完成")
                        continue
                    
                    # 检查是否有新消息需要获取
                    if chat_button.legacy_properties().get('Value'):
                        # 快速遍历会话列表获取新消息
                        total_start_time = time.time()
                        
                        if scrollable:
                            # 快速遍历模式：只获取消息，不立即回复
                            for page in range(max_pages + 1):
                                page_start_time = time.time()
                                
                                # 快速获取当前页的新消息
                                filtered_messages = record()
                                if filtered_messages:
                                    get_messages(filtered_messages)
                                
                                # 快速翻到下一页
                                pyautogui.press('pagedown', _pause=False)
                                
                                page_end_time = time.time()
                                print(f"第{page+1}页处理耗时: {page_end_time - page_start_time:.3f}秒")
                                
                                # 如果单页处理时间过长，跳过剩余页面
                                if page_end_time - page_start_time > 2.0:  # 单页超过2秒就跳过
                                    print("单页处理时间过长，跳过剩余页面")
                                    break
                            
                            # 快速回到顶部
                            pyautogui.press('Home')
                        else:
                            # 非滚动模式：快速处理
                            filtered_messages = record()
                            if filtered_messages:
                                get_messages(filtered_messages)
                        
                        total_end_time = time.time()
                        print(f"本轮获取消息耗时: {total_end_time - total_start_time:.3f}秒")
                    
                    # 执行回复操作（检查是否有已处理完的消息需要回复）
                    reply()
                    
                    # 智能回到文件传输助手，释放聊天窗口（避免重复切换）
                    smart_switch_to_friend(filetransfer)
                    
                    # 等待一段时间，让异步线程有时间处理消息
                    time.sleep(reply_duration)
                    
                except Exception as loop_error:
                    error_msg = str(loop_error)
                    print(f"主循环发生异常: {loop_error}")
                    print(f"异常类型: {type(loop_error)}")
                    
                    # 检查是否是微信窗口相关的COM错误
                    if "事件无法调用任何订户" in error_msg or "COMError" in str(type(loop_error)):
                        print("检测到微信窗口COM错误，可能微信已经退出，将在下次循环尝试重新启动")
                    else:
                        print("发生其他类型异常，继续运行...")
                    
                    # 等待一段时间后继续循环
                    time.sleep(2)
                
        except Exception as e:
            print(f'微信自动回复异常: {e}')
            
        finally:
            # 优雅关闭异步处理线程
            print("正在关闭异步处理线程...")
            shutdown_event.set()  # 设置关闭信号
            message_available.set()  # 唤醒可能在等待的线程
            
            if processing_future:
                try:
                    processing_future.result(timeout=3)  # 等待最多3秒
                except Exception as e:
                    print(f"关闭异步线程时出错: {e}")
            
            if executor:
                executor.shutdown(wait=True)  # 等待线程完成
                
            Systemsettings.close_listening_mode()
            
            if close_wechat:
                main_window.close()
