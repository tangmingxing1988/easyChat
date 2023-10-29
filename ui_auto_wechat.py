import time
import uiautomation as auto
import subprocess
import numpy as np
import pyperclip
import os
import pyautogui
import requests
import http.server
import socketserver
import threading
import queue
import logging

from PIL import ImageGrab
from clipboard import setClipboardFiles
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QMimeData, QUrl
from typing import List
from urllib.parse import quote
from urllib.parse import parse_qs

location = ''
location_name = ''
destination = '30.298782,120.183518'
destination_name = '朗诗乐府'
my_queue = queue.Queue()

# 配置日志
logging.basicConfig(
    level=logging.INFO,  # 设置日志级别
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',  # 时间戳格式
    handlers=[
        logging.FileHandler('my.log'),  # 将日志输出到文件
        logging.StreamHandler()  # 将日志输出到控制台
    ]
)

# 创建一个日志记录器
logger = logging.getLogger(__name__)

def minutes_to_hours_and_minutes(total_minutes):
    if total_minutes < 0:
        return "3分钟"
    elif total_minutes == 0:
        return "3分钟"

    hours = total_minutes // 60
    minutes = total_minutes % 60

    return f"{hours}小时{minutes}分钟"

def translate_location(origin_cood):
    url = f"https://apis.map.qq.com/ws/coord/v1/translate?key=4C7BZ-OMMKT-CN7XP-VQOGV-U7CFO-I7F5B&locations={origin_cood}&type=1"
    # 发送GET请求
    response = requests.get(url)

    # 检查响应状态码
    if response.status_code == 200:
        # 解析JSON响应
        data = response.json()
        
        # 提取lng和lat的值
        if data['status'] == 0:
            lng = data["locations"][0]["lng"]
            lat = data["locations"][0]["lat"]
            return f"{lat},{lng}"
        
    return origin_cood

# 查看距离和剩余时间
def from_destination():
    if len(location) > 0 and len(destination) > 0:
        # 定义请求的URL
        url = f"https://apis.map.qq.com/ws/direction/v1/driving/?from={location}&to={destination}&output=json&key=4C7BZ-OMMKT-CN7XP-VQOGV-U7CFO-I7F5B"

        # 发送GET请求
        response = requests.get(url)

        # 检查响应状态码
        if response.status_code == 200:
            # 解析JSON响应
            data = response.json()
            
            # 提取distance和duration的值
            if data['status'] == 0:
                distance = data["result"]["routes"][0]["distance"]
                duration = data["result"]["routes"][0]["duration"]

                return (distance, duration)
        
    # 否则返回0
    return (0, 0)
    
class MyHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        logger.info("收到Get请求")
        if self.path == "/location":
            self.send_response(302)

            # 拼接URL模板
            marker_template = "coord:{};title:夏维英;addr:{}"
            marker = marker_template.format(location, location_name)
            
            url = f"https://apis.map.qq.com/tools/poimarker?type=0&marker={quote(marker)}&key={quote('4C7BZ-OMMKT-CN7XP-VQOGV-U7CFO-I7F5B')}&referer=myapp"
            self.send_header("Location", url)
            self.end_headers()
            logger.info("发送完Get请求")

    def do_POST(self):
        content_length = int(self.headers["Content-Length"])
        post_data = self.rfile.read(content_length).decode("utf-8")
        post_params = parse_qs(post_data)

        if "from" in post_params and "content" in post_params:
            # 处理接收到的数据，你可以根据需要进行操作
            received_from = post_params["from"][0]
            received_content = post_params["content"][0]
            print("Received from:", received_from)
            print("Received content:", received_content)
            # 使用 splitlines() 方法将字符串分割成行
            lines = received_content.splitlines()

            # 获取第二行（索引为 1）
            if len(lines) > 1:
                second_line = lines[1]
                received_from = f"(来自{received_from})" if len(received_from) <= 11 else ''
                my_queue.put(f"{second_line}{received_from}")
            else:
                print("字符串中没有第二行。")


        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write("".encode("utf-8"))  # 返回一个空文本

# 鼠标移动到控件上
def move(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(x, y)

# 鼠标快速点击控件
def click(element):
    x, y = element.GetPosition()
    auto.Click(x, y)


# 鼠标右键点击控件
def right_click(element):
    x, y = element.GetPosition()
    auto.RightClick(x, y)


# 鼠标快速点击两下控件
def double_click(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(x, y)
    element.DoubleClick()


# 微信的控件介绍。注意"depth"是直接调用auto进行控件搜索的深度（见函数内部代码示例）
# 以群名“测试”为例：
# 左侧聊天列表“测试”群               Name: '测试'     ControlType: ListItemControl    depth: 10
# 左侧聊天列表“测试”群               Name: '测试'     ControlType: ButtonControl      depth: 12
# 进入“测试”群界面之后上方的群名       Name: '测试'     ControlType: ButtonControl      depth: 14
# “测试”群界面的内容框               Name: '消息'     ControlType: ListControl        depth: 12
# 聊天界面的聊天记录按钮              Name: '聊天记录'   ControlType: ButtonControl      depth: 14
# 聊天记录界面的图片按钮              Name: '图片与视频'     ControlType: TabItemControl      depth: 6
# 聊天记录复制图片按钮               Name: '复制'   ControlType: MenuItemControl      depth: 5

class WeChat:
    def __init__(self, path):
        # 微信打开路径
        self.path = path
        
        # 用于复制内容到剪切板
        self.app = QApplication([])
        
        # 自动回复的联系人列表
        self.auto_reply_contacts = []
        
        # 自动回复的内容
        self.auto_reply_msg = "[自动回复]您好，我现在正在忙，稍后会主动联系您，感谢理解。"
        

    def get_location(self):
        # 定义请求的URL和请求体数据
        url = "http://x.hupai.vip:9999/location/query"
        data = {
            "data": {},
            "timestamp": 1652590258638,
            "sign": ""
        }

        # 发送POST请求
        response = requests.post(url, json=data)

        # 检查响应状态码是否为200
        if response.status_code == 200:
            # 解析响应的JSON数据
            response_data = response.json()

            # 获取address字段
            address = response_data.get("data", {}).get("address")

            if address:
                latitude = response_data.get("data", {}).get("latitude")
                longitude = response_data.get("data", {}).get("longitude")
                
                global location, location_name
                location = translate_location(f"{latitude},{longitude}")
                location_name = address

                dest_reply = ''
                dest_distance, dest_duration = from_destination()
                if dest_duration > 0:
                    dest_reply = f"距离{destination_name}还有{dest_distance//1000}公里，大约需要{minutes_to_hours_and_minutes(dest_duration)}，"

                return f"我在【{address}】，{dest_reply}具体位置是：http://x.hupai.vip:8080/location"
            
        return "我在火星，别来烦我"

    # 打开微信客户端
    def open_wechat(self):
        subprocess.Popen(self.path)
    
    # 搜寻微信客户端控件
    def get_wechat(self):
        return auto.WindowControl(Depth=1, Name="微信")
    
    # 搜索指定用户
    def get_contact(self, name):
        self.open_wechat()
        self.get_wechat()
        
        search_box = auto.EditControl(Depth=8, Name="搜索")
        click(search_box)
        
        pyperclip.copy(name)
        auto.SendKeys("{Ctrl}v")
        # 等待客户端搜索联系人
        time.sleep(0.1)
        search_box.SendKeys("{enter}")
    
    # 鼠标移动到发送按钮处点击发送消息
    def press_enter(self):
        # 获取发送按钮
        send_button = auto.ButtonControl(Depth=15, Name="发送(S)")
        click(send_button)
    
    # 在指定群聊中@他人（若@所有人需具备@所有人权限）
    def at(self, name, at_name):
        self.get_contact(name)
        
        # 如果at_name为空则代表@所有人
        if at_name == "":
            auto.SendKeys("@{UP}{enter}")
            self.press_enter()
        
        else:
            auto.SendKeys(f"@{at_name}")
            # 按下回车键确认要at的人
            auto.SendKeys("{enter}")
            self.press_enter()
    
    # 搜索指定用户名的联系人发送信息
    def send_msg(self, name, text):
        self.get_contact(name)
        pyperclip.copy(text)
        auto.SendKeys("{Ctrl}v")
        self.press_enter()
    
    # 搜索指定用户名的联系人发送文件
    def send_file(self, name: str, path: str):
        """
        Args:
            name: 指定用户名的名称，输入搜索框后出现的第一个人
            path: 发送文件的本地地址
        """
        
        # 粘贴文件发送给用户
        self.get_contact(name)
        # 将文件复制到剪切板
        setClipboardFiles([path])
        
        auto.SendKeys("{Ctrl}v")
        self.press_enter()
    
    # 获取所有通讯录中所有联系人
    def find_all_contacts(self):
        self.open_wechat()
        self.get_wechat()
        
        # 获取通讯录管理界面
        click(auto.ButtonControl(Name="通讯录"))
        list_control = auto.ListControl(Name="联系人")
        scroll_pattern = list_control.GetScrollPattern()
        scroll_pattern.SetScrollPercent(-1, 0)
        contacts_menu = list_control.ButtonControl(Name="通讯录管理")
        click(contacts_menu)
        
        # 切换到通讯录管理界面
        contacts_window = auto.GetForegroundControl()
        list_control = contacts_window.ListControl()
        scroll_pattern = list_control.GetScrollPattern()
        
        # 读取用户
        contacts = []
        # 如果不存在滑轮则直接读取
        if scroll_pattern is None:
            for contact in contacts_window.ListControl().GetChildren():
                # 获取用户的昵称以及备注
                name = contact.TextControl().Name
                note = contact.ButtonControl(foundIndex=2).Name

                # 有备注的用备注，没有备注的用昵称
                if note == "":
                    contacts.append(name)
                else:
                    contacts.append(note)
        else:
            for percent in np.arange(0, 1.002, 0.001):
                scroll_pattern.SetScrollPercent(-1, percent)
                for contact in contacts_window.ListControl().GetChildren():
                    # 获取用户的昵称以及备注
                    name = contact.TextControl().Name
                    note = contact.ButtonControl(foundIndex=2).Name

                    # 有备注的用备注，没有备注的用昵称
                    if note == "":
                        contacts.append(name)
                    else:
                        contacts.append(note)
        
        # 返回去重过后的联系人列表
        return list(set(contacts))
    
    # 检测微信是否收到新消息
    def check_new_msg(self):
        self.open_wechat()
        self.get_wechat()
        
        # 获取左侧聊天按钮
        chat_btn = auto.ButtonControl(Name="聊天")
        item = auto.ListItemControl()
        double_click(chat_btn)
        # 持续点击聊天按钮，直到获取完全部新消息
        first_name = item.Name
        while True:
            print(item.Name)
            # 判断该联系人是否需要自动回复
            if item.Name in self.auto_reply_contacts:
                self.auto_reply(item, self.auto_reply_msg)
                # print("需要自动回复")
            
            # 跳转到下一个新消息
            double_click(chat_btn)
            item = auto.ListItemControl()
            # 已经完成遍历，退出循环
            if first_name == item.Name:
                break
    
    # 设置自动回复的联系人
    def set_auto_reply(self, contacts):
        # contacts是一个列表
        self.auto_reply_contacts = contacts
    
    # 自动回复
    def auto_reply(self, element, text):
        click(element)
        pyperclip.copy(text)
        auto.SendKeys("{Ctrl}v")
        self.press_enter()

    def direct_reply(self, text):
        pyperclip.copy(text)
        auto.SendKeys("{Ctrl}v")
        self.press_enter()

    # 识别聊天内容的类型
    # 0：用户发送    1：时间信息  2：红包信息  3：”查看更多消息“标志 4：撤回消息
    def _detect_type(self, list_item_control: auto.ListItemControl) -> int:
        value = None
        # 判断内容框是否为时间框，如果是时间框则子控件不是PaneControl
        if not isinstance(list_item_control.GetFirstChildControl(), auto.PaneControl):
            value = 1
        
        else:
            cnt = 0
            for child in list_item_control.PaneControl().GetChildren():
                cnt += len(child.GetChildren())
            
            # 判断是否为用户发送的信息
            if cnt > 0:
                value = 0
            # 判断是否为“查看更多消息”
            elif list_item_control.Name == "查看更多消息":
                value = 3
            # 或者是红包信息
            elif "红包" in list_item_control.Name:
                value = 2
            # 或者是撤回消息
            elif "撤回了一条消息" in list_item_control.Name:
                value = 4
            # 分割线
            elif "以下为新消息" in list_item_control.Name:
                value = 5
            elif "邀请你加入了群聊" in list_item_control.Name:
                value = 6
                
        if value is None:
            value = 7
        
        return value
    
    # 获取聊天窗口
    def _get_chat_frame(self, name: str):
        self.get_contact(name)
        return auto.ListControl(Name="消息")
    
    def save_dialog_pictures(self, name: str, num: int, save_dir: str) -> None:
        """
        保存指定聊天记录中的图片。图片的名字代表图片在聊天记录中的顺序，从1开始代表最新的图片。
        Args:
            name: 聊天窗口的名字
            num: 保存的最大数量（从最新图片开始保存）
            save_dir: 保存的目录
        """
        
        # 进入图片聊天记录界面
        self.get_contact(name)
        click(auto.ButtonControl(Name="聊天记录", Depth=14))
        click(auto.TabItemControl(Name="图片与视频", Depth=6))
        
        # 图片栏控件
        list_control = auto.ListControl(Name="图片与视频", Depth=6)
        
        # 如果图片数量 < num，则继续往上翻直到满足条件或无法上翻为止
        move(list_control.GetLastChildControl())
        pictures = set()
        cnt = 0
        while cnt < num:
            ori_cnt = cnt
            for list_item_control in list_control.GetChildren()[::-1]:
                # 如果标签不是图片则跳过
                if len(list_item_control.GetFirstChildControl().GetChildren()) == 3:
                    continue
                
                if cnt < num:
                    # 复制图片到剪切板
                    right_click(list_item_control)
                    menu = auto.ListControl(Depth=4)
                    copy = menu.GetFirstChildControl()
                    # 如果图片已经被清理则跳过
                    if copy.Name != "复制":
                        continue
                    else:
                        click(auto.MenuItemControl(Name="复制", Depth=5))
                    
                    # 获取图片路径防止重复存储
                    pic_hash = ImageGrab.grabclipboard()[0]

                    # 获取后缀
                    suffix = pic_hash.split(".")[-1]
                    
                    # 保存图片
                    if pic_hash not in pictures:
                        cnt += 1
                        pictures.add(pic_hash)
                        save_path = os.path.join(save_dir, f"{cnt}.{suffix}")
                        os.system(f"copy \"{pic_hash}\" \"{save_path}\"")
            # 上滑
            pyautogui.scroll(300)
            # 如果无法上滑则退出
            if ori_cnt == cnt:
                break

    # 获取当前聊天窗口的聊天记录
    def get_current_contents(self) -> List:
        list_control = auto.ListControl(Depth=12, Name="消息")
        
        dialogs = []
        value_to_info = {0: '用户发送', 1: '时间信息', 2: '红包信息', 3: '"查看更多消息"标志', 4: '撤回消息', 5: '新消息分割', 6: '邀请进群', 7: '其他'}

        for list_item_control in list_control.GetChildren()[::-1]:
            v = self._detect_type(list_item_control)
            msg = list_item_control.Name
            name = list_item_control.ButtonControl().Name if v == 0 else ''
            
            dialogs.append((value_to_info[v], name, msg))
            
        return dialogs
            
    # 获取指定聊天窗口的聊天记录
    def get_dialogs(self, name: str, n_msg: int) -> List:
        """
        Args:
            name: 聊天窗口的姓名
            n_msg: 获取聊天记录的最大数量（从最后一条往上算）

        Return:
            dialogs: 聊天记录列表，内部元素为三元组（信息类型，发送人，发送内容）
        """
        list_control = self._get_chat_frame(name)
        scroll_pattern = list_control.GetScrollPattern()
        move(scroll_pattern)
        
        # 如果聊天记录数量 < n_msg，则继续往上翻直到满足条件或无法上翻为止
        while len(list_control.GetChildren()) < n_msg:
            # 将聊天记录翻到“查看更多消息”
            scroll_pattern.SetScrollPercent(-1, 0)
            # 如果无法上翻则退出
            first_item = list_control.GetFirstChildControl()
            if self._detect_type(first_item) != 3:
                break
            # 否则点击“查看更多消息”
            else:
                click(first_item)

        cnt = 0
        dialogs = []
        value_to_info = {0: '用户发送', 1: '时间信息', 2: '红包信息', 3: '"查看更多消息"标志', 4: '撤回消息'}
        # 从下往上依次记录聊天内容。
        for list_item_control in list_control.GetChildren()[::-1]:
            v = self._detect_type(list_item_control)
            msg = list_item_control.Name
            name = list_item_control.ButtonControl().Name if v == 0 else ''
            
            cnt += 1
            dialogs.append((value_to_info[v], name, msg))
            
            # 如果达到n_msg则退出
            if cnt == n_msg:
                break
        
        # 将聊天记录列表翻转
        dialogs = dialogs[::-1]
        return dialogs


def start_http_server():
    port = 8080
    with socketserver.TCPServer(("", port), MyHandler) as httpd:
        print("HTTP server is running at port", port)
        httpd.serve_forever()

if __name__ == '__main__':
    wechat_path = "C:\Program Files\Tencent\WeChat\WeChat.exe"
    wechat = WeChat(wechat_path)
    
    # 监听消息
        
    # 创建并启动HTTP服务器线程
    http_server_thread = threading.Thread(target=start_http_server)
    http_server_thread.start()

    # print(wechat.get_location())
    # time.sleep(1000)

    # name = "文件传输助手"
    # text = "你好"
    # file = "C:/Users/Dell/Pictures/takagi.jpeg"
    
    # wechat.send_msg(name, text)
    # wechat.send_file(name, file)
    print('开始使用微信')
    while True:
        dialogs = wechat.get_current_contents()
        if len(dialogs) > 0:
                latest_message = dialogs[0][2]
                for msg in dialogs:
                    if msg[0] == '用户发送' and msg[1] == '夏维英' and '我在' in msg[2]:
                        break
                    elif msg[0] == '用户发送' and msg[1] != '夏维英' and '妈' in msg[2] and '哪' in msg[2]:
                        logger.info('请求位置')
                        wechat.direct_reply(wechat.get_location())
                        break

        try:
            item = my_queue.get_nowait()
            wechat.direct_reply(item)
        except queue.Empty:
            logger.info("没有需要发送的消息")
        time.sleep(1)
    
    # contacts = wechat.find_all_contacts()
    # print(len(contacts))
    
    # res = wechat.get_dialogs("easychat", 100)
    # for i in res:
    #     print(i)
    
    # wechat.save_dialog_pictures("xx", 15, "C:/Users/LTEnj/Desktop/")