import PySimpleGUI as sg
import win32api
import win32con
from time import sleep
import win32gui
import requests
import pyperclip as clip
from pynput.keyboard import Key, Listener
import collections
import datetime
import random
import urllib3.contrib.pyopenssl
import os
import ctypes
import psutil

def AutoCloseMessageBoxW(text, title):
    ctypes.windll.user32.MessageBoxA(0, text.encode('gb2312'), title.encode('gb2312'))

def kill_process_one():
    n=0
    while True:
        try:
            pidx = os.getpid()
            pids = psutil.pids()
            process_name_list=[psutil.Process(pid).name() for pid in pids]
            x=process_name_list.count("daxformat_one_new1.exe")
            if x>2:
                sg.popup_auto_close('Program already exists', auto_close_duration=2)
                psutil.Process(pidx).kill()
            break
        except:
            n=n+1
            if n==3:
                sg.popup_auto_close('start error', auto_close_duration=5)
                os._exit(0)


def Ctrl_X(key):

    if key == "v":
        win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
        win32api.keybd_event(86, 0, 0, 0)  # v键位码是86
        win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    elif key == "a":
        win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
        win32api.keybd_event(65, 0, 0, 0)  # a键位码是65
        win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    elif key == "c":
        win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
        win32api.keybd_event(67, 0, 0, 0)  # c键位码是67
        win32api.keybd_event(67, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    elif key == "ac":
        win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
        sleep(0.3)
        win32api.keybd_event(65, 0, 0, 0)  # a键位码是65
        sleep(0.2)
        win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        sleep(0.3)
        win32api.keybd_event(67, 0, 0, 0)  # c键位码是67
        win32api.keybd_event(67, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        sleep(0.6)
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
        sleep(0.6)
    elif key == "av":
        win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
        sleep(0.3)
        win32api.keybd_event(65, 0, 0, 0)  # a键位码是65
        sleep(0.2)
        win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        sleep(0.3)
        win32api.keybd_event(86, 0, 0, 0)  # v键位码是86
        sleep(0.2)
        win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        sleep(0.6)
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
        sleep(0.6)

# 键盘按压
def on_press(key):
    pass

def get_dax():
    Ctrl_X("ac")  # 全选复制
    # urllib3.contrib.pyopenssl.inject_into_urllib3()
    # requests.packages.urllib3.disable_warnings()
    # requests.adapters.DEFAULT_RETRIES = 3
    # session = requests.session()
    # session.keep_alive = False
    user_agent_list = [
        'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95',
        'Safari/537.36 OPR/26.0.1656.60',
        'Opera/8.0 (Windows NT 5.1; U; en)',
        'Mozilla/5.0 (Windows NT 5.1; U; en; rv:1.8.1) Gecko/20061208 Firefox/2.0.0 Opera 9.50',
        'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:34.0) Gecko/20100101 Firefox/34.0',
        'Mozilla/5.0 (X11; U; Linux x86_64; zh-CN; rv:1.9.2.10) Gecko/20100922 Ubuntu/10.10 '
        '(maverick) Firefox/3.6.10',
        'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/39.0.2171.71 Safari/537.36',
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 '
        '(KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
    ]
    UserAgent = random.choice(user_agent_list)
    headers = {'User-Agent': UserAgent}               #, 'Connection': 'close'}
    url = 'https://daxtest02.azurewebsites.net/api/daxformatter/daxtokenformat/'
    # url = 'https://daxformatter.azurewebsites.net/api/DaxTokenFormat/'
    input_win_title0 = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    if  "Tabular Editor" in input_win_title0:
        daxtext="x="+clip.paste()
    else:
        daxtext=clip.paste()
    x = {'Dax': daxtext, 'ListSeparator': ',', 'DecimalSeparator': '.', 'MaxLineLenght': 0}
    r = requests.post(url, data=x, timeout=25, headers=headers)
    dax_dict = r.json()
    print(dax_dict)
    d1 = dax_dict["formatted"]
    if len(d1) == 0:
        # print("DAX公式错误")
        AutoCloseMessageBoxW("DAX error", 'DAX error')
    else:
        result = ""
        for i, v in enumerate(d1):
            if i > 0:
                result = result + '\r\n'
            for x in v:
                result = result + x["string"]
        # print(result)
        if  "Tabular Editor" in input_win_title0:
            result=result[4:]
        else:
            result=result
        clip.copy(result)  # 结果存储到剪切板
        text = "measure : " + d1[0][0]["string"][0:-1] + " --------has been formatted"
        print(text)

        Ctrl_X("av")



def on_release(key):
    global all_key
    input_win_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    win_title_list = ["Power BI Desktop", "记事本", "Microsoft Visual Studio", "PyCharm","Tabular Editor"]
    if any(win_title_name in input_win_title for win_title_name in win_title_list) and str(key)=="Key.f8":
        print(str(key))
        time_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print("start:" + time_str)
        get_dax()


    # if key == Key.f12:
    if str(key) =="Key.f9":
        return False


def start_listen():
    # pidx = os.getpid()
    with Listener(on_press=None, on_release=on_release) as listenerX:
        listenerX.join()


def daxformat_main():
    global all_key
    all_key = []
    kill_process_one()  #保证运行一个
    # sg.popup_auto_close('start', auto_close_duration=0.6)
    start_listen()


if __name__ == '__main__':
    # F8 触发格式化
    # F9 退出监听
    daxformat_main()


