# coding: UTF-8

import os
import re
import time
import json
import datetime
import keyboard
import win32gui
import win32process
import win32api
import wmi

DEBUG    = 1
WAITTIME = 1 # seconds

ODIR = "output"
OJSON_FILE = f"{ODIR}/output.json"
OCSV_FILE  = f"{ODIR}/output.csv"

class CalcKousu:
    def __init__(self):
        self.start = 0
        self.pre_app_name           = ""
        self.pre_active_window_name = ""
        self.app_name               = ""
        self.active_window_name     = ""
        self.c = wmi.WMI()

        self.kousu_dict = {}
        if (os.path.exists(OJSON_FILE)):
            with open(OJSON_FILE, "r") as f:
                self.kousu_dict = json.load(f)

        self.keyword_dict = {}
        if (os.path.exists("keyword.json")):
            with open("keyword.json", "r", encoding="utf-8") as f:
                self.keyword_dict = json.loads(f.read())

        self.terminate_type = [
            "PRESS_CTRL+C",
            "BREAK_EVENT",
            "PRESS_CROSS_BUTTON",
            "UNKNOWN_TYPE_3",
            "UNKNOWN_TYPE_4",
            "LOFOFF_EVENT",
            "SHUTDOWN_EVENT"
        ]
        win32api.SetConsoleCtrlHandler(self.on_exit_with_cross_button, True)

    # アクティブウィンドウ名とアプリ名を取得
    def get_active_window_and_app_name(self):
        try:
            foreground_window       = win32gui.GetForegroundWindow()
            # アクティブウィンドウ名取得
            self.active_window_name = win32gui.GetWindowText(foreground_window)
            # アクティブアプリ名取得
            _, pid = win32process.GetWindowThreadProcessId(foreground_window)
            for p in self.c.query(f"SELECT Name FROM Win32_Process WHERE ProcessId = {pid}"):
                self.app_name = p.Name
                break
            self.update_app_name()
        except:
            print("Error.")
            self.active_window_name = "Error"
            self.app_name = "Error"

        if (DEBUG == 1):
            print(f"{self.app_name}: {self.active_window_name}")

    # 指定のワードがアクティブウィンドウ名に入ってたら、アプリ名を自分指定の名前に変更する
    # 各種指定のワードは「keyword.jsonc」に記載する
    # 正規表現にも対応
    def update_app_name(self):
        for key, value in self.keyword_dict.items():
            key_search = re.search(key, self.active_window_name)
            if (key_search != None):    # 正規表現で文字列を検索し、matchしたら一定の工数を付けたい場合
                self.app_name = value
                break
            elif (self.active_window_name.find(key) != -1): # keyを含む場合に､､､
                re_search = re.search(value, self.active_window_name)
                if (re_search != None): # 工数名は文字列によって柔軟い変更したい場合
                    self.app_name = re_search.group()
                else:                   # 工数名は自分で決めた文字列で設定する場合
                    self.app_name = value 
                break
            else:
                pass

    # 以前の状態を更新する
    def update_pre_state(self):
        self.start                  = time.time()
        self.pre_app_name           = self.app_name
        self.pre_active_window_name = self.active_window_name

    # 結果格納している辞書を更新
    def update_dict(self):
        proc_time = time.time() - self.start
        # 1時間でどれくらい
        dt_now = datetime.datetime.now().strftime("%Y%m%d_%H")
        self.add_dict(dt_now, proc_time)
        # その日でどれくらい
        dt_now = datetime.datetime.now().strftime("%Y%m%d")
        self.add_dict(dt_now, proc_time)

    def add_dict(self, dt_now, proc_time):
        if dt_now in self.kousu_dict:
            if self.pre_app_name in self.kousu_dict[dt_now]:
                self.kousu_dict[dt_now][self.pre_app_name] += proc_time
            else:
                self.kousu_dict[dt_now][self.pre_app_name] = proc_time
        else:
            self.kousu_dict[dt_now] = {}
            self.kousu_dict[dt_now][self.pre_app_name] = proc_time

    # 結果出力
    def SaveResult(self):
        # 出力フォルダがない場合は作成する
        if not(os.path.exists(ODIR)):
            os.makedirs(ODIR)

        # jsonで出力
        with open(OJSON_FILE, "w", encoding="utf-8") as f:
            json.dump(self.kousu_dict, f, indent=2)

        # 一応CSVとしても出力
        with open(OCSV_FILE, "w", encoding="utf-8") as f:
            f.write("day or hour, date, app_name, time(sec), time(min), time(hour)\n")
            for key in self.kousu_dict:
                id = "day" if (key.find("_") == -1) else "hour"
                for k, v in self.kousu_dict[key].items():
                    f.write(f"{id}, {key}, {k}, {v}, {v/60}, {v/(60*60)}\n")

    # 処理記述
    def run(self):
        while True:
            if keyboard.is_pressed("ctrl+shift+alt"):
                self.update_dict()
                self.SaveResult()
                break

            time.sleep(WAITTIME)

            self.get_active_window_and_app_name()

            if (self.pre_app_name == ""):
                self.update_pre_state()
            elif (self.pre_app_name != self.app_name):
                self.update_dict()
                self.update_pre_state()
            else:
                pass

    # 何かしらで強制終了されたときに実行する関数
    def on_exit_with_cross_button(self, signal_type):
        print(self.terminate_type[signal_type])
        self.update_dict()
        self.SaveResult()

def main():
    calc_kousu = CalcKousu()
    calc_kousu.run()

if __name__ == "__main__":
    main()