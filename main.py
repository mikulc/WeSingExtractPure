import pygetwindow as gw
import win32gui
import win32con
import pyautogui
import time
import os
import shutil
import pyperclip
from datetime import datetime


class WeSingExtractor:
    def __init__(self):
        # 获取窗口信息
        self.window = gw.getWindowsWithTitle('全民K歌')[0]
        self.x = self.window.left
        self.y = self.window.top
        self.width = self.window.width
        self.height = self.window.height
        self.song_name = None
        self.current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 伴奏保存路径
        self.output_path = r'D:\WeSingCache\WeSingDL\Output'
        # Temp路径
        self.temp_path = r'D:\WeSingCache\WeSingDL\Temp'
        # Res路径
        self.res_path = r'D:\WeSingCache\WeSingDL\Res'

        self.positions = {
            'search_box': (1.3, 25),
            'search_btn': (1.17, 25),
            'fir_k_btn': (1.13, 2.9),
            'exit_btn': (1.035, 32),
            'exit_cfm': (1.9, 1.8),
            'exit_btn_': (1.08, 32),
            'exit_cfm_': (1.9, 1.8),
        }

    def get_position(self, position_name):
        x, y = self.positions[position_name]
        return self.x + self.width / x, self.y + self.height / y

    # 置顶窗口
    def set_window_top(self):
        # 获取窗口句柄
        hwnd = self.window._hWnd
        # 如果窗口最小化，将其恢复
        if win32gui.IsIconic(hwnd):  # 判断窗口是否最小化
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        # 激活并置顶窗口
        win32gui.SetForegroundWindow(hwnd)
        print("-" * 30)
        print(self.current_time)
        print("置顶窗口成功！")
        time.sleep(1)

    def search_song(self, song_name):
        print("正在搜索歌曲...")
        pyautogui.doubleClick(self.get_position("search_box"))
        time.sleep(0.5)
        pyautogui.write(f"{song_name}", interval=0.1)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'c')
        self.song_name = pyperclip.paste().replace(' ', ' - ')
        print(f"歌曲名称：{self.song_name}")

        for song_ in os.listdir(self.output_path):
            if self.song_name == song_.split('.')[0]:
                print("歌曲已存在！")
                return False
            else:
                continue
        pyautogui.click(self.get_position("search_btn"))
        return True


    def click_k_btn(self, ):
        pyautogui.click(self.get_position("fir_k_btn"))

    def exit_click(self):
        pyautogui.click(self.get_position("exit_btn"))
        time.sleep(1)
        pyautogui.click(self.get_position("exit_cfm"))

    def exit_click_(self):
        pyautogui.click(self.get_position("exit_btn_"))
        time.sleep(1)
        pyautogui.click(self.get_position("exit_cfm_"))

    def modify_file(self):
        newest_folder = self.find_newest_folder(self.res_path)
        if newest_folder:
            self.process_folder(newest_folder)
        else:
            print("没有找到任何文件夹或无法进入文件夹")

    def process_folder(self, folder_path):
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path) and item.endswith('.note'):
                try:
                    os.remove(item_path)
                    print(f"已删除文件: {item_path}")

                    folder_name = item
                    os.makedirs(os.path.join(folder_path, folder_name), exist_ok=True)
                    print(f"已创建文件夹: {folder_name}")
                except Exception as e:
                    print(f"无法删除文件或创建文件夹: {e}")

    def find_newest_folder(self, dir_path):
        if not os.path.isdir(dir_path):
            print(f"提供的路径 {dir_path} 不是一个有效的目录")
            return None

        newest_folder = None
        newest_time = 0

        for item in os.listdir(dir_path):
            item_path = os.path.join(dir_path, item)
            if os.path.isdir(item_path):
                mod_time = os.path.getmtime(item_path)
                if mod_time > newest_time:
                    newest_time = mod_time
                    newest_folder = item_path
        return newest_folder

    def save_song(self):
        print("正在寻找歌曲缓存文件...")
        for file in os.listdir(self.temp_path):
            if file.endswith('.wav'):
                new_name = os.path.join(self.output_path, self.song_name + '.wav')
                try:
                    shutil.copy(os.path.join(self.temp_path, file), new_name)
                    print("保存歌曲成功！")
                except FileNotFoundError as e:
                    print("没有找到文件！")

    def start_script(self, song_):
        # 将窗口置顶
        self.set_window_top()
        # 搜索歌曲
        if self.search_song(f'{song_}'):
            # 点击K歌
            time.sleep(2)
            self.click_k_btn()
            # 因为歌曲要加载，延迟时间可设置久点，默认10s
            print("正在进入K歌界面,等待歌曲加载...\n")
            time.sleep(10)
            # 退出K歌
            self.exit_click_()
            # 修改文件
            self.modify_file()
            # 点击K歌
            self.click_k_btn()
            print("请耐心等待歌曲播放完成...")
            # 等待播放完一首歌时间，默认350秒
            time.sleep(350)
            self.exit_click_()
            print("歌曲播放完毕！")
            # 保存歌曲
            self.save_song()
            # 退出K歌
            self.exit_click_()
            print("伴奏保存完成！")


if __name__ == '__main__':
    we_sing = WeSingExtractor()

    with open('./SongList.txt') as f:
        for song in f:
            we_sing.start_script(song)
