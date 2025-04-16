import time
import threading
import pygetwindow as gw
import psutil
import win32process
from pystray import Icon, MenuItem, Menu
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import scrolledtext

# 設定 Webex 相關程式名稱清單（小寫）
target_processes = [
    "webex.exe",
    "webexhost.exe",
    "webexmta.exe",
    "atmgr.exe",
    "ciscocollabhost.exe"
]

# 設定 Webex 相關視窗標題關鍵字（小寫）
target_keywords = ["webex", "meeting center", "cisco"]

# 常見瀏覽器執行檔名（小寫）
browsers = ["chrome.exe", "msedge.exe", "firefox.exe", "opera.exe", "brave.exe"]

# 建立任務欄圖示
def create_image():
    image = Image.new('RGB', (64, 64), (50, 150, 255))
    draw = ImageDraw.Draw(image)
    draw.rectangle((16, 16, 48, 48), fill=(255, 255, 255))
    return image

# 寫入 log 檔案
def log_to_file(message):
    with open("window_log.txt", "a", encoding="utf-8") as f:
        f.write(message + "\n")

# 取得目前前景視窗的執行檔名與標題
def get_active_process_name():
    active_window = gw.getActiveWindow()
    if active_window:
        hwnd = active_window._hWnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return process.name(), active_window.title
    else:
        return None, None

# 更新前景程式資訊
def update_window_info():
    global last_process_name
    now = time.strftime('%Y-%m-%d %H:%M:%S')
    process_name, window_title = get_active_process_name()

    if process_name:
        process_name_lower = process_name.lower()
        window_title_lower = window_title.lower()

        is_webex_title = any(keyword in window_title_lower for keyword in target_keywords)
        is_webex_process = process_name_lower in target_processes
        is_browser = process_name_lower in browsers

        # 判斷是否為 Webex
        if is_webex_process or (is_webex_title and not is_browser):
            log_message = f"{now} - ✔️目前前景程式：{process_name} | 視窗：{window_title}"
            color = "green"
        else:
            log_message = f"{now} - ❌目前前景程式：{process_name} | 視窗：{window_title}"
            color = "red"

        if process_name != last_process_name:
            log_to_file(log_message)
            text_area.insert(tk.END, log_message + "\n", color)
            text_area.see(tk.END)
            last_process_name = process_name
    else:
        log_message = f"{now} - ❌沒有前景視窗"
        log_to_file(log_message)
        text_area.insert(tk.END, log_message + "\n", "red")
        text_area.see(tk.END)

    root.after(1000, update_window_info)

# 任務欄選單開啟視窗
def show_window(icon, item):
    root.deiconify()

# 任務欄選單結束程式
def on_exit(icon, item):
    icon.stop()
    root.destroy()

# 建立任務欄選單
menu = Menu(
    MenuItem('開啟視窗', show_window),
    MenuItem('結束程式', on_exit)
)
icon = Icon("window_monitor", create_image(), "前景視窗記錄中", menu)

# 任務欄執行緒
def run_tray():
    icon.run()

# 建立 Tkinter 主視窗
root = tk.Tk()
root.title("前景視窗記錄器")
root.geometry("680x420")
root.protocol("WM_DELETE_WINDOW", lambda: root.withdraw())

# 說明標籤
label = tk.Label(root, text=(
    "自動紀錄前景視窗變化並生成window_log.txt記錄檔。\n"
    "--\n"
    "Webex啟動前請先按X關閉視窗，\n"
	"程式將最小化至工作列右側小圖示。\n"
    "--\n"
    "工作列小圖示按右鍵選單結束程式。"
), justify="left", wraplength=400)
label.pack(padx=10, pady=5)

# 建立滾動文字區
text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("Arial", 12))
text_area.pack(expand=True, fill='both', padx=10, pady=10)

# 定義顏色樣式
text_area.tag_config("green", foreground="green")
text_area.tag_config("red", foreground="red")

# 初始化變數
last_process_name = ""

# 啟動監控與任務欄
root.after(1000, update_window_info)
threading.Thread(target=run_tray, daemon=True).start()

# 執行 Tkinter 事件迴圈
root.mainloop()