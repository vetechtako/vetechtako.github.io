import pyautogui
import time
import threading
import tkinter as tk
from PIL import Image, ImageDraw
import pystray
import sys

INTERVAL = 30  # 每隔幾秒模擬一次滑鼠

# === 主程式狀態 ===
is_running = True

# === 滑鼠模擬行為 ===
def simulate_mouse_activity():
    while is_running:
        current_x, current_y = pyautogui.position()
        pyautogui.moveTo(current_x + 1, current_y)
        pyautogui.moveTo(current_x, current_y)
        log_message(f"滑鼠模擬：({current_x+1}, {current_y})")
        time.sleep(INTERVAL)

# === 記錄訊息到主視窗 ===
def log_message(message):
    if text_area:
        text_area.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        text_area.see(tk.END)

# === 顯示主視窗 ===
def show_window(icon=None, item=None):
    if icon:
        icon.stop()  # 關閉系統欄
    root.after(0, root.deiconify)

# === 隱藏到任務欄（包括按X）===
def hide_window():
    root.withdraw()
    create_tray_icon()

# === 結束程式 ===
def exit_app(icon=None, item=None):
    global is_running
    is_running = False
    if icon:
        icon.stop()
    root.quit()
    sys.exit()

# === 創建任務欄圖示 ===
def create_image():
    image = Image.new("RGB", (64, 64), "white")
    d = ImageDraw.Draw(image)
    d.rectangle((16, 16, 48, 48), fill="green")
    return image

# === 建立任務欄選單 ===
def create_tray_icon():
    image = create_image()
    menu = pystray.Menu(
        pystray.MenuItem("開啟視窗", show_window),
        pystray.MenuItem("結束程式", exit_app)
    )
    icon = pystray.Icon("mouse_simulator", image, "滑鼠移動模擬中", menu)
    threading.Thread(target=icon.run, daemon=True).start()

# === 主視窗 ===
root = tk.Tk()
root.title("滑鼠移動模擬器")
root.geometry("420x340")

# 說明標籤
label = tk.Label(root, text=(
    "滑鼠每30秒自動移動1像素，保持滑鼠活動，\n"
	"防止因滑鼠停滯而被Webex專心度追蹤判定為異常狀態。\n"
    "--\n"
    "Webex啟動前請先按X關閉視窗，\n"
	"程式將最小化至工作列右側小圖示。\n"
    "--\n"
    "工作列小圖示按右鍵選單結束程式。"
), justify="left", wraplength=400)
label.pack(padx=10, pady=5)

# 紀錄文字區域
text_area = tk.Text(root, height=15)
text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

# 攔截關閉按鈕 (X)
root.protocol("WM_DELETE_WINDOW", hide_window)

# 開始滑鼠模擬執行緒
threading.Thread(target=simulate_mouse_activity, daemon=True).start()

# 啟動主視窗
root.mainloop()
