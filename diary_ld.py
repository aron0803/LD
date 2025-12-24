import tkinter as tk
from tkinter import filedialog, ttk
import cv2
import numpy as np
import time
import threading
import os
import subprocess
from PIL import Image
from configparser import ConfigParser

# ==================== LdPlayer 控制類 ====================
class LdPlayerManager:
    def __init__(self, ld_path, screenshot_dir, index):
        self.ld_path = ld_path
        self.screenshot_dir = screenshot_dir
        self.index = index

    def set_ld_path(self, path):
        self.ld_path = path

    def set_screenshot_dir(self, path):
        self.screenshot_dir = path

    def set_index(self, idx):
        self.index = idx

    def click(self, x, y):
##        print([self.ld_path, "-s", str(self.index),  "input", "tap", str(x), str(y)])
        subprocess.run([self.ld_path, "-s", str(self.index),  "input", "tap", str(x), str(y)], creationflags=subprocess.CREATE_NO_WINDOW)

    def screencap(self):
        filepath =  f"/sdcard/Pictures/cap_{self.index}.png"
##        print([self.ld_path, "-s", str(self.index), "screencap", filepath])
        subprocess.run([self.ld_path, "-s", str(self.index), "screencap",  filepath], creationflags=subprocess.CREATE_NO_WINDOW)
        return Image.open(os.path.join(self.screenshot_dir, f"cap_{self.index}.png"))

# ==================== 匹配函式 ====================
def match(cap, pic2_path, threshold=0.98):
    try:
        # Ensure both images are RGB (drop alpha if exists) to avoid shape mismatch or color errors
        img1 = np.array(cap.convert('RGB'))
        pic2 = Image.open(pic2_path).convert('RGB')
        img2 = np.array(pic2)
        
        # Convert to BGR for OpenCV
        img1_bgr = cv2.cvtColor(img1, cv2.COLOR_RGB2BGR)
        img2_bgr = cv2.cvtColor(img2, cv2.COLOR_RGB2BGR)
        
        result = cv2.matchTemplate(img1_bgr, img2_bgr, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        
        # Debug: Print similarity score to console
        print(f"比對 {os.path.basename(pic2_path)}: 相似度={max_val:.4f} (門檻={threshold})")
        
        return max_loc if max_val >= threshold else (0, 0)
    except Exception as e:
        print(f"比對錯誤: {e}")
        return (0, 0)

# ==================== Config 處理 ====================
config_path = "config.ini"
config = ConfigParser()
if os.path.exists(config_path):
    config.read(config_path)
    ld_path = config.get("Settings", "ld_path", fallback="D:\\LDPlayer\\LDPlayer9\\ld.exe")
    screenshot_dir = config.get("Settings", "screenshot_dir", fallback="D:\\Screenshots")
    ld_index = config.getint("Settings", "ld_index", fallback=0)
else:
    ld_path = "D:\\LDPlayer\\LDPlayer9\\ld.exe"
    screenshot_dir = "D:\\Screenshots"
    ld_index = 0

ld_manager = LdPlayerManager(ld_path, screenshot_dir, ld_index)

# ==================== GUI ====================
# ==================== GUI ====================
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

app = ttk.Window(themename="cosmo")
app.title("LDPlayer 自動化腳本")
# app.geometry("600x550") # Removed fixed size for dynamic resizing
app.minsize(600, 500) # Set minimum size instead

status_var = tk.StringVar(value="狀態：準備就緒")

# 設定檔 GUI 互動
ld_path_var = tk.StringVar(value=ld_path)
screenshot_dir_var = tk.StringVar(value=screenshot_dir)
index_var = tk.IntVar(value=ld_index)

png_vars = []
running = False

def select_ld_path():
    path = filedialog.askopenfilename(title="選擇 ld.exe 路徑", filetypes=[("ld.exe", "*.exe")])
    if path:
        ld_path_var.set(path)
        ld_manager.set_ld_path(path)

def select_screenshot_dir():
    path = filedialog.askdirectory(title="選擇截圖資料夾")
    if path:
        screenshot_dir_var.set(path)
        ld_manager.set_screenshot_dir(path)

def update_status(msg):
    status_var.set(msg)
    status_label.update_idletasks()

def load_png_files():
    # 清除舊的 checkbox (如果有)
    for widget in png_frame.winfo_children():
        widget.destroy()
    png_vars.clear()

    png_directory = os.path.join(os.getcwd(), 'pic')
    if not os.path.exists(png_directory):
        os.makedirs(png_directory)
    png_files = [f for f in os.listdir(png_directory) if f.endswith('.png')]
    
    col_count = 0
    row_count = 0
    max_cols = 3
    
    for file in png_files:
        var = tk.BooleanVar()
        chk = ttk.Checkbutton(png_frame, text=file, variable=var, bootstyle="round-toggle")
        chk.grid(row=row_count, column=col_count, sticky='w', padx=10, pady=5)
        png_vars.append((file, var))
        
        col_count += 1
        if col_count >= max_cols:
            col_count = 0
            row_count += 1

def run_script():
    global running
    ld_manager.set_index(index_var.get())
    update_status(f"執行模擬器 Index: {index_var.get()}")
    while running:
        selected_pngs = [var[0] for var in png_vars if var[1].get()]
        ld_manager.click(400, 380)
        time.sleep(1)
        cap = ld_manager.screencap()
        x, y = match(cap, "diary.png", 0.90)
        if x:
            update_status("點日記本")
            ld_manager.click(x, y)
            time.sleep(2)
        cap = ld_manager.screencap()
        x, y = match(cap, "ball.png", 0.98)
        if x:
            update_status("點神秘珠子")
            ld_manager.click(x, y)
            time.sleep(2)
        cap = ld_manager.screencap()
        x, y = match(cap, "check1.png", 0.98)
        if x:
            update_status("確認")
            ld_manager.click(x + 240, y + 130)
            time.sleep(2)
            cap = ld_manager.screencap()
            x, y = match(cap, "check2.png", 0.98)
            if x:
                while running:
                    for png_file in selected_pngs:
                        png_path = os.path.join(os.getcwd(), 'pic', png_file)
                        update_status(f"判斷 {png_file}")
                        x, y = match(cap, png_path, 0.98)
                        if x:
                            update_status(f"製作 {png_file}")
                            for _ in range(6):
                                ld_manager.click(405, 357)
                                time.sleep(1)
                            break
                    if x:
                        break
                    update_status("變更")
                    ld_manager.click(470, 300)
                    time.sleep(1.5)
                    cap = ld_manager.screencap()

def start_script():
    global running
    if not running:
        running = True
        threading.Thread(target=run_script, daemon=True).start()

def stop_script():
    global running
    running = False
    update_status("狀態：已停止")

def save_config():
    config['Settings'] = {
        'ld_path': ld_path_var.get(),
        'screenshot_dir': screenshot_dir_var.get(),
        'ld_index': index_var.get()
    }
    with open(config_path, 'w') as configfile:
        config.write(configfile)

# ==================== GUI Layout Construction ====================

# Main Container with padding
main_frame = ttk.Frame(app, padding="20")
main_frame.pack(fill=BOTH, expand=YES)

# --- Settings Group ---
settings_frame = ttk.Labelframe(main_frame, text="設定 (Settings)", padding="15")
settings_frame.pack(fill=X, pady=(0, 15))

# Grid layout for settings
settings_frame.columnconfigure(1, weight=1)

# LDPlayer Path
ttk.Label(settings_frame, text="LDPlayer 路徑:").grid(row=0, column=0, sticky=W, pady=5)
ttk.Entry(settings_frame, textvariable=ld_path_var).grid(row=0, column=1, sticky=EW, padx=10, pady=5)
ttk.Button(settings_frame, text="瀏覽", command=select_ld_path, bootstyle="outline").grid(row=0, column=2, padx=5, pady=5)

# Screenshot Dir
ttk.Label(settings_frame, text="截圖資料夾:").grid(row=1, column=0, sticky=W, pady=5)
ttk.Entry(settings_frame, textvariable=screenshot_dir_var).grid(row=1, column=1, sticky=EW, padx=10, pady=5)
ttk.Button(settings_frame, text="瀏覽", command=select_screenshot_dir, bootstyle="outline").grid(row=1, column=2, padx=5, pady=5)

# Index
ttk.Label(settings_frame, text="模擬器 Index (0~30):").grid(row=2, column=0, sticky=W, pady=5)
ttk.Spinbox(settings_frame, from_=0, to=30, textvariable=index_var, width=5).grid(row=2, column=1, sticky=W, padx=10, pady=5)

# --- Target Images Group ---
target_frame = ttk.Labelframe(main_frame, text="目標圖片 (Target Images)", padding="15")
target_frame.pack(fill=BOTH, expand=YES, pady=(0, 15))

# Container for checkboxes (png_frame)
png_frame = ttk.Frame(target_frame)
png_frame.pack(fill=BOTH, expand=YES)

load_png_files()

# --- Control Group ---
control_frame = ttk.Frame(main_frame)
control_frame.pack(fill=X, pady=(0, 10))

# Center the buttons
button_container = ttk.Frame(control_frame)
button_container.pack(anchor=CENTER)

ttk.Button(button_container, text="開始執行 (Start)", command=start_script, bootstyle="success", width=15).pack(side=LEFT, padx=5)
ttk.Button(button_container, text="停止執行 (Stop)", command=stop_script, bootstyle="danger", width=15).pack(side=LEFT, padx=5)
ttk.Button(button_container, text="儲存設定 (Save)", command=save_config, bootstyle="info-outline", width=15).pack(side=LEFT, padx=5)

# --- Status Bar ---
status_label = ttk.Label(app, textvariable=status_var, bootstyle="inverse-secondary", anchor=W, padding=(10, 5))
status_label.pack(side=BOTTOM, fill=X)

app.mainloop()
