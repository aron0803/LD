import os
import sys
import time
import threading
import subprocess
import configparser
import cv2
import numpy as np
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
try:
    from win32com.shell import shell, shellcon
except ImportError:
    shell = None

def cv2_imread_unicode(path):
    try:
        # Support Chinese characters in path
        img = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
        return img
    except Exception as e:
        print(f"Error reading {path}: {e}")
        return None

class LdPlayerManager:
    def __init__(self, ld_path, screenshot_dir, index=0):
        self.ld_path = ld_path
        self.screenshot_dir = screenshot_dir
        self.index = index

    def run_command(self, *args):
        try:
            cmd = [self.ld_path, "-s", str(self.index)] + list(args)
            subprocess.run(cmd, check=False, creationflags=subprocess.CREATE_NO_WINDOW)
        except Exception as e:
            print(f"RunCommand Error: {e}")

    def click(self, x, y):
        self.run_command("input", "tap", str(x), str(y))

    def screencap(self):
        filename = f"trace_cap_{self.index}.png"
        local_path = os.path.join(self.screenshot_dir, filename)
        remote_path = f"/sdcard/Pictures/{filename}"
        
        # Remove old file if exists
        if os.path.exists(local_path):
            try: os.remove(local_path)
            except: pass

        try:
            self.run_command("screencap", remote_path)
            # Wait for file to appear (up to 5 seconds)
            for _ in range(50):
                if os.path.exists(local_path):
                    # Try to read it to ensure it's not being written
                    img = cv2.imread(local_path)
                    if img is not None:
                        return img
                time.sleep(0.1)
            return None
        except Exception as e:
            print(f"Screencap Error: {e}")
            return None

class ImageMatcher:
    def __init__(self):
        self.template_cache = {}

    def load_templates(self, trace_dir):
        self.template_cache.clear()
        if not os.path.exists(trace_dir):
            return
        
        for file in os.listdir(trace_dir):
            if file.lower().endswith(".png"):
                path = os.path.join(trace_dir, file)
                img = cv2_imread_unicode(path)
                if img is not None:
                    self.template_cache[file] = img

    def find_image(self, source_img, template_name, threshold=0.9, search_region=None):
        if template_name not in self.template_cache:
            return None
        
        template = self.template_cache[template_name]
        
        # Apply search region if provided (x, y, w, h)
        offset_x, offset_y = 0, 0
        if search_region:
            x, y, w, h = search_region
            if x >= 0 and y >= 0 and x + w <= source_img.shape[1] and y + h <= source_img.shape[0]:
                source_img = source_img[y:y+h, x:x+w]
                offset_x, offset_y = x, y

        res = cv2.matchTemplate(source_img, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

        if max_val >= threshold:
            return {
                "success": True,
                "location": (max_loc[0] + offset_x, max_loc[1] + offset_y),
                "score": max_val
            }
        return {"success": False}

class TraceLDApp:
    def __init__(self, root):
        self.root = root
        self.root.title("LDPlayer Trace 腳本 (Python)")
        self.root.geometry("550x750")
        
        # Set icon if exists
        if os.path.exists("favicon.ico"):
            try:
                self.root.iconbitmap("favicon.ico")
            except:
                pass
        
        self.style = ttk.Style(theme="superhero")
        
        self.is_running = False
        self.worker_thread = None
        self.ld_manager = None
        self.matcher = ImageMatcher()
        # For Nuitka onefile: use the original exe location, not temp extraction dir
        # Check if running as compiled exe by looking at sys.argv[0]
        exe_path = sys.argv[0] if sys.argv else sys.executable
        if exe_path.lower().endswith('.exe'):
            # Running as compiled exe - use argv[0] which has the real path
            self.script_dir = os.path.dirname(os.path.abspath(exe_path))
        elif getattr(sys, 'frozen', False):
            self.script_dir = os.path.dirname(sys.executable)
        else:
            self.script_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.trace_dir = os.path.join(self.script_dir, "trace")
        self.config_path = os.path.join(self.script_dir, "trace.ini")
        
        self.setup_ui()
        self.load_config()
        self.load_images()
        self.refresh_emulators()
        
        # Add trace to LD path for auto-detection
        self.txt_ld_path.bind("<FocusOut>", lambda e: self.on_ld_path_changed())
        self.txt_ld_path.bind("<Return>", lambda e: self.on_ld_path_changed())

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=BOTH, expand=YES)

        # --- Settings Group ---
        settings_group = ttk.Labelframe(main_frame, text="設定 (Settings)", padding=10)
        settings_group.pack(fill=X, pady=5)

        # LD Path
        path_frame = ttk.Frame(settings_group)
        path_frame.pack(fill=X, pady=2)
        ttk.Label(path_frame, text="LDPlayer 路徑:", width=15).pack(side=LEFT)
        self.txt_ld_path = ttk.Entry(path_frame)
        self.txt_ld_path.pack(side=LEFT, fill=X, expand=YES, padx=5)
        ttk.Button(path_frame, text="瀏覽", command=self.browse_ld_path, width=8).pack(side=LEFT)

        # Screenshot Dir
        scr_frame = ttk.Frame(settings_group)
        scr_frame.pack(fill=X, pady=2)
        ttk.Label(scr_frame, text="截圖資料夾:", width=15).pack(side=LEFT)
        self.txt_screenshot_dir = ttk.Entry(scr_frame)
        self.txt_screenshot_dir.pack(side=LEFT, fill=X, expand=YES, padx=5)
        ttk.Button(scr_frame, text="瀏覽", command=self.browse_screenshot_dir, width=8).pack(side=LEFT)

        # Wait Seconds
        wait_frame = ttk.Frame(settings_group)
        wait_frame.pack(fill=X, pady=2)
        ttk.Label(wait_frame, text="比對間隔 (秒):", width=15).pack(side=LEFT)
        self.num_wait_seconds = ttk.Spinbox(wait_frame, from_=0.1, to=60.0, increment=0.5, width=10)
        self.num_wait_seconds.set(1.0)
        self.num_wait_seconds.pack(side=LEFT)

        # --- Lists Container ---
        lists_frame = ttk.Frame(main_frame)
        lists_frame.pack(fill=BOTH, expand=YES, pady=5)

        # Emulator List (Left) - same width as image list
        emu_group = ttk.Labelframe(lists_frame, text="模擬器選擇", padding=5)
        emu_group.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 5))
        
        self.emu_list_frame = ttk.Frame(emu_group)
        self.emu_list_frame.pack(fill=BOTH, expand=YES)
        
        # Scrollable canvas for checkboxes
        self.emu_canvas = tk.Canvas(self.emu_list_frame, highlightthickness=0, width=150)
        self.emu_scrollbar = ttk.Scrollbar(self.emu_list_frame, orient=VERTICAL, command=self.emu_canvas.yview)
        self.emu_scrollable_frame = ttk.Frame(self.emu_canvas)

        self.emu_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.emu_canvas.configure(scrollregion=self.emu_canvas.bbox("all"))
        )

        self.emu_canvas.create_window((0, 0), window=self.emu_scrollable_frame, anchor="nw")
        self.emu_canvas.configure(yscrollcommand=self.emu_scrollbar.set)

        self.emu_canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        self.emu_scrollbar.pack(side=RIGHT, fill=Y)

        emu_btn_frame = ttk.Frame(emu_group)
        emu_btn_frame.pack(fill=X, pady=5)
        ttk.Button(emu_btn_frame, text="全選", command=self.select_all_emus, width=6).pack(side=LEFT, padx=1)
        ttk.Button(emu_btn_frame, text="全不選", command=self.deselect_all_emus, width=6).pack(side=LEFT, padx=1)
        ttk.Button(emu_btn_frame, text="整理", command=self.refresh_emulators, width=6).pack(side=LEFT, padx=1)

        # Image List (Right) - same width as emulator list
        img_group = ttk.Labelframe(lists_frame, text="比對清單 (優先順序)", padding=5)
        img_group.pack(side=LEFT, fill=BOTH, expand=YES, padx=(5, 0))
        
        img_list_container = ttk.Frame(img_group)
        img_list_container.pack(fill=BOTH, expand=YES)
        
        self.lst_images = tk.Listbox(img_list_container, selectmode=tk.SINGLE, font=("Microsoft JhengHei", 10))
        self.lst_images.pack(side=LEFT, fill=BOTH, expand=YES)
        
        # Drag and Drop bindings
        self.lst_images.bind('<Button-1>', self.on_drag_start)
        self.lst_images.bind('<B1-Motion>', self.on_drag_motion)
        self.lst_images.bind('<ButtonRelease-1>', self.on_drag_drop)
        self._drag_index = None

        img_scroll = ttk.Scrollbar(img_list_container, orient=VERTICAL, command=self.lst_images.yview)
        img_scroll.pack(side=RIGHT, fill=Y)
        self.lst_images.config(yscrollcommand=img_scroll.set)

        img_btn_frame = ttk.Frame(img_group)
        img_btn_frame.pack(fill=X, pady=5)
        ttk.Button(img_btn_frame, text="上移", command=lambda: self.move_item(-1), width=6).pack(side=LEFT, padx=2)
        ttk.Button(img_btn_frame, text="下移", command=lambda: self.move_item(1), width=6).pack(side=LEFT, padx=2)
        ttk.Button(img_btn_frame, text="整理", command=self.load_images, width=6).pack(side=LEFT, padx=2)

        # --- Action Controls ---
        action_frame = ttk.Frame(main_frame, padding=5)
        action_frame.pack(fill=X)

        self.btn_start = ttk.Button(action_frame, text="開始 (Start)", command=self.start_script, bootstyle=SUCCESS, width=15)
        self.btn_start.pack(side=LEFT, padx=5)

        self.btn_stop = ttk.Button(action_frame, text="停止 (Stop)", command=self.stop_script, bootstyle=DANGER, width=15, state=DISABLED)
        self.btn_stop.pack(side=LEFT, padx=5)

        ttk.Button(action_frame, text="儲存設定", command=self.save_config, width=15).pack(side=LEFT, padx=5)

        self.chk_debug = ttk.Checkbutton(action_frame, text="Debug 模式")
        self.chk_debug.pack(side=LEFT, padx=10)
        self.chk_debug.state(['!selected'])

        # --- Status Bar ---
        self.lbl_status = ttk.Label(main_frame, text="狀態：準備就緒", relief=SUNKEN, anchor=W, padding=5)
        self.lbl_status.pack(fill=X, pady=(5, 0))

        self.emu_vars = [] # List of (index, name, var, checkbox)

    def on_drag_start(self, event):
        self._drag_index = self.lst_images.nearest(event.y)

    def on_drag_motion(self, event):
        i = self.lst_images.nearest(event.y)
        if i != self._drag_index:
            x = self.lst_images.get(self._drag_index)
            self.lst_images.delete(self._drag_index)
            self.lst_images.insert(i, x)
            self._drag_index = i

    def on_drag_drop(self, event):
        self._drag_index = None

    def on_ld_path_changed(self):
        ld_path = self.txt_ld_path.get()
        if os.path.exists(ld_path):
            _, auto_scr = self.auto_detect_paths(ld_path)
            current_scr = self.txt_screenshot_dir.get()
            if not current_scr or not os.path.exists(current_scr):
                self.txt_screenshot_dir.delete(0, END)
                self.txt_screenshot_dir.insert(0, auto_scr)

    def browse_ld_path(self):
        filename = filedialog.askopenfilename(filetypes=[("Executable", "*.exe")])
        if filename:
            self.txt_ld_path.delete(0, END)
            self.txt_ld_path.insert(0, filename)
            self.on_ld_path_changed()

    def browse_screenshot_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.txt_screenshot_dir.delete(0, END)
            self.txt_screenshot_dir.insert(0, directory)

    def move_item(self, direction):
        selection = self.lst_images.curselection()
        if not selection:
            return
        
        index = selection[0]
        new_index = index + direction
        
        if 0 <= new_index < self.lst_images.size():
            item = self.lst_images.get(index)
            self.lst_images.delete(index)
            self.lst_images.insert(new_index, item)
            self.lst_images.selection_set(new_index)

    def load_config(self):
        config = configparser.ConfigParser()
        
        ld_path = ""
        screenshot_dir = ""
        wait_seconds = "1.0"
        debug_mode = "False"
        selected_emus = ""
        
        if os.path.exists(self.config_path):
            # Try UTF-8 first, then CP950 (Big5)
            try:
                config.read(self.config_path, encoding='utf-8')
            except UnicodeDecodeError:
                try:
                    config.read(self.config_path, encoding='cp950')
                except Exception as e:
                    print(f"Error reading config: {e}")
            
            if 'Settings' in config:
                ld_path = config['Settings'].get('ld_path', '')
                screenshot_dir = config['Settings'].get('screenshot_dir', '')
                wait_seconds = config['Settings'].get('wait_seconds', '1.0')
                debug_mode = config['Settings'].get('debug_mode', 'False')
                selected_emus = config['Settings'].get('selected_emulators', '')

        # Auto detect missing paths
        auto_ld, auto_scr = self.auto_detect_paths(ld_path)
        
        if not ld_path or not os.path.exists(ld_path):
            ld_path = auto_ld
            # If LD path changed, re-detect screenshot dir
            _, auto_scr = self.auto_detect_paths(ld_path)
            
        if not screenshot_dir or not os.path.exists(screenshot_dir):
            screenshot_dir = auto_scr

        self.txt_ld_path.delete(0, END)
        self.txt_ld_path.insert(0, ld_path)
        self.txt_screenshot_dir.delete(0, END)
        self.txt_screenshot_dir.insert(0, screenshot_dir)
        self.num_wait_seconds.set(wait_seconds)
        if debug_mode.lower() == "true":
            self.chk_debug.state(['selected'])
        
        self.saved_selected_emus = selected_emus.split('|') if selected_emus else []

    def auto_detect_paths(self, current_ld=""):
        common_paths = [
            r"C:\LDPlayer\LDPlayer9\ld.exe",
            r"D:\LDPlayer\LDPlayer9\ld.exe",
            r"E:\LDPlayer\LDPlayer9\ld.exe",
            r"F:\LDPlayer\LDPlayer9\ld.exe",
            r"C:\Program Files\LDPlayer\LDPlayer9\ld.exe",
            r"C:\XuanZhi\LDPlayer9\ld.exe"
        ]
        
        ld_path = current_ld
        if not ld_path or not os.path.exists(ld_path):
            ld_path = ""
            for p in common_paths:
                if os.path.exists(p):
                    ld_path = p
                    break
        
        if not ld_path:
            path_env = os.environ.get("PATH", "")
            for p in path_env.split(os.pathsep):
                potential = os.path.join(p, "ld.exe")
                if os.path.exists(potential):
                    ld_path = potential
                    break
        
        scr_dir = ""
        if ld_path:
            try:
                ld_dir = os.path.dirname(ld_path)
                # Check multiple possible config locations
                cfg_candidates = [
                    os.path.join(ld_dir, "leidian.config"),
                    os.path.join(ld_dir, "vms", "config", "leidian.config"),
                    os.path.join(ld_dir, "vms", "config", "leidians.config")
                ]
                for cfg_path in cfg_candidates:
                    if os.path.exists(cfg_path):
                        with open(cfg_path, 'r', encoding='utf-8', errors='ignore') as f:
                            for line in f:
                                if '"picturePath"' in line:
                                    parts = line.split(':')
                                    if len(parts) > 1:
                                        val = parts[1].strip().strip('", ').replace('\\\\', '\\')
                                        if os.path.exists(val):
                                            scr_dir = val
                                            break
                    if scr_dir: break
            except: pass
            
        # Fallback 1: Documents/XuanZhi9/Pictures (Using win32com for reliability)
        if not scr_dir and shell:
            try:
                docs = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0)
                xz_pics = os.path.join(docs, "XuanZhi9", "Pictures")
                if os.path.exists(xz_pics):
                    scr_dir = xz_pics
            except: pass

        # Fallback 2: Pictures/LDPlayer
        if not scr_dir and shell:
            try:
                pics = shell.SHGetFolderPath(0, shellcon.CSIDL_MYPICTURES, None, 0)
                ld_pics = os.path.join(pics, "LDPlayer")
                if os.path.exists(ld_pics):
                    scr_dir = ld_pics
            except: pass

        # Final Fallback
        if not scr_dir:
            scr_dir = r"D:\Screenshots"
            
        return ld_path, scr_dir

    def save_config(self):
        config = configparser.ConfigParser()
        config['Settings'] = {
            'ld_path': self.txt_ld_path.get(),
            'screenshot_dir': self.txt_screenshot_dir.get(),
            'wait_seconds': self.num_wait_seconds.get(),
            'debug_mode': str('selected' in self.chk_debug.state()),
            'selected_emulators': '|'.join([str(e[0]) for e in self.emu_vars if e[2].get()]),
            'image_order': '|'.join(self.lst_images.get(0, END))
        }
        
        with open(self.config_path, 'w', encoding='utf-8') as f:
            config.write(f)
        messagebox.showinfo("提示", "設定已儲存")

    def load_images(self):
        # Fallback to TraceLD_Project/trace if local trace is missing or empty
        if not os.path.exists(self.trace_dir) or not any(f.lower().endswith(".png") for f in os.listdir(self.trace_dir) if os.path.isdir(self.trace_dir)):
            fallback = os.path.join(self.script_dir, "TraceLD_Project", "trace")
            if os.path.exists(fallback) and any(f.lower().endswith(".png") for f in os.listdir(fallback)):
                self.trace_dir = fallback
        
        if not os.path.exists(self.trace_dir):
            os.makedirs(self.trace_dir, exist_ok=True)
            messagebox.showwarning("提示", f"找不到 trace 資料夾，已自動建立：\n{self.trace_dir}\n請將比對圖片放入此資料夾。")
            return

        config = configparser.ConfigParser()
        try:
            config.read(self.config_path, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                config.read(self.config_path, encoding='cp950')
            except:
                pass
        
        saved_order = config.get('Settings', 'image_order', fallback='')
        saved_items = [i for i in saved_order.split('|') if i]
        
        current_files = set([f for f in os.listdir(self.trace_dir) if f.lower().endswith(".png")])
        
        if not current_files and not saved_items:
            self.update_status("trace 資料夾內無圖片")

        self.lst_images.delete(0, END)
        for item in saved_items:
            if item in current_files:
                self.lst_images.insert(END, item)
                current_files.remove(item)
        
        for file in sorted(list(current_files)):
            self.lst_images.insert(END, file)

    def refresh_emulators(self):
        ld_path = self.txt_ld_path.get()
        if not os.path.exists(ld_path):
            return

        ld_dir = os.path.dirname(ld_path)
        console_path = os.path.join(ld_dir, "ldconsole.exe")
        if not os.path.exists(console_path):
            console_path = ld_path

        try:
            # Try UTF-8 first
            try:
                result = subprocess.run([console_path, "list2"], capture_output=True, text=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW)
                stdout = result.stdout
            except UnicodeDecodeError:
                result = subprocess.run([console_path, "list2"], capture_output=True, text=True, encoding='cp950', creationflags=subprocess.CREATE_NO_WINDOW)
                stdout = result.stdout

            lines = stdout.strip().split('\n')
            
            # Clear existing
            for _, _, _, cb in self.emu_vars:
                cb.destroy()
            self.emu_vars = []

            for line in lines:
                if not line: continue
                parts = line.split(',')
                if len(parts) >= 5:
                    idx = int(parts[0])
                    name = parts[1]
                    is_running = parts[4] == "1"
                    
                    var = tk.BooleanVar(value=is_running)
                    # If we had saved selection, override
                    if hasattr(self, 'saved_selected_emus') and str(idx) in self.saved_selected_emus:
                        var.set(True)

                    cb = ttk.Checkbutton(self.emu_scrollable_frame, text=f"{idx}: {name}", variable=var)
                    cb.pack(fill=X, anchor=W, padx=5, pady=2)
                    self.emu_vars.append((idx, name, var, cb))
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取模擬器清單失敗: {e}")

    def select_all_emus(self):
        for _, _, var, _ in self.emu_vars:
            var.set(True)

    def deselect_all_emus(self):
        for _, _, var, _ in self.emu_vars:
            var.set(False)

    def update_status(self, msg):
        self.lbl_status.config(text=f"狀態：{msg}")
        if 'selected' in self.chk_debug.state():
            self.log(msg)

    def log(self, msg):
        now = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        with open("trace_log.txt", "a", encoding='utf-8') as f:
            f.write(f"[{now}] {msg}\n")

    def check_emulator_status(self, index):
        ld_path = self.txt_ld_path.get()
        if not os.path.exists(ld_path):
            return False

        ld_dir = os.path.dirname(ld_path)
        console_path = os.path.join(ld_dir, "ldconsole.exe")
        if not os.path.exists(console_path):
            console_path = ld_path

        try:
            # Try UTF-8 first
            try:
                result = subprocess.run([console_path, "list2"], capture_output=True, text=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW)
                stdout = result.stdout
            except UnicodeDecodeError:
                result = subprocess.run([console_path, "list2"], capture_output=True, text=True, encoding='cp950', creationflags=subprocess.CREATE_NO_WINDOW)
                stdout = result.stdout
                
            lines = stdout.strip().split('\n')
            for line in lines:
                parts = line.split(',')
                if len(parts) >= 5 and parts[0] == str(index):
                    return parts[4] == "1"
        except:
            pass
        return False

    def start_script(self):
        if self.is_running:
            return

        ld_path = self.txt_ld_path.get()
        if not os.path.exists(ld_path):
            messagebox.showerror("錯誤", "找不到 LDPlayer 執行檔 (ld.exe)！")
            return

        selected_indices = [e[0] for e in self.emu_vars if e[2].get()]
        if not selected_indices:
            messagebox.showwarning("提示", "請至少勾選一個模擬器！")
            return

        if 'selected' in self.chk_debug.state():
            if os.path.exists("trace_log.txt"):
                os.remove("trace_log.txt")

        self.update_status("載入圖片快取...")
        self.matcher.load_templates(self.trace_dir)
        
        self.ld_manager = LdPlayerManager(ld_path, self.txt_screenshot_dir.get())
        self.is_running = True
        self.btn_start.config(state=DISABLED)
        self.btn_stop.config(state=NORMAL)
        
        self.worker_thread = threading.Thread(target=self.run_automation, daemon=True)
        self.worker_thread.start()

    def stop_script(self):
        self.is_running = False
        self.btn_start.config(state=NORMAL)
        self.btn_stop.config(state=DISABLED)
        self.update_status("已停止")

    def run_automation(self):
        try:
            self.update_status("開始執行...")
            
            while self.is_running:
                images = list(self.lst_images.get(0, END))
                selected_indices = [e[0] for e in self.emu_vars if e[2].get()]
                wait_sec = float(self.num_wait_seconds.get())

                if not selected_indices:
                    self.root.after(0, lambda: self.update_status("未選擇任何模擬器"))
                    break

                for emu_index in selected_indices:
                    if not self.is_running:
                        break
                    
                    self.ld_manager.index = emu_index
                    self.root.after(0, lambda idx=emu_index: self.update_status(f"[{idx}] 截圖中..."))

                    cap_img = self.ld_manager.screencap()
                    if cap_img is not None:
                        matched = False
                        # Search region from C#: Rectangle(180, 120, 430, 60)
                        search_region = (180, 120, 430, 60)

                        for img_name in images:
                            self.root.after(0, lambda idx=emu_index, name=img_name: self.update_status(f"[{idx}] 比對: {name}"))
                            result = self.matcher.find_image(cap_img, img_name, threshold=0.9, search_region=search_region)
                            
                            if result and result["success"]:
                                loc = result["location"]
                                self.root.after(0, lambda idx=emu_index, name=img_name, x=loc[0]: self.update_status(f"[{idx}] 匹配: {name} -> 點擊 ({x}, 320)"))
                                self.ld_manager.click(loc[0], 320)
                                matched = True
                                break
                        
                        if not matched:
                            self.root.after(0, lambda idx=emu_index: self.update_status(f"[{idx}] 無匹配項"))
                    else:
                        self.root.after(0, lambda idx=emu_index: self.update_status(f"[{idx}] 截圖失敗，檢查狀態..."))
                        if not self.check_emulator_status(emu_index):
                            self.root.after(0, lambda idx=emu_index: self.update_status(f"[{idx}] 模擬器已關閉，取消勾選"))
                            # Find the var and uncheck it
                            for idx, _, var, _ in self.emu_vars:
                                if idx == emu_index:
                                    self.root.after(0, lambda v=var: v.set(False))
                                    break

                if self.is_running:
                    time.sleep(wait_sec)
        except Exception as e:
            self.root.after(0, lambda msg=str(e): self.update_status(f"錯誤: {msg}"))
            self.is_running = False
            self.root.after(0, lambda: self.btn_start.config(state=NORMAL))
            self.root.after(0, lambda: self.btn_stop.config(state=DISABLED))

if __name__ == "__main__":
    root = ttk.Window()
    app = TraceLDApp(root)
    root.mainloop()
