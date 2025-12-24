import os
import configparser
from pathlib import Path

def auto_detect_paths():
    common_paths = [
        r"C:\LDPlayer\LDPlayer9\ld.exe",
        r"D:\LDPlayer\LDPlayer9\ld.exe",
        r"E:\LDPlayer\LDPlayer9\ld.exe",
        r"F:\LDPlayer\LDPlayer9\ld.exe",
        r"C:\Program Files\LDPlayer\LDPlayer9\ld.exe",
        r"C:\XuanZhi\LDPlayer9\ld.exe"
    ]
    
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
            cfg_path = os.path.join(ld_dir, "leidian.config")
            print(f"Checking config: {cfg_path}")
            if os.path.exists(cfg_path):
                with open(cfg_path, 'r', encoding='utf-8', errors='ignore') as f:
                    for line in f:
                        if '"picturePath"' in line:
                            parts = line.split(':')
                            if len(parts) > 1:
                                val = parts[1].strip().strip('", ').replace('\\\\', '\\')
                                print(f"Found picturePath in config: {val}")
                                if os.path.exists(val):
                                    scr_dir = val
                                    break
        except Exception as e:
            print(f"Error reading config: {e}")
            
    # Fallback 1: Documents/XuanZhi9/Pictures
    if not scr_dir:
        try:
            docs = Path.home() / "Documents"
            xz_pics = docs / "XuanZhi9" / "Pictures"
            print(f"Checking fallback 1: {xz_pics}")
            if xz_pics.exists():
                scr_dir = str(xz_pics)
        except: pass

    # Fallback 2: Pictures/LDPlayer
    if not scr_dir:
        try:
            pics = Path.home() / "Pictures"
            ld_pics = pics / "LDPlayer"
            print(f"Checking fallback 2: {ld_pics}")
            if ld_pics.exists():
                scr_dir = str(ld_pics)
        except: pass

    # Final Fallback
    if not scr_dir:
        scr_dir = r"D:\Screenshots"
        print(f"Final fallback: {scr_dir}")
        
    return ld_path, scr_dir

ld, scr = auto_detect_paths()
print(f"\nRESULT:")
print(f"LD Path: {ld}")
print(f"Screenshot Dir: {scr}")
