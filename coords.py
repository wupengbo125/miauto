"""
coords.py - 鼠标坐标抓取小工具
按 Ctrl+Z 后，把当前鼠标坐标复制到剪贴板，格式为 click "test",x:y
按 Ctrl+C 或 Esc 退出程序
"""

import pyautogui
import pyperclip
import keyboard
import time
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--only", action="store_true", help="只复制坐标 x:y")
args = parser.parse_args()

print("🖱️  坐标抓取工具已启动！")
print("   按 Ctrl+Z  → 抓取当前鼠标坐标并复制到剪贴板")
print("   按 Esc     → 退出程序")
print("-" * 40)

running = True

def on_ctrl_z():
    x, y = pyautogui.position()
    if args.only:
        text = f"{x}:{y}"
    else:
        text = f'click "test",{x}:{y}'
    pyperclip.copy(text)
    print(f"  📋 已复制: {text}")

def on_esc():
    global running
    print("\n👋 退出坐标抓取工具。")
    running = False

keyboard.add_hotkey("ctrl+z", on_ctrl_z)
keyboard.add_hotkey("esc", on_esc)

while running:
    time.sleep(0.1)
