import time
import threading
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageDraw
from comtypes import CLSCTX_ALL, CoInitialize
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import os
import sys
import winreg
from infi.systray import SysTrayIcon

APP_NAME = "MicVolumeEnforcerByAsid"
STARTUP_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_PATH = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)

running = True
thread = None
systray = None

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS  # PyInstaller extracts files here at runtime
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_volume_control():
    devices = AudioUtilities.GetMicrophone()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return interface.QueryInterface(IAudioEndpointVolume)


def volume_loop():
    global running
    CoInitialize()
    volume = get_volume_control()
    max_volume_value = volume.GetVolumeRange()[1]
    while running:
        volume.SetMasterVolumeLevel(max_volume_value, None)
        time.sleep(0.1)


def toggle_loop():
    global running, thread
    running = not running
    if running:
        toggle_btn.config(text="Stop")
        thread = threading.Thread(target=volume_loop, daemon=True)
        thread.start()
    else:
        toggle_btn.config(text="Start")


def hide_gui():
    root.withdraw()


def show_gui():
    root.deiconify()
    root.after(0, root.lift)


def on_quit(systray):
    global running
    running = False
    root.quit()


def set_startup(enabled: bool):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY, 0, winreg.KEY_ALL_ACCESS) as key:
            if enabled:
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{APP_PATH}"')
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass
    except PermissionError:
        messagebox.showerror("Permission Denied", "Run as admin to change startup behavior.")


def is_startup_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY) as key:
            winreg.QueryValueEx(key, APP_NAME)
            return True
    except FileNotFoundError:
        return False


# ---------------- GUI ----------------
root = tk.Tk()
root.title("Mic Volume Enforcer By Asid")
root.geometry("400x180")
root.protocol("WM_DELETE_WINDOW", hide_gui)

toggle_btn = tk.Button(root, text="Stop", command=toggle_loop)
toggle_btn.pack(expand=False, fill='x', padx=20, pady=10)

var = tk.BooleanVar(value=is_startup_enabled())
startup_check = tk.Checkbutton(root, text="Run at Windows startup", variable=var, command=lambda: set_startup(var.get()))
startup_check.pack(pady=5)

quit_btn = tk.Button(root, text="Quit", fg="white", bg="red", command=lambda: on_quit(None))
quit_btn.pack(pady=10, fill='x', padx=20)

# ---------------- Tray Setup ----------------
icon_path = resource_path("ses100.ico")

# Create a basic .ico file if not exists
if not os.path.exists(icon_path):
    image = Image.new("RGB", (64, 64), "black")
    draw = ImageDraw.Draw(image)
    draw.rectangle((16, 16, 48, 48), fill="white")
    image.save(icon_path)

menu_options = (("Show Window", None, lambda systray: show_gui()),)

systray = SysTrayIcon(icon_path, APP_NAME, menu_options, on_quit=on_quit)

# Handle tray left-click to show window
def tray_click_handler(hwnd, msg, wparam, lparam):
    WM_LBUTTONDOWN = 0x0201
    if lparam == WM_LBUTTONDOWN:
        show_gui()

systray._SysTrayIcon__on_notify = tray_click_handler

# Start systray and background volume thread
threading.Thread(target=systray.start, daemon=True).start()
thread = threading.Thread(target=volume_loop, daemon=True)
thread.start()

# Start GUI hidden
hide_gui()
root.mainloop()
