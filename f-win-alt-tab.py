import os
import sys
import subprocess
import ctypes

title = 'f-win-alt-tab by Burgil - 2025'

os.system('title ' + title)
os.system('color a')

def is_title_open():
    windows = []
    def enum_windows_callback(hwnd, _):
        text = ctypes.create_unicode_buffer(512)
        ctypes.windll.user32.GetWindowTextW(hwnd, text, 512)
        if text.value == title:
            windows.append(hwnd)
        return True
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
    ctypes.windll.user32.EnumWindows(EnumWindowsProc(enum_windows_callback), 0)
    if len(windows) > 1:
        print("App already open!")
        sys.exit(1)

is_title_open()

# Function to run pip install with subprocess to avoid import issues
def install_package(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Check and fix pywin32 installation if needed
try:
    import win32com
except ImportError:
    install_package('pywin32')
    # Run post-install script to register DLLs
    import os.path
    import sys
    pythonpath = sys.executable
    binpath = os.path.dirname(pythonpath)
    os.system(f'"{binpath}/Scripts/pywin32_postinstall.py" -install')
    import win32com

try:
    import keyboard
except ImportError:
    os.system('pip install keyboard')
    import keyboard

try:
    import pystray
except ImportError:
    os.system('pip install pystray')
    import pystray

import time
import win32com.client
from ctypes import wintypes
import pystray
from PIL import Image
import win32gui
import win32con
import win32api
import win32console

os.system('title ' + title)
os.system('color a')

DEBUG = False
def debug_print(*args):
    global DEBUG
    if DEBUG:
        print(*args)

objShell = win32com.client.Dispatch("Shell.Application")
alt_pressed = False

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002
VK_RIGHT = 0x27

def send_right_key():
    ctypes.windll.user32.keybd_event(VK_RIGHT, 0, KEYEVENTF_EXTENDEDKEY, 0)
    ctypes.windll.user32.keybd_event(VK_RIGHT, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)

def is_win_tab_open():
    """Check if the Windows Task View window is currently open by checking foreground window"""
    foreground_window = ctypes.windll.user32.GetForegroundWindow() # Get foreground window handle
    text = ctypes.create_unicode_buffer(512) # Get window text
    ctypes.windll.user32.GetWindowTextW(foreground_window, text, 512)
    return text.value == "Task View" # Check if the foreground window is Task View

def is_alt_tab_open():
    """Check if the Windows Task View window is currently open by checking foreground window"""
    foreground_window = ctypes.windll.user32.GetForegroundWindow() # Get foreground window handle
    text = ctypes.create_unicode_buffer(512) # Get window text
    ctypes.windll.user32.GetWindowTextW(foreground_window, text, 512)
    return text.value == "Task Switching" # Check if the foreground window is Task Switching

def on_alt(event):
    global alt_pressed
    if is_win_tab_open():
        if event.event_type == keyboard.KEY_DOWN:
            debug_print("Blocked ALT KEY_DOWN press")
        elif event.event_type == keyboard.KEY_UP:
            if is_win_tab_open():
                keyboard.press_and_release('enter')
                debug_print("Closing Switcher A...")
                alt_pressed = False
                return True
            else:
                debug_print("Blocked ALT KEY_UP press")
        return False
    if event.event_type == keyboard.KEY_DOWN:
        alt_pressed = True
        debug_print("ALT pressed")
    elif event.event_type == keyboard.KEY_UP:
        alt_pressed = False
        if is_win_tab_open():
            debug_print("ALT released while SWITCHER Open")
            keyboard.press_and_release('enter')
            debug_print("Closing Switcher B...")
        else:
            debug_print("ALT released while SWITCHER Closed")
    return True

def on_tab(event):
    global alt_pressed
    if event.event_type == keyboard.KEY_DOWN and alt_pressed:
        debug_print("ALT+TAB detected!")
        keyboard.block_key('tab')
        if not is_win_tab_open() and not is_alt_tab_open():
            objShell.WindowSwitcher()
            debug_print("Opened the Window Switcher!")
        else:
            send_right_key()
            debug_print("Moving selection right")
        keyboard.unblock_key('tab')
        return False
    if is_win_tab_open():
        return False
    return True

current_window = win32console.GetConsoleWindow()

def toggle_console():
    global DEBUG
    if "--hide" in sys.argv:
        if win32gui.IsWindowVisible(current_window):
            DEBUG = False
            win32gui.ShowWindow(current_window, win32con.SW_HIDE)
        else:
            DEBUG = True
            win32gui.ShowWindow(current_window, win32con.SW_SHOW)
    else:
        print("DEBUG True2")
        DEBUG = True
        print("Disabled during development")

def create_tray_icon():
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.png')
    image = Image.open(icon_path)
    
    def on_exit(icon, item):
        icon.stop()
        sys.exit()

    menu = (
        pystray.MenuItem("Toggle Console", toggle_console),
        pystray.MenuItem("Exit", on_exit)
    )
    
    icon = pystray.Icon("f-win-alt-tab", image, "F-WIN-ALT-TAB", menu)
    return icon

if "--hide" in sys.argv:
    win32gui.ShowWindow(current_window, win32con.SW_HIDE)

# Fix CTRL+C

def handler(type):
    icon.stop()
    sys.exit(0)
    return True

win32api.SetConsoleCtrlHandler(handler, True)

keyboard.hook_key('alt', on_alt, suppress=True)
keyboard.hook_key('tab', on_tab, suppress=True)

print("F-WIN-ALT-TAB is now running...")
print("Press ALT+TAB to switch between windows on current monitor")
print("Press CTRL+C to exit")

icon = create_tray_icon()
try:
    icon.run()
except KeyboardInterrupt:
    icon.stop()
    sys.exit(0)
