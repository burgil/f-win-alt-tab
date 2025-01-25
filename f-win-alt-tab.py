import keyboard
import time
import ctypes
import sys
import win32com.client

DEBUG = True
def debug_print(*args):
    if DEBUG:
        print(*args)

objShell = win32com.client.Dispatch("Shell.Application")
alt_pressed = False

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
        else:
            print("IMPOSSIBLE Blocked ALT")
        return False
    if event.event_type == keyboard.KEY_DOWN:
        alt_pressed = True
        debug_print("ALT pressed")
    elif event.event_type == keyboard.KEY_UP:
        alt_pressed = False
        time.sleep(0.1)
        if is_win_tab_open():
            debug_print("ALT released while SWITCHER Open")
            keyboard.press_and_release('enter')
            debug_print("Closing Switcher B...")
        else:
            debug_print("ALT released while SWITCHER Closed")
    return True

def on_tab(event):
    global alt_pressed
    current_time = time.time()
    if event.event_type == keyboard.KEY_DOWN and alt_pressed:
        debug_print("ALT+TAB detected!")
        if not is_win_tab_open() and not is_alt_tab_open():
            objShell.WindowSwitcher()
            debug_print("Opened the Window Switcher!")
        else:
            keyboard.press_and_release('right')
            debug_print("Moving selection right")
    if is_win_tab_open():
        if event.event_type == keyboard.KEY_DOWN:
            debug_print("Blocked TAB KEY_DOWN press")
        elif event.event_type == keyboard.KEY_UP:
            debug_print("Blocked TAB KEY_UP press")
        else:
            print("IMPOSSIBLE Blocked TAB")
        return False
    return True

# Register hooks
keyboard.hook_key('alt', on_alt, suppress=True)
keyboard.hook_key('tab', on_tab, suppress=True)

debug_print("F-WIN-ALT-TAB is now running...")
debug_print("Press ALT+TAB to switch between windows on current monitor")
debug_print("Press CTRL+C to exit")

# Just keep the process running
while True:
    try:
        time.sleep(1)
    except:
        sys.exit()
        pass
