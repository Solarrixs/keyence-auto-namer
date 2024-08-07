import pyautogui
import time
import logging
from pywinauto.application import Application
from constants import WIDE_IMAGE_VIEWER_TITLE, MAX_DELAY_TIME

def select_stitch_type(stitchtype):
    if stitchtype == "F":
        pyautogui.press('f')
        pyautogui.press('enter')
    elif stitchtype == "L":
        pyautogui.press('l')

def start_stitching(overlay):
    pyautogui.press('tab', presses=6)
    pyautogui.press('right')
    if overlay == "Y":
        pyautogui.press('tab', presses=3)
    elif overlay == "N":
        pyautogui.press('tab', presses=2)
    pyautogui.press('enter')

def wait_for_wide_image_viewer(desktop):
    start_time = time.time()
    while True:
        windows = desktop.windows()
        matching_windows = [win for win in windows if WIDE_IMAGE_VIEWER_TITLE in win.window_text()]
        if matching_windows:
            break
        time.sleep(0.1)
    delay_time = min(time.time() - start_time, float(MAX_DELAY_TIME))
    return delay_time

def disable_caps_lock():
    import ctypes
    hllDll = ctypes.WinDLL ("User32.dll")
    VK_CAPITAL = 0x14
    if hllDll.GetKeyState(VK_CAPITAL):
        pyautogui.press('capslock')
        logging.info("Caps Lock disabled.")

def click_file_button(window):
    app = Application(backend="uia").connect(handle=window.handle)
    main_window = app.window(auto_id="MainForm", control_type="Window")
    toolbar = main_window.child_window(title="toolStrip1", control_type="ToolBar")
    file_button = toolbar.child_window(title="File", control_type="Button")
    file_button.click_input()

def export_in_original_scale():
    pyautogui.press('tab', presses=4)
    pyautogui.press('enter')
    pyautogui.press('tab', presses=1)
    pyautogui.press('enter')

def close_image(delay, channel):
    time.sleep(delay/3) # ! Arbitrary value. Figure out a way to prevent this hardcoded result.
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.1)
    pyautogui.press('tab', presses=1)
    time.sleep(0.1)
    pyautogui.press('enter')
    print(f"{channel} image closed.")