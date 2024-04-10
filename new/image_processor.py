import os
import time
import pyautogui
import pywinauto

def check_for_image(overlay):
    if overlay == "Y":
        file_name = os.path.join(os.path.dirname(__file__), 'overlay.png')
    elif overlay == "N":
        file_name = os.path.join(os.path.dirname(__file__), 'nooverlay.png')
    else :
        raise ValueError("Invalid overlay value. Must be 'Y' or 'N'.")
    assert os.path.exists(file_name)
    start_time = time.time()
    while True:
        try:
            location = pyautogui.locateOnScreen(file_name, grayscale=True, confidence=0.99)
            if location is not None:
                print("Found image!")
                pyautogui.click(location)
                return False
        except Exception:
            time.sleep(2)
        if time.time() - start_time > 5 * 60:  # 5 minutes
            print("Image not found after 5 minutes... terminating search.")
            return False

def wait_for_viewer():
    start_time = time.time()
    while True:
        windows = pywinauto.Desktop(backend="win32").windows()
        matching_windows = [win for win in windows if "BZ-X800 Wide Image Viewer" in win.window_text()]
        if matching_windows:
            break
        time.sleep(0.1)
    end_time = time.time()
    return end_time - start_time