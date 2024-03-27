import os
import time
import pyautogui

def check_for_image(image_path):
    # Checks if an image exists on the screen by locating it using the given image path.
    start_time = time.time()
    while True:
        try:
            location = pyautogui.locateOnScreen(image_path, grayscale=True, confidence=0.95)
            if location is not None:
                print("Debug")
                return True
        except Exception:
            print("Image not found, trying again...",)
            time.sleep(0.5)
        if time.time() - start_time > 5 * 60:  # 5 minutes
            print("Image not found after 5 minutes... terminating search.")
            return False

file_name = os.path.join(os.path.dirname(__file__), 'image1.png')
assert os.path.exists(file_name)
check_for_image(file_name)