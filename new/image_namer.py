import automation
import time
import pyautogui
import pywinauto

def name_files(naming_template, placeholder_values, xy_name, delay, channel_orders_list):
    windows = pywinauto.Desktop(backend="win32").windows()
    matching_windows = [win for win in windows if "BZ-X800 Wide Image Viewer" in win.window_text()]

    for i in range(len(matching_windows)):
        last_window = matching_windows[i]
        last_window.set_focus()
        
        # Click File Button and Export in Original Scale
        automation.export_image(last_window)
        
        # Naming Code
        channel = channel_orders_list[i]
        file_name = naming_template.format(C=channel, **placeholder_values[xy_name])
        print(f"Naming file: {file_name}")
        pyautogui.write(file_name)
        pyautogui.press('tab', presses=2)
        pyautogui.press('enter')
        
        # Close Image
        time.sleep(delay)
        pyautogui.hotkey('alt', 'f4')
        pyautogui.press('tab', presses=1)
        pyautogui.press('enter')
        print(channel + " image closed.")