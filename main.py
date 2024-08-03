import pywinauto
import pyautogui
import time
import os
import sys

# Constants
IMAGE_PATH = os.path.join(os.path.dirname(__file__), 'image.png')
WIDE_IMAGE_VIEWER_TITLE = "BZ-X800 Wide Image Viewer"
MAX_DELAY_TIME = 20

def main():
    run_configs = get_multiple_run_configs()
    failed = []

    for run_config in run_configs:
        run_name, stitchtype, overlay, naming_template, filepath, start_child, end_child = run_config
        placeholder_values = get_placeholder_values(naming_template, start_child, end_child)

        try:
            process_xy_sequences(failed, run_name, stitchtype, overlay, naming_template, start_child, end_child, placeholder_values, filepath)
        except Exception as e:
            print(f"Failed on running {run_name}. Error: {e}. Moving on to the next run.")

    if failed:
        print("All XY sequences have been processed except for: ", failed)
    else:   
        print("All XY sequences have been processed successfully!")

def get_multiple_run_configs():
    run_configs = []
    while True:
        run_name = input("Enter Run Name (or press Enter to finish): ")
        if not run_name:
            break
        
        stitchtype = input("Stitch Type? Full (F) or Load (L): ").upper()
        overlay = input("Overlay Image? (Y/N): ").upper()
        naming_template = input("Enter the naming template (use {key1}, {key2}, etc. for placeholders and {C} for channel): ")
        filepath = input("Enter the EXACT filepath to save the images. Press ENTER if you want to use the previous filepath loaded by Keyence: ")
        
        start_child, end_child = get_xy_sequence_range(run_name)
        
        run_configs.append((run_name, stitchtype, overlay, naming_template, filepath, start_child, end_child))
    
    return run_configs

def get_xy_sequence_range(run_name):
    run_tree_item = main_window.child_window(title=run_name, control_type="TreeItem")
    children = run_tree_item.children()
    print(f"Detected {len(children)} XY sequences for {run_name}.")
    start_child = int(input(f"Enter the starting XY number for {run_name}: "))
    end_child = int(input(f"Enter the ending XY number for {run_name}: "))
    return start_child, end_child

def get_placeholder_values(naming_template, start_child, end_child):
    placeholder_values = {}
    xy_names = [f"XY{i:02}" for i in range(start_child, end_child + 1)]

    # Find all unique placeholders in the naming template
    placeholders = set(i for i in range(1, 10) if f"{{key{i}}}" in naming_template)

    for xy_name in xy_names:
        placeholder_values[xy_name] = {}
        for placeholder in placeholders:
            value = input(f"Enter placeholder {{key{placeholder}}} for {xy_name}: ")
            placeholder_values[xy_name][f'key{placeholder}'] = value

    return placeholder_values

def process_xy_sequences(failed, run_name, stitchtype, overlay, naming_template, start_child, end_child, placeholder_values, filepath):
    main_window.set_focus()
    run_tree_item = main_window.child_window(title=run_name, control_type="TreeItem")
    run_tree_item.expand()
    run_tree_item.click_input()

    for i in range(start_child - 1, end_child):
        try:
            child = run_tree_item.children()[i]
            xy_name = child.window_text()
            print(f"Processing {xy_name}")
            child.click_input()
            stitch_button.click_input()

            select_stitch_type(stitchtype)
            check_for_image()
            
            image_stitch = app.window(title="BZ-X800 Analyzer").child_window(auto_id="ImageJointMainForm", class_name="WindowsForms10.Window.8.app.0.3553390_r8_ad1", title = "Image Stitch")
            
            close_button = image_stitch.child_window(auto_id="_buttonCancel", title = "Cancel", class_name="WindowsForms10.BUTTON.app.0.3553390_r8_ad1")
            
            image_stitch.set_focus()
            start_stitching(overlay)
            
            delay_time = wait_for_wide_image_viewer()
            process_delay_time = delay_time*(len(channel_orders_list)-0.7)
            print(f"Waiting for {process_delay_time:.2f} seconds")
            time.sleep(process_delay_time) # ! Arbitary value. Figure out a way to prevent this hardcoded result.
            disable_caps_lock()
            name_files(naming_template, placeholder_values, xy_name, delay_time, filepath)
            
            image_stitch.set_focus()
            close_button.click_input()
            print("Closing stitch image...")
            
        except Exception as e:
            print(f"Failed on running {xy_name}. Error: {e}")
            failed.append(xy_name)

def select_stitch_type(stitchtype):
    # Full Focus
    if stitchtype == "F":
        pyautogui.press('f')
        pyautogui.press('enter')
    # Load
    elif stitchtype == "L":
        pyautogui.press('l')

def check_for_image():
    assert os.path.exists(IMAGE_PATH)
    start_time = time.time()
    print(IMAGE_PATH)
    while True:
        try:
            location = pyautogui.locateOnScreen(IMAGE_PATH, grayscale=True, confidence=0.95)
            if location is not None:
                pyautogui.click(location)
                print("Found image!")
                return
        except Exception:
            print("Time elapsed:", round(time.time() - start_time, 0), "s")
            time.sleep(2)
            if time.time() - start_time > 10 * 60:
                print("Image not found after 10 minutes... terminating search.")
                sys.exit()

def start_stitching(overlay):
    pyautogui.press('tab', presses=6)
    pyautogui.press('right')
    if overlay == "Y":
        pyautogui.press('tab', presses=3)
    elif overlay == "N":
        pyautogui.press('tab', presses=2)
    pyautogui.press('enter')

def wait_for_wide_image_viewer():
    start_time = time.time()
    while True:
        windows = pywinauto.Desktop(backend="win32").windows()
        matching_windows = [win for win in windows if WIDE_IMAGE_VIEWER_TITLE in win.window_text()]
        if matching_windows:
            break
        time.sleep(0.1)
    delay_time = min(time.time() - start_time, MAX_DELAY_TIME)
    return delay_time

def disable_caps_lock():
    import ctypes
    hllDll = ctypes.WinDLL ("User32.dll")
    VK_CAPITAL = 0x14
    if hllDll.GetKeyState(VK_CAPITAL):
        pyautogui.press('capslock')
        print("Caps Lock disabled.")

def name_files(naming_template, placeholder_values, xy_name, delay, filepath):
    windows = pywinauto.Desktop(backend="win32").windows()
    matching_windows = [win for win in windows if WIDE_IMAGE_VIEWER_TITLE in win.window_text()]

    for i in range(len(matching_windows)):
        last_window = matching_windows[i]
        last_window.set_focus()

        click_file_button(last_window)
        export_in_original_scale()
        
        if i == 0 and filepath != "":
            pyautogui.press('tab', presses=6)
            pyautogui.press('enter')
            pyautogui.write(filepath)
            pyautogui.press('enter')
            pyautogui.press('tab', presses=6)
            print(f"Filepath set to: {filepath}")

        channel = channel_orders_list[i]
        try:
            # Create a dictionary with all possible placeholders
            format_dict = {f'key{k}': '' for k in range(1, 10)}  # Initialize all placeholders
            format_dict.update(placeholder_values[xy_name])  # Update with actual values
            format_dict['C'] = channel

            file_name = naming_template.format(**format_dict)
            print(f"Naming file: {file_name}")
            time.sleep(1)  # ! Arbitary value. Figure out a way to prevent this hardcoded result.
            pyautogui.write(file_name)
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
        except KeyError as e:
            print(f"Error: Missing placeholder {e} in naming template for {xy_name}")
        except Exception as e:
            print(f"Unexpected error occurred while naming file for {xy_name}: {e}")

        close_image(delay, channel)

def click_file_button(window):
    app = pywinauto.Application(backend="uia").connect(handle=window.handle)
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
    time.sleep(delay/3) # ! Arbitary value. Figure out a way to prevent this hardcoded result.
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.1)
    pyautogui.press('tab', presses=1)
    time.sleep(0.1)
    pyautogui.press('enter')
    print(f"{channel} image closed.")

def close_stitch_image():
    pyautogui.press('tab', presses=2)
    pyautogui.press('enter')
    time.sleep(1) # ! Arbitary value. Figure out a way to prevent this hardcoded result.

def define_channel_orders():
    channel_count = int(input("How many channels were imaged? "))
    channel_orders_list = []
    print(f"Enter the channel name type of the {channel_count} channels from opened first to opened last. The last one should be Overlay.")
    for i in range(channel_count):
        order = input(f"Channel {i+1} of {channel_count}: ")
        channel_orders_list.append(order)
    channel_orders_list.reverse()
    return channel_orders_list

def display_splash_art():
    splash_art = r"""
    =======================================
            Keyence Auto Namer
    =======================================
            Created by: Maxx Yung
            Version: 1.0.2
            Last Updated: 2024-08-01
    =======================================
    """
    print(splash_art)
    time.sleep(0.5)

# Initial setup
while True:
    try:
        app = pywinauto.Application(backend="uia").connect(title="BZ-X800 Analyzer")
        main_window = app.window(title="BZ-X800 Analyzer")
        stitch_button = app.window(title="BZ-X800 Analyzer").child_window(title="Stitch")
        break
    except Exception:
        print("BZ-X800 Analyzer not found. Please open the application and try again. Press Enter to retry.")
        input()

display_splash_art()
channel_orders_list = define_channel_orders()
print("Channel Orders:", channel_orders_list)

# Run Program
if __name__ == "__main__":
    main()