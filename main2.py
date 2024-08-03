import pywinauto
import pyautogui
import time
import os
import sys
import logging
from pywinauto.application import Application
from pywinauto import Desktop, timings
from pywinauto.keyboard import send_keys

# Constants
LOG_FILE = 'keyence_auto_namer.log'
IMAGE_PATH = os.path.join(os.path.dirname(__file__), 'image.png')
WIDE_IMAGE_VIEWER_TITLE = 'BZ-X800 Wide Image Viewer'
MAX_DELAY_TIME = 20
ANALYZER_TITLE = 'BZ-X800 Analyzer'

# Setup logging
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    logging.info("Starting Keyence Auto Namer")
    run_configs = get_multiple_run_configs()
    failed = []

    for run_config in run_configs:
        run_name, stitchtype, overlay, naming_template, filepath, start_child, end_child = run_config
        placeholder_values = get_placeholder_values(naming_template, start_child, end_child)

        try:
            process_xy_sequences(failed, run_name, stitchtype, overlay, naming_template, start_child, end_child, placeholder_values, filepath)
        except Exception as e:
            logging.error(f"Failed on running {run_name}. Error: {str(e)}")
            print(f"Failed on running {run_name}. Error: {str(e)}. Moving on to the next run.")

    if failed:
        logging.warning(f"All XY sequences have been processed except for: {failed}")
        print("All XY sequences have been processed except for: ", failed)
    else:   
        logging.info("All XY sequences have been processed successfully!")
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
    main_window = get_main_window()
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
    main_window = get_main_window()
    main_window.set_focus()
    run_tree_item = main_window.child_window(title=run_name, control_type="TreeItem")
    run_tree_item.expand()
    run_tree_item.click_input()

    for i in range(start_child - 1, end_child):
        try:
            child = run_tree_item.children()[i]
            xy_name = child.window_text()
            logging.info(f"Processing {xy_name}")
            print(f"Processing {xy_name}")
            child.click_input()
            stitch_button = main_window.child_window(title="Stitch")
            stitch_button.click_input()

            select_stitch_type(stitchtype)
            check_for_image()
            
            image_stitch = main_window.child_window(auto_id="ImageJointMainForm", title="Image Stitch")
            
            close_button = image_stitch.child_window(auto_id="_buttonCancel", title="Cancel")
            
            image_stitch.set_focus()
            start_stitching(overlay)
            
            delay_time = wait_for_wide_image_viewer()
            process_delay_time = delay_time * (len(channel_orders_list) - 0.7)
            logging.info(f"Waiting for {process_delay_time:.2f} seconds")
            print(f"Waiting for {process_delay_time:.2f} seconds")
            time.sleep(process_delay_time)
            disable_caps_lock()
            name_files(naming_template, placeholder_values, xy_name, delay_time, filepath)
            
            image_stitch.set_focus()
            close_button.click_input()
            logging.info("Closing stitch image...")
            print("Closing stitch image...")
            
        except Exception as e:
            logging.error(f"Failed on running {xy_name}. Error: {str(e)}")
            print(f"Failed on running {xy_name}. Error: {str(e)}")
            failed.append(xy_name)

def select_stitch_type(stitchtype):
    if stitchtype == "F":
        send_keys('f{ENTER}')
    elif stitchtype == "L":
        send_keys('l')

def check_for_image():
    image_path = IMAGE_PATH
    assert os.path.exists(image_path)
    start_time = time.time()
    logging.info(f"Searching for image: {image_path}")
    print(f"Searching for image: {image_path}")
    while True:
        try:
            location = pyautogui.locateOnScreen(image_path, grayscale=True, confidence=0.95)
            if location is not None:
                pyautogui.click(location)
                logging.info("Image found and clicked!")
                print("Image found and clicked!")
                return
        except Exception:
            elapsed_time = round(time.time() - start_time, 0)
            logging.info(f"Time elapsed: {elapsed_time} s")
            print(f"Time elapsed: {elapsed_time} s")
            time.sleep(2)
            if elapsed_time > 10 * 60:
                logging.error("Image not found after 10 minutes... terminating search.")
                print("Image not found after 10 minutes... terminating search.")
                sys.exit()

def start_stitching(overlay):
    send_keys('{TAB 6}{RIGHT}')
    if overlay == "Y":
        send_keys('{TAB 3}')
    elif overlay == "N":
        send_keys('{TAB 2}')
    send_keys('{ENTER}')

def wait_for_wide_image_viewer():
    start_time = time.time()
    while True:
        windows = Desktop(backend="uia").windows()
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
        send_keys('{CAPSLOCK}')
        logging.info("Caps Lock disabled.")
        print("Caps Lock disabled.")

def name_files(naming_template, placeholder_values, xy_name, delay, filepath):
    windows = Desktop(backend="uia").windows()
    matching_windows = [win for win in windows if WIDE_IMAGE_VIEWER_TITLE in win.window_text()]

    for i in range(len(matching_windows)):
        last_window = matching_windows[i]
        last_window.set_focus()

        click_file_button(last_window)
        export_in_original_scale()
        
        if i == 0 and filepath != "":
            send_keys('{TAB 6}{ENTER}')
            send_keys(filepath)
            send_keys('{ENTER}{TAB 6}')
            logging.info(f"Filepath set to: {filepath}")
            print(f"Filepath set to: {filepath}")

        channel = channel_orders_list[i]
        try:
            # Create a dictionary with all possible placeholders
            format_dict = {f'key{k}': '' for k in range(1, 10)}  # Initialize all placeholders
            format_dict.update(placeholder_values[xy_name])  # Update with actual values
            format_dict['C'] = channel

            file_name = naming_template.format(**format_dict)
            logging.info(f"Naming file: {file_name}")
            print(f"Naming file: {file_name}")
            time.sleep(1)
            send_keys(file_name)
            send_keys('{TAB 2}{ENTER}')
        except KeyError as e:
            logging.error(f"Error: Missing placeholder {e} in naming template for {xy_name}")
            print(f"Error: Missing placeholder {e} in naming template for {xy_name}")
        except Exception as e:
            logging.error(f"Unexpected error occurred while naming file for {xy_name}: {str(e)}")
            print(f"Unexpected error occurred while naming file for {xy_name}: {str(e)}")

        close_image(delay, channel)

def click_file_button(window):
    app = Application(backend="uia").connect(handle=window.handle)
    main_window = app.window(auto_id="MainForm", control_type="Window")
    toolbar = main_window.child_window(title="toolStrip1", control_type="ToolBar")
    file_button = toolbar.child_window(title="File", control_type="Button")
    file_button.click_input()

def export_in_original_scale():
    send_keys('{TAB 4}{ENTER}{TAB}{ENTER}')

def close_image(delay, channel):
    time.sleep(delay/3)
    send_keys('%{F4}')
    time.sleep(0.1)
    send_keys('{TAB}{ENTER}')
    logging.info(f"{channel} image closed.")
    print(f"{channel} image closed.")

def close_stitch_image():
    send_keys('{TAB 2}{ENTER}')
    time.sleep(1)

def define_channel_orders():
    channel_orders_list = []
    print("Enter the channel names in order, from first opened to last.")
    print("Press Enter on an empty line when you're finished.")
    print("Note: The last channel should typically be 'Overlay'.")
    
    while True:
        channel = input(f"Channel {len(channel_orders_list) + 1}: ").strip()
        if not channel:  # If the input is empty (user just pressed Enter)
            if not channel_orders_list:
                print("You must enter at least one channel.")
                continue
            break  # Exit the loop if we have at least one channel and the user enters a blank line
        channel_orders_list.append(channel)

    channel_orders_list.reverse()  
    return channel_orders_list

def display_splash_art():
    splash_art = r"""
    =======================================
            Keyence Auto Namer
    =======================================
            Created by: Maxx Yung
            Version: 1.1.0
            Last Updated: 2024-08-03
    =======================================
    """
    print(splash_art)
    time.sleep(0.5)

def get_main_window():
    try:
        app = Application(backend="uia").connect(title=ANALYZER_TITLE)
        main_window = app.window(title=ANALYZER_TITLE)
        return main_window
    except Exception as e:
        logging.error(f"BZ-X800 Analyzer not found. Error: {str(e)}")
        print("BZ-X800 Analyzer not found. Please open the application and try again. Press Enter to retry.")
        input()
        return get_main_window()

def run_tests():
    logging.info("Running tests...")
    print("Running tests...")
    
    # Test main window connection
    try:
        main_window = get_main_window()
        assert main_window.exists(), "Main window connection failed"
    except Exception as e:
        logging.error(f"Main window connection test failed: {str(e)}")
        print(f"Main window connection test failed: {str(e)}")
        return False

    # TODO Add more tests as needed

    logging.info("All tests passed successfully.")
    print("All tests passed successfully.")
    return True

# Initial setup
display_splash_art()
channel_orders_list = define_channel_orders()
logging.info(f"Channel Orders: {channel_orders_list}")
print(f"Channel Orders: {channel_orders_list}")

# Run Program
if __name__ == "__main__":
    try:
        if run_tests():
            main()
        else:
            logging.error("Tests failed. Please check the logs and fix any issues before running the main program.")
            print("Tests failed. Please check the logs and fix any issues before running the main program.")
    except Exception as e:
        logging.critical(f"An unexpected error occurred: {str(e)}")
        print(f"An unexpected error occurred: {str(e)}")
        print("Please check the log file for more details.")
    finally:
        logging.info("Program execution completed.")
        print("Program execution completed.")