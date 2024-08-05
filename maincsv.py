import pyautogui
import time
import os
import sys
import logging
import csv
from pywinauto.application import Application
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

# Constants
LOG_FILE = 'keyence_auto_namer.log'
IMAGE_PATH = os.path.join(os.path.dirname(__file__), 'image.png')
WIDE_IMAGE_VIEWER_TITLE = 'BZ-X800 Wide Image Viewer'
MAX_DELAY_TIME = 20
ANALYZER_TITLE = 'BZ-X800 Analyzer'
REQUIRED_FIELDS = ['Run Name', 'Stitch Type', 'Overlay', 'Naming Template', 'Filepath', 'Start Child', 'End Child', 'XY Name']

# Setup logging
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler(LOG_FILE),
                        logging.StreamHandler()
                    ])

# Global variable for channel order
channel_orders_list = []

def main():
    global channel_orders_list
    logging.info("Starting Keyence Auto Namer")
    
    # Prompt for channel order
    channel_orders_list = get_channel_orders()
    
    csv_file_path = input("Enter the path to your CSV configuration file: ")
    
    # Validate CSV
    validation_errors = validate_csv(csv_file_path)
    if validation_errors:
        logging.error("CSV Validation Errors:")
        for error in validation_errors:
            logging.error(error)
        return

    run_configs, placeholder_values = read_csv_config(csv_file_path)

    failed = []

    for run_config in run_configs:
        run_name, stitchtype, overlay, naming_template, filepath, start_child, end_child = run_config

        try:
            process_xy_sequences(failed, run_name, stitchtype, overlay, naming_template, start_child, end_child, placeholder_values[run_name], filepath)
        except Exception as e:
            logging.error(f"Failed on running {run_name}. Error: {str(e)}")

    if failed:
        logging.warning(f"All XY sequences have been processed except for: {failed}")
    else:   
        logging.info("All XY sequences have been processed successfully!")

def get_channel_orders():
    print("\nEnter the channel names in order, from first opened to last.")
    print("Press Enter on an empty line when you're finished.")
    print("Note: The last channel should typically be 'Overlay'.")
    
    channels = []
    while True:
        channel = input(f"Channel {len(channels) + 1}: ").strip()
        if not channel:
            if not channels:
                print("You must enter at least one channel.")
                continue
            break
        channels.append(channel)
    
    logging.info(f"Channel orders set: {channels}")
    return channels

def validate_csv(csv_file_path):
    errors = []

    try:
        with open(csv_file_path, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            
            # Check for required fields
            missing_fields = set(REQUIRED_FIELDS) - set(reader.fieldnames)
            if missing_fields:
                errors.append(f"Missing required fields: {', '.join(missing_fields)}")

            # Validate data in each row
            for row_num, row in enumerate(reader, start=2):
                if row['Run Name']:
                    # Validate integer fields
                    for field in ['Start Child', 'End Child']:
                        try:
                            int(row[field])
                        except ValueError:
                            errors.append(f"Row {row_num}: '{field}' must be an integer")
                    
                    # Validate Stitch Type
                    if row['Stitch Type'] not in ['F', 'L']:
                        errors.append(f"Row {row_num}: 'Stitch Type' must be 'F' or 'L'")
                    
                    # Validate Overlay
                    if row['Overlay'] not in ['Y', 'N']:
                        errors.append(f"Row {row_num}: 'Overlay' must be 'Y' or 'N'")

    except FileNotFoundError:
        errors.append(f"CSV file not found: {csv_file_path}")
    except csv.Error as e:
        errors.append(f"Error reading CSV file: {str(e)}")

    return errors

def read_csv_config(csv_file_path):
    run_configs = []
    placeholder_values = {}

    try:
        with open(csv_file_path, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            
            current_run = None
            for row in reader:
                if row['Run Name']:
                    current_run = row['Run Name']
                    run_configs.append((
                        current_run,
                        row['Stitch Type'],
                        row['Overlay'],
                        row['Naming Template'],
                        row['Filepath'],
                        int(row['Start Child']),
                        int(row['End Child'])
                    ))
                    placeholder_values[current_run] = {}
                
                if current_run:
                    xy_name = row['XY Name']
                    if xy_name:
                        placeholder_values[current_run][xy_name] = {}
                        for i in range(1, 10):
                            key = f'key{i}'
                            if key in row and row[key]:
                                placeholder_values[current_run][xy_name][key] = row[key]

        logging.info(f"Successfully read configuration from {csv_file_path}")
    except Exception as e:
        logging.error(f"Error reading CSV file: {str(e)}")
        raise

    return run_configs, placeholder_values

def process_xy_sequences(failed, run_name, stitchtype, overlay, naming_template, start_child, end_child, placeholder_values, filepath):
    logging.info(f"Starting process_xy_sequences for {run_name}")
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
            process_delay_time = max(delay_time * (len(channel_orders_list) - 0.7), delay_time)
            logging.info(f"Waiting for {process_delay_time:.2f} seconds")
            time.sleep(process_delay_time)
            disable_caps_lock()
            name_files(naming_template, placeholder_values, xy_name, delay_time, filepath)
            
            image_stitch.set_focus()
            close_button.click_input()
            logging.info("Closing stitch image...")
            
        except Exception as e:
            logging.error(f"Failed on running {xy_name}. Error: {str(e)}")
            failed.append(xy_name)

    logging.info(f"Completed processing for {run_name}")

def select_stitch_type(stitchtype):
    if stitchtype == "F":
        send_keys('f{ENTER}')
    elif stitchtype == "L":
        send_keys('l')

def check_for_image():
    assert os.path.exists(IMAGE_PATH)
    start_time = time.time()
    while True:
        try:
            location = pyautogui.locateOnScreen(IMAGE_PATH, grayscale=True, confidence=0.95)
            if location is not None:
                pyautogui.click(location)
                logging.info("Image found and clicked!")
                return
        except Exception:
            elapsed_time = round(time.time() - start_time, 0)
            logging.info(f"Time elapsed: {elapsed_time} s")
            time.sleep(2)
            if elapsed_time > 10 * 60:
                logging.error("Image not found after 10 minutes... terminating search.")
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

def name_files(naming_template, placeholder_values, xy_name, delay, filepath):
    global channel_orders_list
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

        channel = channel_orders_list[i]
        try:
            format_dict = {f'key{k}': '' for k in range(1, 10)}  # Initialize all placeholders
            format_dict.update(placeholder_values[xy_name])  # Update with actual values
            format_dict['C'] = channel

            file_name = naming_template.format(**format_dict)
            logging.info(f"Naming file: {file_name}")
            time.sleep(1)
            send_keys(file_name)
            send_keys('{TAB 2}{ENTER}')
        except KeyError as e:
            logging.error(f"Error: Missing placeholder {e} in naming template for {xy_name}")
        except Exception as e:
            logging.error(f"Unexpected error occurred while naming file for {xy_name}: {str(e)}")

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
        input()
        return get_main_window()

def run_tests():
    logging.info("Running tests...")
    
    tests_passed = True

    # Test main window connection
    try:
        main_window = get_main_window()
        assert main_window.exists(), "Main window connection failed"
        logging.info("Main window connection test passed.")
    except Exception as e:
        logging.error(f"Main window connection test failed: {str(e)}")
        tests_passed = False

    # Test CSV reading
    try:
        test_csv_path = 'test_config.csv'
        with open(test_csv_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(REQUIRED_FIELDS)
            writer.writerow(['Test Run', 'F', 'Y', '{key1}_{C}', 'C:\\Test', '1', '3', 'XY01'])
        
        errors = validate_csv(test_csv_path)
        assert not errors, f"CSV validation failed: {errors}"
        
        run_configs, placeholder_values = read_csv_config(test_csv_path)
        assert run_configs and placeholder_values, "CSV reading failed"
        
        os.remove(test_csv_path)
        logging.info("CSV reading and validation test passed.")
    except Exception as e:
        logging.error(f"CSV reading and validation test failed: {str(e)}")
        tests_passed = False

    if tests_passed:
        logging.info("All tests passed successfully.")
    else:
        logging.error("Some tests failed. Please check the logs for details.")

    return tests_passed

# Initial setup
display_splash_art()

# Run Program
if __name__ == "__main__":
    try:
        if run_tests():
            main()
        else:
            logging.error("Tests failed. Please check the logs and fix any issues before running the main program.")
    except Exception as e:
        logging.critical(f"An unexpected error occurred: {str(e)}")
    finally:
        logging.info("Program execution completed.")