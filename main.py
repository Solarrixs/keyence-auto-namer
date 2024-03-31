import pywinauto
import pyautogui
import tkinter as tk
import time
import os

def main(run_name, xy_count, naming_template, placeholder_values, channel_orders_list, output_text):
    try:
        main_window.set_focus()
        run_tree_item = main_window.child_window(title=run_name, control_type="TreeItem")
        run_tree_item.expand()
        run_tree_item.click_input()  # Click on the run name set above to select it
        children = run_tree_item.children()  # Loop on the XY sequences within the run name defined above
        for child in children:
            xy_name = child.window_text()
            output_text.insert(tk.END, f"Processing child item: {xy_name}\n")
            child.click_input()  # Click on the XY sequences
            stitch_button.click_input()  # Stitch Button
            pyautogui.press('f')  # Full Focus (Can be optionated for other features)
            pyautogui.press('enter')
            
            # Check if stitch image is ready using image1.png
            file_name_1 = os.path.join(os.path.dirname(__file__), 'image1.png')
            assert os.path.exists(file_name_1), "Image file 'image1.png' not found."
            check_for_image(file_name_1)
            time.sleep(0.1)
            
            # Uncompressed + start stitch
            pyautogui.press('tab', presses=6)
            pyautogui.press('right')
            pyautogui.press('tab', presses=3)
            pyautogui.press('enter')

            # Ensure all images are ready by guessing the delay time
            start_time = time.time()
            while True:
                matching_windows = get_matching_windows("BZ-X800 Wide Image Viewer")
                if matching_windows:
                    break
                time.sleep(0.1)
            end_time = time.time()
            delay_time = end_time - start_time
            output_text.insert(tk.END, f"Waiting for {round(delay_time * len(channel_orders_list), 2)} seconds\n")
            time.sleep(delay_time * len(channel_orders_list))
            
            name_files(naming_template, placeholder_values, xy_name, delay_time, output_text)
            
            # Click Cancel on Stitch Image
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
                
    except Exception as e:
        output_text.insert(tk.END, f"Failed on running {run_name}. Error: {e}\n")

def get_placeholder_values(xy_count, naming_template):
    """
    Prompts user to enter placeholder values for each XY sequence.
    
    :param xy_count: The number of XY sequences (just the highest XY number from your Runs).
    :param naming_template: The naming template for the files.
    :return: A dictionary containing the placeholder values for each XY sequence.
    """
    placeholder_values = {}
    for i in range(xy_count):
        xy_name = f"XY{i+1:02}"
        placeholder_values[xy_name] = {}
        num_placeholders = naming_template.count("{") - naming_template.count("{C}")
        for placeholder in range(1, num_placeholders + 1):
            value = input(f"Enter placeholder {{key{placeholder}}} for {xy_name}: ")
            placeholder_values[xy_name][f'key{placeholder}'] = value
    return placeholder_values

def define_channel(channel_count):
    """
    Prompts user to enter the channel order from first displayed to last displayed.
    
    :param channel_count: The number of channels.
    :return: A list of channel orders in reverse order.
    """
    channel_orders_list = []
    print(f"Enter the channel name type of the {channel_count} channels from opened first to opened last. The last one should be Overlay.")
    for i in range(channel_count):
        order = input(f"Channel {i+1} of {channel_count}: ")
        channel_orders_list.append(order)
    channel_orders_list.reverse()
    return channel_orders_list

def check_for_image(image_path):
    """
    Checks if an image exists on the screen by locating it using the given image path.
    
    :param image_path: The path to the image file.
    :param timeout: The maximum time (in seconds) to search for the image.
    :return: True if the image is found, False otherwise.
    """
    start_time = time.time()
    while True:
        try:
            location = pyautogui.locateOnScreen(image_path, grayscale=True, confidence=0.99)
            if location is not None:
                print("Found image!")
                pyautogui.click(location)
                return True
        except Exception:
            time.sleep(2)
        if time.time() - start_time > 300:
            print("Image not found after 300 seconds... terminating search.")
            return False
        
def name_files(naming_template, placeholder_values, xy_name, delay):
    """
    Names the files using the specified naming template and placeholder values.
    
    :param naming_template: The naming template for the files.
    :param placeholder_values: A dictionary containing the placeholder values for each XY sequence.
    :param xy_name: The name of the current XY sequence.
    :param delay: The delay time between file operations.
    """
    matching_windows = get_matching_windows("BZ-X800 Wide Image Viewer")

    for i, window in enumerate(matching_windows):
        window.set_focus()
        
        # Click File Button
        app = pywinauto.Application(backend="uia").connect(handle=window.handle)
        main_window = app.window(auto_id="MainForm", control_type="Window")
        toolbar = main_window.child_window(title="toolStrip1", control_type="ToolBar")
        file_button = toolbar.child_window(title="File", control_type="Button")
        file_button.click_input()
        
        # Click Export in Original Scale
        pyautogui.press('tab', presses=4)
        pyautogui.press('enter')
        pyautogui.press('tab', presses=1)
        pyautogui.press('enter')
        
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
        print(f"{channel} image closed.")

def get_matching_windows(title):
    """
    Filters windows with titles matching the given title.
    
    :param title: The title to match against window titles.
    :return: A list of matching windows.
    """
    windows = pywinauto.Desktop(backend="win32").windows()
    return [win for win in windows if title in win.window_text()]

# Setting up the app to use the UIA backend of BZ-X800 Analyzer and focusing on it
app = pywinauto.Application(backend="uia").connect(title="BZ-X800 Analyzer")
main_window = app.window(title="BZ-X800 Analyzer")

# Defines the ID of the "Stitch" button
stitch_button = app.window(title="BZ-X800 Analyzer").child_window(title="Stitch", class_name="WindowsForms10.BUTTON.app.0.3553390_r8_ad1")

# Setting up the Channel Orders
channel_orders_list = define_channel(int(input("How many channels were imaged? ")))
print("Channel Orders:", channel_orders_list)

# Run the main function
main()