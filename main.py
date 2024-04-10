import pywinauto
import pyautogui
import time
import os
import sys

def main():
    run_name = input("Enter Run Name: ")
    xy_count = int(input("Enter the number of XY sequences: "))
    naming_template = input("Enter the naming template (use {key1}, {key2}, etc. for placeholders and {C} for channel): ")
    stitchtype = input("Stitch Type 'Full' or 'Load': ")
    overlay  = input("Overlay Image? (Y/N): ")

    # Code for naming the XY sequences with your custom naming template defined above
    placeholder_values = {}
    for i in range(xy_count):
        xy_name = f"XY{i+1:02}"
        placeholder_values[xy_name] = {}
        for placeholder in range(1, naming_template.count("{") - naming_template.count("{C}") + 1):
            value = input(f"Enter placeholder {{key{placeholder}}} for {xy_name}: ")
            placeholder_values[xy_name][f'key{placeholder}'] = value

    try:
        main_window.set_focus()
        run_tree_item = main_window.child_window(title=run_name, control_type="TreeItem")
        run_tree_item.expand()
        run_tree_item.click_input() # Click on the run name set above to select it
        children = run_tree_item.children() # Loop on the XY sequences within the run name defined above
        for child in children:
            xy_name = child.window_text()
            print(f"Processing {xy_name}")
            child.click_input() # Click on the XY sequences
            stitch_button.click_input() # Stitch Button
            
            if stitchtype == "Full":
                pyautogui.press('f')
                pyautogui.press('enter')
            elif stitchtype == "Load":
                pyautogui.press('l')
            
            # Check if stitch image is ready
            file_name_1 = os.path.join(os.path.dirname(__file__), 'image.png')
            assert os.path.exists(file_name_1)
            check_for_image(file_name_1)
            time.sleep(0.1)
            
            # Uncompressed + start stitch
            pyautogui.press('tab', presses=6)
            pyautogui.press('right')
            if overlay == "Y":
                pyautogui.press('tab', presses=3)
            elif overlay == "N":
                pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
            time.sleep(1)

            start_time = time.time()
            while True:
                windows = pywinauto.Desktop(backend="win32").windows()
                matching_windows = [win for win in windows if "BZ-X800 Wide Image Viewer" in win.window_text()]
                if matching_windows:  # When the first instance of the Wide Image Viewer is found, break the loop and check the time
                    break
                time.sleep(0.1)
            end_time = time.time()
            delay_time = end_time - start_time
            print("Waiting for", round(delay_time*len(channel_orders_list), 2), "seconds")
            time.sleep(delay_time*len(channel_orders_list))
            
            name_files(naming_template, placeholder_values, xy_name, delay_time)
            
            # Click Cancel on Stitch Image
            time.sleep(delay_time)
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
                
    except Exception as e:
        print(f"Failed on running {run_name}. Error: {e}")

def defineChannel(channel_count):
    channel_orders_list = []
    print(f"Enter the channel name type of the {channel_count} channels from opened first to opened last. The last one should be Overlay.")
    for i in range(channel_count):
        order = input(f"Channel {i+1} of {channel_count}: ")
        channel_orders_list.append(order)
    channel_orders_list.reverse()
    return channel_orders_list

def check_for_image(image_path):
    # Checks if an image exists on the screen by locating it using the given image path.
    start_time = time.time()
    print(image_path)
    while True:
        try:
            location = pyautogui.locateOnScreen(image_path, grayscale = True, confidence = 0.9)
            if location is not None:
                print("Found image!")
                pyautogui.click(location)
                return False
        except Exception:
            print("Time elapsed: ", round(time.time() - start_time, 0), " s")
            time.sleep(2)
            if time.time() - start_time > 10 * 60:  # 5 minutes
                print("Image not found after 10 minutes... terminating search.")
                sys.exit()
        
def name_files(naming_template, placeholder_values, xy_name, delay):
    # Filter windows with titles matching Wide Image Viewer
    windows = pywinauto.Desktop(backend="win32").windows()
    matching_windows = [win for win in windows if "BZ-X800 Wide Image Viewer" in win.window_text()]

    for i in range(len(matching_windows)):
        # Select the last window from the filtered list
        last_window = matching_windows[i]
        last_window.set_focus()
        
        # Click File Button
        app = pywinauto.Application(backend="uia").connect(handle=last_window.handle)
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
        print(channel + " image closed.")

# Setting up the app to use the UIA backend of BZ-X800 Analyzer and focusing on it
app = pywinauto.Application(backend="uia").connect(title="BZ-X800 Analyzer")
main_window = app.window(title="BZ-X800 Analyzer")

# Defines the ID of the "Stitch" button
stitch_button = app.window(title="BZ-X800 Analyzer").child_window(title="Stitch", class_name="WindowsForms10.BUTTON.app.0.3553390_r8_ad1")

# Setting up the Channel Orders
channel_orders_list = defineChannel(int(input("How many channels were imaged? ")))
print("Channel Orders:", channel_orders_list)

# Run the main function
main()