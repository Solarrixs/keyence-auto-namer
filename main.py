import pywinauto
import pyautogui
import time
import os

def main():
    run_name = input("Enter Run Name: ")
    xy_count = int(input("Enter the number of XY sequences: "))
    naming_template = input("Enter the naming template (use {1}, {2}, etc. for placeholders and {C} for channel): ")

    # Code for naming the XY sequences with your custom naming template defined above
    placeholder_values = {}
    for i in range(xy_count):
        xy_name = f"XY{i+1:02}"
        placeholder_values[xy_name] = {}
        for placeholder in range(1, naming_template.count("{") - naming_template.count("{C}") + 1):
            value = input(f"Enter placeholder {{{placeholder}}} for {xy_name}: ")
            placeholder_values[xy_name][placeholder] = value

    try:
        main_window.set_focus()
        run_tree_item = main_window.child_window(title=run_name, control_type="TreeItem")
        run_tree_item.expand()
        run_tree_item.click_input() # Click on the run name set above to select it
        children = run_tree_item.children() # Loop on the XY sequences within the run name defined above
        for child in children:
            xy_name = child.window_text()
            print(f"Child item: {xy_name}")
            child.click_input() # Cick on the XY sequences
            stitch_button.click_input() # Stitch Button
            pyautogui.press('f') # Full Focus
            pyautogui.press('enter')
            file_name = os.path.join(os.path.dirname(__file__), 'image1.png')
            assert os.path.exists(file_name)
            check_for_image(file_name)
            time.sleep(2)
            pyautogui.press('tab', presses=6)
            pyautogui.press('right')
            pyautogui.press('tab', presses=3)
            pyautogui.press('enter')
            window_location = locate_new_window('image2.png')
            if window_location:
                name_files(naming_template, placeholder_values, xy_name, window_location)
            else:
                print("Matching window not found.")
    except Exception as e:
        print(f"Failed on running {run_name}. Error: {e}")

def name_files(naming_template, placeholder_values, xy_name, window_location):
    file_button_location = pyautogui.locateOnScreen('file_button.png', confidence=0.9, region=window_location)
    if file_button_location:
        file_button_center = pyautogui.center(file_button_location)
        pyautogui.click(file_button_center)
        pyautogui.press('tab', presses=5)
        pyautogui.press('enter')
    else:
        print("File button not found within the window.")
    for channel in channel_orders_list:
        file_name = naming_template.format(**placeholder_values, C=channel)
        print(f"Naming file for {xy_name} - {channel}: {file_name}")
        # Add code here to actually rename the files using the generated file_name
    exit_button_location = pyautogui.locateOnScreen('exit_button.png', confidence=0.9, region=window_location)
    if exit_button_location:
        exit_button_center = pyautogui.center(exit_button_location)
        pyautogui.click(exit_button_center)
        pyautogui.press('tab', presses=1)
        pyautogui.press('enter')
    else:
        print("Exit button not found within the window.")

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
    while True:
        try:
            location = pyautogui.locateOnScreen(image_path, grayscale=True, confidence=0.99)
            if location is not None:
                print("Found image!")
                pyautogui.click(location)
                return False
        except Exception:
            print("Image not found, trying again...",)
            time.sleep(2)
        if time.time() - start_time > 5 * 60:  # 5 minutes
            print("Image not found after 5 minutes... terminating search.")
            return False
        
def locate_new_window(reference_image):
    previous_matches = []
    while True:
        current_matches = list(pyautogui.locateAllOnScreen(reference_image))
        if current_matches == previous_matches:
            break
        previous_matches = current_matches
        time.sleep(0.5)
    return previous_matches[-1]  # Return the last matched window

# Setting up the app to use the UIA backend of BZ-X800 Analyzer and focusing on it
app = pywinauto.Application(backend="uia").connect(title="BZ-X800 Analyzer") #(process=21140)
main_window = app.window(title="BZ-X800 Analyzer")

# Defines the ID of the "Stitch" button
stitch_button = app.window(title="BZ-X800 Analyzer").child_window(title="Stitch", class_name="WindowsForms10.BUTTON.app.0.3553390_r8_ad1")

# Setting up the Channel Orders
channel_orders_list = defineChannel(int(input("How many channels were imaged? ")))
print("Channel Orders:", channel_orders_list)

# Run the main function
main()