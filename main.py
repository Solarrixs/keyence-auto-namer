import pywinauto
import pyautogui
import time

def main():
    run_name = input("Enter Run Name: ")
    xy_count = int(input("Enter the number of XY sequences: "))
    naming_template = input("Enter the naming template (use {1}, {2}, etc. for placeholders and {C} for channel): ")

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
        run_tree_item.click_input()
        children = run_tree_item.children()
        for child in children:
            xy_name = child.window_text()
            print(f"Child item: {xy_name}")
            child.click_input()
            stitch_button.click_input()
            pyautogui.press('f')
            pyautogui.press('enter')
            if check_for_image('stitchimage.png'):
                pyautogui.press('tab', presses=6)
                pyautogui.press('right')
                pyautogui.press('tab', presses=3)
                pyautogui.press('enter')
                name_files(naming_template, placeholder_values[xy_name], xy_name)
            else:
                print("Image not found within the time limit.")
                break
    except Exception as e:
        print(f"Failed on running {run_name}. error: {e}")

def name_files(naming_template, placeholder_values, xy_name):
    for channel in channel_orders_list:
        file_name = naming_template.format(**placeholder_values, C=channel)
        print(f"Naming file for {xy_name} - {channel}: {file_name}")
        # Add code here to actually rename the files using the generated file_name

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
            location = pyautogui.locateOnScreen(image_path)
            if location is not None:
                # Click on the top-left corner of the image
                pyautogui.click(location[0], location[1])
                return True
            else:
                print("Image not found")
        except Exception as e:
            print(f"Error: {e}")
            return False

        # Sleep for 2 seconds
        time.sleep(2)

        # If 5 minutes have passed, stop checking
        if time.time() - start_time > 5 * 60:
            break

    return False

# Setting up the Channel Orders
channel_orders_list = defineChannel(int(input("How many channels were imaged? ")))
print("Channel Orders:", channel_orders_list)

# Setting up the app to use the UIA backend of BZ-X800 Analyzer and focusing on it
app = pywinauto.Application(backend="uia").connect(process=6232)
main_window = app.window(title="BZ-X800 Analyzer")

# Defines the ID of the "Stitch" button
stitch_button = app.window(title="BZ-X800 Analyzer").child_window(title="Stitch", class_name="WindowsForms10.BUTTON.app.0.3553390_r8_ad1")

# Run the main function
main()