import pywinauto
import pyautogui

def set_up_app():
    app = pywinauto.Application(backend="uia").connect(title="BZ-X800 Analyzer")
    main_window = app.window(title="BZ-X800 Analyzer")
    stitch_button = main_window.child_window(title="Stitch", class_name="WindowsForms10.BUTTON.app.0.3553390_r8_ad1")
    return main_window, stitch_button

def define_channel_orders():
    channel_count = int(input("How many channels were imaged? "))
    channel_orders_list = []
    print(f"Enter the channel name type of the {channel_count} channels from opened first to opened last. The last one should be Overlay if you want an overlay image.")
    for i in range(channel_count):
        order = input(f"Channel {i+1} of {channel_count}: ")
        channel_orders_list.append(order)
    channel_orders_list.reverse()
    print("Channel Orders:", channel_orders_list)
    return channel_orders_list

def process_xy_sequence(main_window, run_name, xy_name, stitch_button, stitch_type, overlay):
    main_window.set_focus()
    run_tree_item = main_window.child_window(title=run_name, control_type="TreeItem")
    run_tree_item.expand()
    run_tree_item.click_input()
    child = run_tree_item.child_window(title=xy_name, control_type="TreeItem")
    print(f"Processing {xy_name}")
    child.click_input()
    stitch_button.click_input()
    
    if stitch_type == "Full":
        pyautogui.press('f')
        pyautogui.press('enter')
    elif stitch_type == "Load":
        pyautogui.press('l')
    
    pyautogui.press('tab', presses=6)
    pyautogui.press('right')
    
    if overlay == "Y":
        pyautogui.press('tab', presses=3)
    elif overlay == "N":
        pyautogui.press('tab', presses=2)
        
    pyautogui.press('enter')

def export_image(window):
    app = pywinauto.Application(backend="uia").connect(handle=window.handle)
    main_window = app.window(auto_id="MainForm", control_type="Window")
    toolbar = main_window.child_window(title="toolStrip1", control_type="ToolBar")
    file_button = toolbar.child_window(title="File", control_type="Button")
    file_button.click_input()
    
    pyautogui.press('tab', presses=4)
    pyautogui.press('enter')
    pyautogui.press('tab', presses=1)
    pyautogui.press('enter')