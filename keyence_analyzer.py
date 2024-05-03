import pywinauto
import pyautogui
import time
import os
import sys

class KeyenceAnalyzer:
    def __init__(self):
        """
        Initializes the KeyenceAnalyzer class.

        This function establishes a connection to the "BZ-X800 Analyzer" application using the pywinauto library. Necessary for the entire program to run.
        """
        self.app = pywinauto.Application(backend="uia").connect(title="BZ-X800 Analyzer")
        self.main_window = self.app.window(title="BZ-X800 Analyzer")
        self.stitch_button = self.main_window.child_window(title="Stitch")

    def process_xy_sequences(self, config, start_child, end_child, placeholder_values, channel_orders_list, wide_image_viewer):
        """
        Processes XY sequences based on the provided configuration and parameters.
        
        Parameters:
            config: Description of the configuration parameter.
            start_child: The starting XY child parameter.
            end_child: The ending XY child parameter.
            placeholder_values: The placeholder values parameter.
            channel_orders_list: The channel orders list parameter.
            wide_image_viewer: Uses the wide image viewer class.
        """
        self.main_window.set_focus()
        run_tree_item = self.main_window.child_window(title=config.run_name, control_type="TreeItem")
        run_tree_item.expand()
        run_tree_item.click_input() # Expands the run tree

        for i in range(start_child - 1, end_child): # Processes each XY sequence as defined by start_child and end_child
            try:
                child = run_tree_item.children()[i]
                xy_name = child.window_text()
                print(f"Processing {xy_name}")
                child.click_input()
                self.stitch_button.click_input()

                if config.overlay == "N": # No overlay
                    assert os.path.exists(config.IMAGE_PATH_2)
                    try:
                        location = pyautogui.locateOnScreen(config.IMAGE_PATH_2, grayscale=True, confidence=0.98) # Tries to find the overlay button
                        if location is not None:
                            pyautogui.click(location) # Clicks the overlay button to disable it
                    finally:
                        print("Disabled Overlay!")

                self.select_stitch_type(config.stitchtype)
                self.check_for_image(config.IMAGE_PATH)
                time.sleep(2) # Max delay time to let the image load
                self.start_stitching(config.overlay)

                delay_time = wide_image_viewer.wait_for_viewer(config.MAX_DELAY_TIME)
                process_delay_time = delay_time * (len(channel_orders_list) + 0.5)
                print(f"Waiting for {process_delay_time:.2f} seconds")
                time.sleep(process_delay_time)

                self.disable_caps_lock() # Ensures proper naming of files
                wide_image_viewer.name_files(config.naming_template, placeholder_values, xy_name, delay_time, config.filepath, channel_orders_list)

                self.close_stitch_image(delay_time)

            except Exception as e:
                print(f"Failed on running {xy_name}. Error: {e}")

    def select_stitch_type(self, stitchtype):
        """
        Selects the stitch type based on the input parameter.
        """
        # Full Focus Mode
        if stitchtype == "F":
            pyautogui.press('f')
            pyautogui.press('enter')
        # Load Mode
        elif stitchtype == "L":
            pyautogui.press('l')

    def check_for_image(self, image_path):
        """
        Checks if the specified image exists and waits until it is found on the screen.  image_path is image1.png.
        """
        assert os.path.exists(image_path)
        start_time = time.time()
        print(image_path)
        while True:
            try:
                location = pyautogui.locateOnScreen(image_path, grayscale=True, confidence=0.95) # Tries to find the image to begin stitching
                if location is not None:
                    pyautogui.click(location)
                    print("Found image!")
                    return
            except Exception:
                print("Time elapsed:", round(time.time() - start_time, 0), "s")
                time.sleep(10)
                if time.time() - start_time > 10 * 60: # 10 minutes
                    print("Image not found after 10 minutes... terminating search.") # Likely indictates corrupted image
                    sys.exit()

    def start_stitching(self, overlay):
        """
        Starts the stitching process by simulating keyboard input. Is the fastest way to start the stitching process.
        """
        pyautogui.press('tab', presses=6)
        pyautogui.press('right')
        if overlay == "Y":
            pyautogui.press('tab', presses=3)
        elif overlay == "N":
            pyautogui.press('tab', presses=2)
        pyautogui.press('enter')

    def disable_caps_lock(self):
        """
        Uses the `ctypes` library to interact with the Windows API. It checks the state of the Caps Lock key using the `GetKeyState` function from the `User32.dll` library. If the Caps Lock key is enabled, it simulates a key press by using the `pyautogui.press` function to press the Caps Lock key.
        """
        import ctypes
        hllDll = ctypes.WinDLL ("User32.dll")
        VK_CAPITAL = 0x14
        if hllDll.GetKeyState(VK_CAPITAL):
            pyautogui.press('capslock')

    def close_stitch_image(self, delay):
        """
        Closes the stitch image after a specified delay.
        """
        time.sleep(delay) # Delay calculated from max delay time (above)
        pyautogui.press('tab', presses=2)
        time.sleep(2) # Max delya time to let the image load
        pyautogui.press('enter')