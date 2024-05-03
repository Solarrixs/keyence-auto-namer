import os
from config import Config
from keyence_analyzer import KeyenceAnalyzer
from wide_image_viewer import WideImageViewer

class Main:
    def __init__(self):
        """
        Initializes the Main class by creating instances of the KeyenceAnalyzer and WideImageViewer classes.
        """
        self.keyence_analyzer = KeyenceAnalyzer()
        self.wide_image_viewer = WideImageViewer()

    def run(self):
        """
        Runs the main function of the program. Retrieves core configuration settings by using the `get_config` method of the `self` object.
        """
        config = self.get_config() # Gets the core configuration settings
        start_child, end_child = self.get_xy_sequence_range(config.run_name) # Gets the XY sequence range
        placeholder_values = self.get_placeholder_values(config.naming_template, start_child, end_child) # Gets the placeholder values for each XY sequence
        channel_orders_list = self.define_channel_orders() # Gets the channel orders

        self.keyence_analyzer.process_xy_sequences(config, start_child, end_child, placeholder_values, channel_orders_list, self.wide_image_viewer) # Processes the XY sequences through the KeyenceAnalyzer class

    def get_config(self):
        """
        Retrieves core configuration settings by prompting the user for input.
        """
        run_name = input("Enter Run Name: ") # Gets the run name to be processed. Version 3 will contain a GUI to select multiple runs
        stitchtype = input("Stitch Type? Full (F) or Load (L): ").upper()
        overlay = input("Overlay Image? (Y/N): ").upper() # Determines if the overlay image should be used and changes the stitch type accordingly
        naming_template = input("Enter the naming template (use {key1}, {key2}, etc. for placeholders and {C} for channel): ")
        filepath = input("Enter the EXACT filepath to save the images. Press ENTER if you want to use the previous filepath loaded by Keyence: ")
        image_path = os.path.join(os.path.dirname(__file__), 'image.png') # Used by KeyenceAnalyzer to determine whether to begin the stitching process
        image_path_2 = os.path.join(os.path.dirname(__file__), 'image2.png') # Used by KeyenceAnalyzer to locate the overlay button during the stitching process
        max_delay_time = 20

        return Config(run_name, stitchtype, overlay, naming_template, filepath, image_path, image_path_2, max_delay_time)

    def get_xy_sequence_range(self, run_name):
        """
        Retrieves the range of XY sequences for a given run name.

        Parameters:
            run_name (str): The name of the run.

        Returns:
            tuple: A tuple containing the starting and ending XY numbers.
        """
        run_tree_item = self.keyence_analyzer.main_window.child_window(title=run_name, control_type="TreeItem")
        children = run_tree_item.children() # Reads the XY sequences from the run
        print(f"Detected {len(children)} XY sequences.")
        start_child = int(input("Enter the starting XY number to be processed: "))
        end_child = int(input("Enter the ending XY number to be processed: "))
        return start_child, end_child

    def get_placeholder_values(self, naming_template, start_child, end_child):
        """
        Get placeholder values for each XY sequence based on the naming template.

        Parameters:
            naming_template (str): The naming template for the file names.
            start_child (int): The starting XY sequence number.
            end_child (int): The ending XY sequence number.

        Returns:
            dict: A dictionary containing placeholder values for each XY sequence. The keys are the XY names and the values are dictionaries containing the placeholder values.
        """
        placeholder_values = {}
        xy_names = [f"XY{i+1:02}" for i in range(start_child - 1, end_child)]

        for placeholder in range(1, naming_template.count("{") - naming_template.count("{C}") + 1):
            for xy_name in xy_names:
                if xy_name not in placeholder_values:
                    placeholder_values[xy_name] = {}
                value = input(f"Enter placeholder {{key{placeholder}}} for {xy_name}: ")
                placeholder_values[xy_name][f'key{placeholder}'] = value

        return placeholder_values

    def define_channel_orders(self):
        """
        Defines the order of the channels that were imaged.
        
        Returns:
            channel_orders_list (list): A list containing the order of the channels that were imaged.
        """
        channel_count = int(input("How many channels were imaged? "))
        channel_orders_list = []
        print(f"Enter the channel name type of the {channel_count} channels from opened first to opened last. The last one should be Overlay.")
        for i in range(channel_count):
            order = input(f"Channel {i+1} of {channel_count}: ")
            channel_orders_list.append(order)
        channel_orders_list.reverse()
        return channel_orders_list

def main():
    main_instance = Main()
    main_instance.run()

if __name__ == "__main__":
    main()