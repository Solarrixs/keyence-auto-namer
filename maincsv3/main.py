import logging
import time
import os
import pyautogui
import constants
import interactions
import utils

# Global variable for channel order
channel_orders_list = []

def main():
    global channel_orders_list
    logging.info("Starting Keyence Auto Namer")

    try:
        channel_orders_list = utils.get_channel_orders()
        
        while True:
            csv_file_path = input("\nEnter the path to your CSV configuration file: ")
            
            # Validate CSV
            validation_errors = utils.validate_csv(csv_file_path)
            if validation_errors:
                logging.error("CSV Validation Errors:")
                for error in validation_errors:
                    logging.error(error)
                print("The CSV file is invalid. Please check the following errors and try again:")
                for error in validation_errors:
                    print(f"- {error}")
                continue  # Invalid Path. Continues the loop.
            
            break # Valid CSV

        run_configs, placeholder_values = utils.read_csv_config(csv_file_path)

        failed = []

        for run_config in run_configs:
            run_name, stitchtype, overlay, naming_template, filepath, xy_name = run_config
            process_xy_sequences(failed, run_name, stitchtype, overlay, naming_template, xy_name, placeholder_values[run_name], filepath)

        if failed:
            logging.warning(f"All XY sequences have been processed except for: {failed}")
            print(f"Processing failed for the following XY sequences: {failed}")
        else:   
            logging.info("All XY sequences have been processed successfully!")

    except Exception as e:
        error_message = f"Unexpected error in main function: {str(e)}"
        utils.terminate_program(error_message)

def process_xy_sequences(failed, run_name, stitchtype, overlay, naming_template, xy_name, placeholder_values, filepath):
    logging.info(f"Starting process_xy_sequences for {run_name}")
    try:
        main_window = constants.get_main_window()
        main_window.set_focus()
        run_tree_item = main_window.child_window(title=run_name, control_type="TreeItem")
        run_tree_item.expand()
        run_tree_item.click_input()

        try:
            child = run_tree_item.child_window(title=xy_name, control_type="TreeItem")
            logging.info(f"Processing {xy_name}")
            child.click_input()
            stitch_button = main_window.child_window(title="Stitch")
            stitch_button.click_input()
            
            interactions.select_stitch_type(stitchtype)                
            check_for_image()
            
            image_stitch = main_window.child_window(auto_id="ImageJointMainForm", title="Image Stitch")
            
            close_button = image_stitch.child_window(auto_id="_buttonCancel", title="Cancel")
            
            image_stitch.set_focus()
            interactions.start_stitching(overlay)
            
            delay_time = interactions.wait_for_wide_image_viewer(constants.get_desktop())
            process_delay_time = max(delay_time * (len(channel_orders_list) - 0.7), delay_time)
            logging.info(f"Waiting for {process_delay_time:.2f} seconds")
            time.sleep(process_delay_time)
            interactions.disable_caps_lock()
            name_files(naming_template, placeholder_values, xy_name, delay_time, filepath)
            
            image_stitch.set_focus()
            close_button.click_input()
            logging.info(f"Successfully processed {xy_name}")
            print(f"{xy_name} processed.")

        except Exception as e:
            logging.error(f"Failed on running {xy_name}. Error: {str(e)}")
            failed.append(xy_name)

    except Exception as e:
        error_message = f"Unexpected error in process_xy_sequences: {str(e)}"
        utils.terminate_program(error_message)

    logging.info(f"Completed processing for {run_name}")
    print(f"Completed processing for {run_name}")

def name_files(naming_template, placeholder_values, xy_name, delay, filepath):
    global channel_orders_list
    windows = constants.get_desktop().windows()
    matching_windows = [win for win in windows if constants.WIDE_IMAGE_VIEWER_TITLE in win.window_text()]

    reversed_channels = channel_orders_list[::-1]

    for i in range(len(matching_windows)):
        last_window = matching_windows[i]
        last_window.set_focus()

        interactions.click_file_button(last_window)
        interactions.export_in_original_scale()
        
        save_as_dialog = constants.get_desktop().window(title="Save As")
        address_bar = save_as_dialog.child_window(title="Address", control_type="Edit")
        filename_field = save_as_dialog.child_window(title="File name:", control_type="Edit")
        
        if i == 0 and filepath != "":
            while True:
                try:
                    address_bar.set_focus()
                    address_bar.set_edit_text(filepath)
                    pyautogui.press('enter')
                    logging.info(f"Filepath set to: {filepath}")
                    break
                except Exception as e:
                    logging.error(f"Error setting filepath: {str(e)}")
                    print(f"An error occurred while setting the filepath to: {filepath}")
                    input("Please fix the issue manually and press Enter to continue with the program.")
                    logging.info("User manually fixed the filepath issue. Continuing with the program.")
                    break

        channel = reversed_channels[i]
        try:
            format_dict = placeholder_values.get(xy_name, {})
            format_dict['C'] = channel
            file_name = naming_template.format(**format_dict)
            
            logging.info(f"Naming file: {file_name}")
            filename_field.set_focus()
            time.sleep(1)
            pyautogui.write(file_name)
            pyautogui.press('tab', presses=2)
            time.sleep(0.1)
            pyautogui.press('enter')
        except KeyError as e:
            logging.error(f"Error: Missing placeholder {e} in naming template for {xy_name}")
        except Exception as e:
            logging.error(f"Unexpected error occurred while naming file for {xy_name}: {str(e)}")

        interactions.close_image(delay, channel)
        
def check_for_image():
    logging.info(f"Image path exists: {os.path.exists(constants.IMAGE_PATH)}")
    start_time = time.time()
    while True:
        try:
            location = pyautogui.locateOnScreen(constants.IMAGE_PATH, grayscale=True, confidence=0.95)
            if location is not None:
                pyautogui.click(location)
                print("Image found!")
                logging.info("Image found!")
                return
        except Exception:
            logging.info(f"Time elapsed: {round(time.time() - start_time, 0)} s")
            time.sleep(2)
            if time.time() - start_time > 10 * 60:
                logging.error("Image not found after 10 minutes... terminating search.")
                print("Image not found after 10 minutes... terminating search.")
                raise TimeoutError("Image not found after 10 minutes")

if __name__ == "__main__":
    try:
        utils.display_splash_art()
        if utils.run_tests():
            main()
        else:
            logging.error("Tests failed. Please check the logs and fix any issues before running the main program.")
    except Exception as e:
        logging.critical(f"An unexpected error occurred: {str(e)}")
    finally:
        logging.info("Program execution completed.")