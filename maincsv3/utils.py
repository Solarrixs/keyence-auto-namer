import sys
import logging
import time
import csv
from constants import REQUIRED_FIELDS

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
            current_run_config = None
            
            for row in reader:
                if row['Run Name']:
                    current_run = row['Run Name']
                    current_run_config = (
                        current_run,
                        row['Stitch Type'],
                        row['Overlay'],
                        row['Naming Template'],
                        row['Filepath']
                    )
                    placeholder_values[current_run] = {}
                
                if current_run:
                    xy_name = row['XY Name']
                    if xy_name:
                        run_configs.append(current_run_config + (xy_name,))
                        placeholder_values[current_run][xy_name] = {}
                        for key, value in row.items():
                            if key.startswith('key') and value:
                                placeholder_values[current_run][xy_name][key] = value

        logging.info(f"Successfully read configuration from {csv_file_path}")
    except Exception as e:
        logging.error(f"Error reading CSV file: {str(e)}")
        raise

    return run_configs, placeholder_values

def terminate_program(error_message):
    logging.critical(f"Critical error occurred: {error_message}")
    logging.info("Terminating program due to critical error.")
    print("A critical error occurred. Program terminated. Check the log file for details.")
    sys.exit(1)

def get_channel_orders():
    print("\nEnter the channel names in order, from first opened to last.")
    print("Press Enter on an empty line when you're finished.")
    print("Note: The last channel should typically be 'Overlay'.\n")
    
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

def run_tests():
    logging.info("Running tests...")
    
    tests_passed = True

    # Test main window connection
    try:
        from constants import get_main_window
        main_window = get_main_window()
        assert main_window.exists(), "Main window connection failed"
        logging.info("Main window connection test passed.")
    except Exception as e:
        logging.error(f"Main window connection test failed: {str(e)}")
        tests_passed = False

    # Test CSV reading
    try:
        import csv
        import os
        from csv import validate_csv, read_csv_config
        from constants import REQUIRED_FIELDS
        
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