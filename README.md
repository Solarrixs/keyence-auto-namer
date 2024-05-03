# Automated Keyence Naming Program

The Automated Keyence Naming Program is a Python script that automates the process of processing and naming image files on the Keyence analyzer. This program significantly reduces the time and effort required to process Z-stack images by automating various processing steps for multiple image sequences in a given run.

## Features

- Automates the processing of image sequences on the Keyence Analzyer Software, including the Full Focus Mode, Z-Stack Mode, and Load Mode.
- Allows customization of the naming template for the processed image files.
- Supports processing of multiple channels and XY sequences within a run.
- Significantly reduces the manual processing time from ~3.5m per image to ~5s per image, allowing you to do the work that really matters for your PhD!

## Requirements

- Python 3.x
- `pywinauto` library
- `pyautogui` library

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/Solarrixs/keyence-auto-namer.git
   ```
2. Install the required libraries:
   ```
   pip install pywinauto pyautogui
   ```

## Project Structure

The project is structured into four main files:

1. `main.py`: This file contains the `Main` class, which maintains the overall program flow. It includes methods for retrieving configuration settings, getting the XY sequence range, retrieving placeholder values for naming, and defining the channel order. The `main()` function creates an instance of the `Main` class and runs the program.

2. `config.py`: This file contains the `Config` class, which is responsible for initializing and storing the configuration parameters required for the program.

3. `keyence_analyzer.py`: This file contains the `KeyenceAnalyzer` class, which handles the interaction with the BZ-X800 Analyzer application. It includes methods for processing XY sequences, selecting the stitch type, checking for the presence of an image, starting the stitching process, disabling Caps Lock, and closing the stitch image.

4. `wide_image_viewer.py`: This file contains the `WideImageViewer` class, which handles the interaction with the "BZ-X800 Wide Image Viewer" window. It includes methods for waiting for the Wide Image Viewer window, naming files in the Wide Image Viewer, clicking the File button, exporting the image in original scale, and closing the image.

## Usage

1. Run the script:
   ```
   python main.py
   ```
2. Follow the prompts in the terminal to provide the necessary information:
   - Enter the Run Name
   - Select the Stitch Type (Full Focus Mode or Load Mode)
   - Choose whether to include the Overlay Image
   - Enter the naming template with placeholders (`{key#}` for custom values and `{C}` for channel)
   - Enter the file path to save the images (or press Enter to use the previous file path loaded by Keyence)
   - Specify the number of channels imaged and provide the channel names in the order they were opened (from first to last)
3. The program will automatically process the image sequences based on the provided information.

## Limitations and Future Updates

1. Version 1 supports only supports the Full Focus mode. Version 2 currently introduce the ability to customize the workflow based on user requirements.
2. Version 2 requires manual input of the file save path and channel order. Version 3 will include a file save path chooser and automatically determine the channel order.
3. Version 3 will include a graphical user interface (GUI) and an executable (.exe) file for ease of use by non-programmers.

## Contributing

You can contribute to the project by forking the repository at [https://github.com/Solarrixs/keyence-auto-namer](https://github.com/Solarrixs/keyence-auto-namer) or contacting Maxx Yung to be added as a collaborator.

## Contact

For any questions or suggestions, please contact Maxx Yung.