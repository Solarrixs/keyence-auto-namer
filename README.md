# Automated Keyence Naming Program

The Automated Keyence Naming Program is a Python script that automates the process of processing and naming image files on the Keyence analyzer. This program significantly reduces the time and effort required to process Z-stack images by automating the Stitch, Full Focus, and Uncompressed steps for multiple image sequences in a given run.

## Features

- Automates the processing of image sequences, including Stitch, Full Focus, and Uncompressed steps
- Allows customization of the naming template for the processed image files
- Supports processing of multiple channels and XY sequences within a run
- Significantly reduces the processing time compared to manual processing

## Requirements

- Python 3.x
- `pywinauto` library
- `pyautogui` library

## Installation

1. Clone the repository:
   `git clone https://github.com/Solarrixs/keyence-auto-namer.git`

2. Install the required libraries:
   `pip install pywinauto pyautogui`

## Usage

1. Run the script:
   `python main.py`

2. Follow the prompts in the terminal to provide the necessary information:
   - Number of channels in your images
   - Channel names (from Keyence first-opened to last-opened channel)
   - Name of the Run folder to process
   - Number of XY sequences under the Run folder
   - Naming template with placeholders (`{key#}` for custom values and `{C}` for channel)
   - Placeholder values for each key and XY sequence

3. The program will automatically process the image sequences based on the provided information.

4. To bundle the program within a executable file, run: `pyinstaller --hidden-import comtypes --hidden-import comtypes.stream -F main.py`

## Limitations and Future Updates

1. Version 1 currently supports only the Stitch, Full Focus, and Uncompressed workflow. Version 2 will introduce the ability to customize the workflow based on user requirements.

2. Version 1 requires manual input of the file save path and channel order. Version 2 will include a file save path chooser and automatically determine the channel order.

3. Version 1.0.2 allows for an executable (.exe) file to be used instead.

4. Version 2 will include a graphical user interface (GUI).

## Contributing

You can contribute to the project by forking the repository at [https://github.com/Solarrixs/keyence-auto-namer](https://github.com/Solarrixs/keyence-auto-namer) or contacting Maxx Yung to be added as a collaborator.

## Contact

For any questions or suggestions, please contact Maxx Yung.