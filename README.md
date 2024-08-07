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
   ```bash
   pip install pywinauto pyautogui
   ```

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

## CSV File Structure
The CSV file should contain the following columns:
1. Run Name
2. Stitch Type
3. Overlay
4. Naming Template
5. Filepath
6. Start Child
7. End Child
8. XY Name
9. key1, key2, key3, etc. (as needed)

## Column Descriptions
- **Run Name**: Name of the run (required for each new run)
- **Stitch Type**: Either 'F' for Full or 'L' for Load
- **Overlay**: 'Y' for Yes or 'N' for No
- **Naming Template**: Template for file naming (e.g., "{key1}_{key2}_{C}")
- **Filepath**: Path to save the images
- **Start Child**: Starting XY number (integer)
- **End Child**: Ending XY number (integer)
- **XY Name**: Name of the XY sequence (e.g., XY01, XY02)
- **key1, key2, etc.**: Placeholder values for naming template

## Example CSV Structure

Here's an example of how your CSV might look:

| Run Name | Stitch Type | Overlay | Naming Template | Filepath      | Start Child | End Child | XY Name | key1   | key2 |
|----------|-------------|---------|-----------------|---------------|-------------|-----------|---------|--------|------|
| Run1     | F           | Y       | {key1}_{key2}_{C} | C:\Images\Run1 | 1           | 3         |         |        |      |
|          |             |         |                 |               |             |           | XY01    | sample1| day1 |
|          |             |         |                 |               |             |           | XY02    | sample2| day1 |
|          |             |         |                 |               |             |           | XY03    | sample3| day1 |
| Run2     | L           | N       | {key1}_{C}       | D:\Images\Run2 | 1           | 2         |         |        |      |
|          |             |         |                 |               |             |           | XY01    | exp2   |      |
|          |             |         |                 |               |             |           | XY02    | exp2   |      |

## Limitations and Future Updates

1. Version 1 currently supports only the Stitch, Full Focus, and Uncompressed workflow. Version 2 will introduce the ability to customize the workflow based on user requirements.

2. Version 1 requires manual input of the file save path and channel order. Version 2 will include a file save path chooser and automatically determine the channel order.

3. Version 1.0.2 allows for an executable (.exe) file to be used instead.

4. Version 2 will include a graphical user interface (GUI).

## Contributing

You can contribute to the project by forking the repository at [https://github.com/Solarrixs/keyence-auto-namer](https://github.com/Solarrixs/keyence-auto-namer) or contacting Maxx Yung to be added as a collaborator.

## Contact

For any questions or suggestions, please contact Maxx Yung.