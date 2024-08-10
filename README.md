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
   - Number of channels in your images.
   - Filepath of the CSV file.

3. Setup the CSV file using the provided CSV template.

4. The program will automatically process the image sequences based on the provided information.

## CSV File Structure
The CSV file should contain the following columns:
1. Run Name
2. Stitch Type
3. Overlay
4. Naming Template
5. Filepath
8. XY Name
9. key1, key2, key3, etc. (as needed)

## Column Descriptions
- **Run Name**: Name of the run (required for each new run)
- **Stitch Type**: Either 'F' for Full or 'L' for Load
- **Overlay**: 'Y' for Yes or 'N' for No
- **Naming Template**: Template for file naming (e.g., "{key1}{key2}{C}")
- **Filepath**: Path to save the images
- **XY Name**: Name of the XY sequence (typically XY01, XY02, XY03...)
- **key1, key2, etc.**: Placeholder values for naming template

## Example CSV Structure

Here's an example of how your CSV might look:

| Run Name | Stitch Type | Overlay | Naming Template | Filepath      | XY Name |key1 | key2  |
|----------|-------------|---------|-----------------|---------------|---------|-----|-------|
| Run1     | F           | Y       | {key1}{key2}{C} | C:\Images\Run1| XY01    | ex1 | ex1.1 |
|          |             |         |                 |               | XY02    | ex2 | ex2.1 |
|          |             |         |                 |               | XY03    | ex3 | ex3.1 |
| Run2     | L           | N       | {key1}_{C}      | D:\Images\Run2| XY01    | ex  |       |
|          |             |         |                 |               | XY02    | ex2 |       |

## Limitations

1. This program cannot handle duplicated file names. Make sure that the folder does not contain any duplicated files, otherwise your files will be overwritten.
2. This program cannot handle setting the filepath, which must be setup before by naming once. This is planned to be fixed in a future update.

## Contributing

You can contribute to the project by forking the repository at [https://github.com/Solarrixs/keyence-auto-namer](https://github.com/Solarrixs/keyence-auto-namer) or contacting Maxx Yung to be added as a collaborator.

## Contact

For any questions or suggestions, please contact Maxx Yung.