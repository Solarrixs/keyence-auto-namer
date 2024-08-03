# Keyence Auto Namer: CSV Structure Guide

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