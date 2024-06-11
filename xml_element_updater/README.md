# XML Element Updater

This script updates specified XML elements in multiple XML files within a folder.

## Overview

This Python script searches for XML files in a specified folder and updates specific XML elements with new text. It then saves the modified XML files to another folder.

## Usage

1. Ensure you have Python installed on your system.
2. Install the required dependencies by running `pip install lxml`.
3. Place the XML files you want to update in the input folder.
4. Modify the script to specify the element name and the new text you want to set for that element.
5. Run the script by executing `python update_xml_elements.py`.
6. Check the output folder for the modified XML files.

## Script Details

- **Input Folder**: The folder containing the XML files to be updated.
- **Output Folder**: The folder where the modified XML files will be saved.
- **Element Name**: The name of the XML element to update.
- **New Text**: The new text to set for the specified XML elements.

## Example

Suppose you have several XML files in the input folder and you want to update the `<defautPagCycle>` elements with the text `defaultPagCycle_rf128`. After running the script, all `<defautPagCycle>` elements in the XML files will be updated accordingly.

## Dependencies

- Python (3.x recommended)
- lxml library

