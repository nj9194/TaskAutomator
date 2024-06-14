# Closest Locations Finder for RSI

## Description
This tool calculates the closest location to a point of interest (POI) based on latitude and longitude coordinates using the Haversine formula. It reads data from an Excel file, filters the locations by the selected technology (LTE, 5GNR, or both), and iterates over a range of PRACH Root Sequences Index (RSI) values to find the closest location for each RSI. The results are saved to an Excel file.

## Features
- Calculates distance using the Haversine formula.
- Filters results by technology (LTE, 5GNR, or both).
- Allows user to specify the number of PRACH Root Sequences Index values to iterate over (default is 891).
- Saves results to an Excel file with a timestamp.

## Requirements
- Python 3.x
- pandas
- openpyxl
- tkinter

## Installation
1. Ensure you have Python 3.x installed.
2. Install the required packages using pip:
    ```
    pip install pandas openpyxl
    ```

## Usage
1. Run the script:
    ```
    python RSI_closest_locations_finder.py
    ```
2. Follow the instructions in the GUI to:
   - Enter the latitude and longitude of the point of interest.
   - Select the input Excel file.
   - Select the output folder.
   - Specify the number of PRACH Root Sequences Index values to iterate over (default is 891).
   - Choose the technology (LTE, 5GNR, or Both).
   - Click "Process" to calculate the closest locations.

3. The results will be saved to an Excel file in the selected output folder.

Sample Input File (input_data.xlsx)
This file will contain sample location data with columns: SITE_NAME, LATITUDE, LONGITUDE, TECHNOLOGY, and PRACH_ROOT_SEQUENCES.

| SITE_NAME | LATITUDE | LONGITUDE | TECHNOLOGY | PRACH_ROOT_SEQUENCES |
|-----------|----------|-----------|------------|----------------------|
| Site A    | 40.7128  | -74.0060  | LTE        | 0-100                |
| Site B    | 34.0522  | -118.2437 | 5GNR       | 150-200              |
| Site C    | 41.8781  | -87.6298  | LTE        | 50-75                |
| Site D    | 29.7604  | -95.3698  | Both       | 80-120               |

Sample Output File
| Point of Interest Lat | Point of Interest Long | RSI | Closest Location | Technology | Distance (miles) |
|-----------------------|------------------------|-----|------------------|------------|------------------|
| 40.0                  | -75.0                  | 0   | Site A           | LTE        | 4.514981150776129|
| 40.0                  | -75.0                  | 1   | Site A           | LTE        | 4.514981150776129|
| 40.0                  | -75.0                  | 2   | Site A           | LTE        | 4.514981150776129|
...
| 40.0                  | -75.0                  | 97  | Site A           | LTE        | 4.514981150776129|
| 40.0                  | -75.0                  | 98  | Site A           | LTE        | 4.514981150776129|
| 40.0                  | -75.0                  | 99  | Site A           | LTE        | 4.514981150776129|

## Contributing
Contributions are welcome! Please open an issue or submit a pull request with your changes.

