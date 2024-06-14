# RSI Tuner

RSI Tuner is a Python script for analyzing and processing location data based on (RSI) values. It calculates the distance between a problem sector and nearby locations, identifies the closest location for each RSI group, and generates output files for analysis and visualization.

## Features

- Calculate distances between locations using the Haversine formula.
- Identify the closest location to a problem sector for each RSI group.
- Generate KML files for visualization.

## Usage

To use RSI Tuner, follow these steps:

1. **Install Dependencies**:
   - Ensure you have Python 3 installed.
   - Install the required dependencies using `pip install <package_name>`:
     - haversine
     - pandas
     - PySimpleGUI
     - simplekml

2. **Run the Script**:
   - Open a terminal or command prompt.
   - Navigate to the directory containing the script (`rsi_tuner.py`).
   - Run the following command:
     ```bash
     python rsi_tuner.py
     ```

3. **Follow the Instructions**:
   - A Graphical User Interface (GUI) window will appear.
   - Select a CSV file containing location data with 'Site ID', 'RSI', 'Lat', and 'Long' columns using the file browser.
   - Enter the problem sector's RSI value, latitude, and longitude in the corresponding input fields.
   - Click the "Submit" button to start the processing.

## Sample Input

**Input CSV file:**<br>
Site ID, RSI, Lat, Long<br>
1, 75, 40.7128, -74.0060<br>
2, 80, 34.0522, -118.2437<br>
3, 70, 41.8781, -87.6298<br>
4, 75, 29.7604, -95.3698<br>
5, 80, 39.9526, -75.1652

**Problem Sector:**
- RSI: 80
- Latitude: 37.7749
- Longitude: -122.4194

## Sample Output

**Output CSV file (`Output_RSI.csv`):**<br>
RSI, Site ID, Location, Distance<br>
75, 1, (40.7128, -74.0060), 2425.5<br>
70, 3, (41.8781, -87.6298), 1835.6<br>

**KML File (`Output_RSI.kml`):**<br>
- Visualizes the problem sector and the closest locations for each RSI group.

## Dependencies

- Python 3
- haversine
- pandas
- PySimpleGUI
- simplekml
