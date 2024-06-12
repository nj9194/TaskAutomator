import logging
import traceback
import csv
from haversine import haversine, Unit
import pandas as pd
import PySimpleGUI as sg
import simplekml

# Set up logging
logging.basicConfig(filename='RSI_Tuner.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to log messages
def log_message(message):
    """
    Log messages to a file.

    Parameters:
        message (str): Message to be logged.
    """
    logging.info(message)

# Function to calculate distance between two points using Haversine formula
def calculate_distance(dest, src):
    """
    Calculate the distance between two points using the Haversine formula.

    Parameters:
        dest (tuple): Destination coordinates (latitude, longitude).
        src (tuple): Source coordinates (latitude, longitude).

    Returns:
        float: Distance between the two points in miles.
    """
    return haversine(src, dest, unit=Unit.MILES)

# Function to find the minimum distance and associated data for a given RSI
def find_minimum_distance(k, data_list):
    """
    Find the minimum distance and associated data for a given RSI.

    Parameters:
        k (int): RSI value.
        data_list (list): List of data points containing site ID, location, and distance.

    Returns:
        list: Data with minimum distance for the given RSI.
    """
    updated_list = []
    for item in data_list:
        item.insert(0, k)
        updated_list.append(item)
        updated_list.sort(key=lambda i: i[3])

    return updated_list[0]

# Function to find the maximum distance from a list of data points
def find_maximum_distance(input_list, output_folder):
    """
    Find the maximum distance from a list of data points and save the result to a CSV file.

    Parameters:
        input_list (list): List of data points.
        output_folder (str): Path to the output folder.
    """
    output_df = pd.DataFrame(input_list, columns=["RSI", "Site_ID", "Location", "Distance"])
    sorted_df = output_df.sort_values("Distance", ascending=False)
    sorted_df.to_csv(f'{output_folder}\\Output_RSI.csv', index=False)

# Function to group data points based on RSI values and find minimum distances
def group_and_process_data(location_file, problem_rsi, problem_location, output_folder):
    """
    Group data points based on RSI values and find minimum distances for each RSI.

    Parameters:
        location_file (str): Path to the CSV file containing location data.
        problem_rsi (int): RSI value of the problem sector.
        problem_location (tuple): Coordinates (latitude, longitude) of the problem sector.
        output_folder (str): Path to the output folder.
    """
    final_list = []
    location_df = pd.read_csv(location_file)
    location_df.columns = [c.replace(" ", "") for c in location_df.columns]
    location_df["lat_long"] = list(zip(location_df.Lat, location_df.Long))
    location_df = location_df.drop(columns=["Lat", "Long"])
    location_df["distance"] = location_df["lat_long"].apply(calculate_distance, src=problem_location)
    location_data = location_df.to_dict("records")
    rsi_data = {}
    for record in location_data:
        if record["distance"] == 0.0 or record["RSI"] == problem_rsi:
            continue
        else:
            key = record["RSI"]
            value1 = record["Site_ID"]
            value2 = record["lat_long"]
            value3 = record["distance"]
            temp_list = [[value1, value2, value3]]
            if key not in rsi_data.keys():
                rsi_data[key] = temp_list
            else:
                rsi_data[key].extend(temp_list)

    for rsi, group in rsi_data.items():
        final_list.append(find_minimum_distance(rsi, group))

    find_maximum_distance(final_list, output_folder)

    # Generate KML file
    kml = simplekml.Kml()
    problem_sector_name = "Problem Sector"
    problem_sector_lat, problem_sector_long = problem_location
    problem_sector_coords = [(problem_sector_long, problem_sector_lat)]
    kml.newpoint(name=problem_sector_name, coords=problem_sector_coords, description=f"RSI: {problem_rsi}")
    
    with open(f'{output_folder}\\Output_RSI.csv', "r") as fh:
        file = csv.reader(fh)
        for row in file:
            coordinates = [tuple(map(float, coord.strip().split(','))) for coord in row[2][1:-1].split()]
            kml.newpoint(name=f"{row[1]} ({row[0]})", coords=coordinates, description=f"RSI: {row[0]}")
            kml.newlinestring(name=f"Distance: {row[3]}", coords=[problem_sector_coords[0], coordinates[0]])

    kml.save(f'{output_folder}\\Output_RSI.kml')

# GUI layout
layout = [
    [sg.Text("Choose a CSV file with 'Site ID', 'RSI', 'Lat', 'Long' columns: "), sg.Input(key="-FILE-"), sg.FileBrowse(key="-IN-")],
    [sg.Text("Choose a destination folder: "), sg.Input(key="-FOLDER-"), sg.FolderBrowse(key="-OUT-")],
    [sg.Text('Problem sector name:'), sg.InputText(key="problem_sector")],
    [sg.Text('Problem sector RSI:'), sg.InputText(key="problem_rsi")],
    [sg.Text('Problem sector Lat:'), sg.InputText(key="problem_lat")],
    [sg.Text('Problem sector Long:'), sg.InputText(key="problem_long")],
    [sg.Button("Submit")]
]

# Create the GUI window
window = sg.Window('RSI Tuner', layout, size=(800, 400))

try:
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Exit":
            break
        elif event == "Submit":
            location_file = values["-FILE-"]
            output_folder = values["-FOLDER-"]
            problem_rsi = values["problem_rsi"]
            problem_lat = values["problem_lat"]
            problem_long = values["problem_long"]
            problem_sector = values["problem_sector"]

            if not location_file or not problem_sector or not problem_rsi or not problem_lat or not problem_long or not output_folder:
                sg.popup_error("Please fill in all the required fields.")
                continue

            problem_rsi = int(problem_rsi)
            problem_lat = float(problem_lat)
            problem_long = float(problem_long)
            problem_location = (problem_lat, problem_long)

            group_and_process_data(location_file, problem_rsi, problem_location, output_folder)

            sg.popup(f'File has been created! Please check your destination folder:\n{output_folder}', title="Success!")
            window.close()

except Exception as ex:
    traceback_info = traceback.format_exc()
    sg.popup(f'An error occurred. Here is the info:', ex, title="Error!")
    log_message(f'Exception: {ex}\n{traceback_info}')
    exit()
