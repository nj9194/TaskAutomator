import pandas as pd
from math import radians, sin, cos, sqrt, atan2
from tkinter import Tk, Label, Entry, Button, filedialog, StringVar, OptionMenu, ttk
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def calculate_distance(lat1, lon1, lat2, lon2):
    """Calculate the distance between two points on the Earth using the Haversine formula."""
    R = 3958.8  # Radius of the Earth in miles

    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])

    dlon = lon2 - lon1
    dlat = lat2 - lat1

    a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))

    distance = R * c
    return distance

def check_prach_sequence(sequence, value):
    """Check if a given PRACH_ROOT_SEQUENCES contains a specific value."""
    sequence_parts = str(sequence).split('-')
    start, end = 0, 0
    if len(sequence_parts) == 1 and sequence_parts[0].isdigit():
        start = int(sequence_parts[0])
        end = start + 1
    elif len(sequence_parts) == 2 and sequence_parts[0].isdigit() and sequence_parts[1].isdigit():
        start = int(sequence_parts[0])
        end = int(sequence_parts[1]) + 1
    return value in range(start, end)

def find_closest_locations(df, poi_lat, poi_lon, selected_technology, num_rsi):
    """Find the closest locations for each PRACH Root Sequences Index."""
    results_data = []

    for p1_value in range(num_rsi):
        if selected_technology != "Both":
            filtered_locations = df[df['TECHNOLOGY'] == selected_technology]
        else:
            filtered_locations = df

        filtered_locations = filtered_locations[
            filtered_locations['PRACH_ROOT_SEQUENCES'].apply(lambda x: check_prach_sequence(x, p1_value))
        ]

        closest_location = None
        closest_distance = float('inf')
        for _, loc in filtered_locations.iterrows():
            distance = calculate_distance(poi_lat, poi_lon, loc["LATITUDE"], loc["LONGITUDE"])
            if distance < closest_distance:
                closest_location = loc
                closest_distance = distance

        results_row = {
            "Point of Interest Lat": poi_lat,
            "Point of Interest Long": poi_lon,
            "RSI": p1_value,
            "Closest Location": closest_location["SITE_NAME"] if closest_location is not None else "No location found",
            "Technology": closest_location["TECHNOLOGY"] if closest_location is not None else "N/A",
            "Distance (miles)": closest_distance if closest_location is not None else "N/A"
        }
        results_data.append(results_row)

    return pd.DataFrame(results_data)

def read_excel_file(file_path):
    """Read an Excel file and return a DataFrame."""
    return pd.read_excel(file_path)

def save_to_excel(df, output_folder, selected_technology):
    """Save the DataFrame to an Excel file with a timestamp."""
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")
    output_file = f"{output_folder}/closest_locations_with_RSIs_output_{selected_technology}_{current_time}.xlsx"
    df.to_excel(output_file, index=False)
    return output_file

def select_output_folder(entry_output_folder):
    """Handle button click event for selecting output folder."""
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    if output_folder:
        entry_output_folder.delete(0, 'end')
        entry_output_folder.insert(0, output_folder)

def select_input_file(entry_input_file):
    """Handle button click event for selecting input file."""
    input_file = filedialog.askopenfilename(title="Select Input File", filetypes=[("Excel files", "*.xlsx")])
    if input_file:
        entry_input_file.delete(0, 'end')
        entry_input_file.insert(0, input_file)

def calculate_closest_locations(entry_lat, entry_lon, entry_input_file, entry_output_folder, technology_var, num_rsi_entry, progress_bar, result_label, root):
    """Handle button click event to perform calculations."""
    try:
        poi_lat = float(entry_lat.get())
        poi_lon = float(entry_lon.get())
        input_path = entry_input_file.get()
        num_rsi = int(num_rsi_entry.get())

        df = read_excel_file(input_path)
        selected_technology = technology_var.get()

        progress_bar["maximum"] = num_rsi
        progress_bar["value"] = 0
        progress_bar.start()

        results = find_closest_locations(df, poi_lat, poi_lon, selected_technology, num_rsi)

        progress_bar.stop()

        output_folder = entry_output_folder.get()
        output_file = save_to_excel(results, output_folder, selected_technology)
        result_label.config(text=f"Output saved to {output_file}")
        logging.info(f"Output saved to {output_file}")
    except ValueError as e:
        result_label.config(text=f"Please enter valid input values. Error: {e}")
        logging.error(f"ValueError: {e}")
    except Exception as e:
        result_label.config(text=f"An error occurred: {e}")
        logging.error(f"Exception: {e}")

def create_gui():
    """Create and run the Tkinter GUI."""
    root = Tk()
    root.title("Closest Locations Finder")

    label_lat = Label(root, text="Enter Latitude:")
    label_lon = Label(root, text="Enter Longitude:")
    entry_lat = Entry(root)
    entry_lon = Entry(root)
    label_input_file = Label(root, text="Input File:")
    entry_input_file = Entry(root)
    button_input = Button(root, text="Select Input File", command=lambda: select_input_file(entry_input_file))
    label_output_folder = Label(root, text="Output Folder:")
    entry_output_folder = Entry(root)
    button_output = Button(root, text="Select Output Folder", command=lambda: select_output_folder(entry_output_folder))
    label_num_rsi = Label(root, text="Number of PRACH Root Sequences (default 891):")
    num_rsi_entry = Entry(root)
    num_rsi_entry.insert(0, "891")
    button_process = Button(root, text="Process", command=lambda: calculate_closest_locations(entry_lat, entry_lon, entry_input_file, entry_output_folder, technology_var, num_rsi_entry, progress_bar, result_label, root))
    result_label = Label(root, text="")
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")

    technology_label = Label(root, text="Select Technology:")
    technology_options = ["LTE", "5GNR", "Both"]
    technology_var = StringVar(root)
    technology_var.set("LTE")
    technology_menu = OptionMenu(root, technology_var, *technology_options)

    label_lat.grid(row=0, column=0, sticky="w", padx=5, pady=5)
    entry_lat.grid(row=0, column=1, padx=5, pady=5)
    label_lon.grid(row=1, column=0, sticky="w", padx=5, pady=5)
    entry_lon.grid(row=1, column=1, padx=5, pady=5)
    label_input_file.grid(row=2, column=0, sticky="w", padx=5, pady=5)
    entry_input_file.grid(row=2, column=1, padx=5, pady=5)
    button_input.grid(row=2, column=2, padx=5, pady=5)
    label_output_folder.grid(row=3, column=0, sticky="w", padx=5, pady=5)
    entry_output_folder.grid(row=3, column=1, padx=5, pady=5)
    button_output.grid(row=3, column=2, padx=5, pady=5)
    label_num_rsi.grid(row=4, column=0, sticky="w", padx=5, pady=5)
    num_rsi_entry.grid(row=4, column=1, padx=5, pady=5)
    technology_label.grid(row=5, column=0, sticky="w", padx=5, pady=5)
    technology_menu.grid(row=5, column=1, columnspan=2, padx=5, pady=5)
    button_process.grid(row=6, column=0, columnspan=2, padx=5, pady=5)
    result_label.grid(row=7, column=0, columnspan=3, padx=5, pady=5)
    progress_bar.grid(row=8, column=0, columnspan=3, padx=5, pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
