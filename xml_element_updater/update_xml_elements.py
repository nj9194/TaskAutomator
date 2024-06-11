import os
from lxml import etree

def update_xml_elements(input_folder, output_folder, element_name, new_text):
    """
    Update specified XML elements in all XML files within the input folder
    and save the modified files to the output folder.

    Args:
        input_folder (str): Path to the folder containing XML files.
        output_folder (str): Path to the folder to save modified XML files.
        element_name (str): The name of the XML element to update.
        new_text (str): The new text to set for the specified XML elements.
    """
    # Iterate through XML files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith('.xml'):
            input_file_path = os.path.join(input_folder, filename)
            output_file_path = os.path.join(output_folder, filename)

            # Read XML data from input file
            with open(input_file_path, 'r') as file:
                xml_data = file.read()

            # Parse XML data
            root = etree.fromstring(xml_data)

            # Iterate through all elements in the XML tree
            for element in root.iter():
                # Check if the element's local name matches the desired element name
                if etree.QName(element).localname == element_name:
                    # Update the text of the matching elements if it's not already set to the new text
                    if element.text != new_text:
                        element.text = new_text

                        # Write the updated XML data to the output file
                        with open(output_file_path, 'w') as output_file:
                            output_file.write(etree.tostring(root, encoding='utf-8').decode('utf-8'))
                            print(f"Updated element '{element_name}' in file: {filename}")

# Define input and output folder paths
input_folder_path = r'C:\Users\niyati.joshi\Documents\post-gsp\Dest_Folder'
output_folder_path = r'C:\Users\niyati.joshi\Documents\combined_1\Output3'

# Define the element name and new text
element_name = "defautPagCycle"
new_text = "defaultPagCycle_rf128"

# Update XML elements
update_xml_elements(input_folder_path, output_folder_path, element_name, new_text)
