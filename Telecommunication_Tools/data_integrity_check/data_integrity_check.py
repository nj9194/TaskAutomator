import pandas as pd
import numpy as np
import re


def load_templates(template_path):
    """
    Load the Excel templates.

    Args:
    - template_path (str): Path to the Excel template file.

    Returns:
    - tuple: DataFrames of the templates.
    """
    template = pd.read_excel(template_path)
    cu_bedc_template = pd.read_excel(template_path, sheet_name="Sheet2")
    gnb_tac_template = pd.read_excel(template_path, sheet_name="Sheet3")
    return template.astype(str), cu_bedc_template.astype(str), gnb_tac_template.astype(str)


def parse_gnb_tac_ranges(gnb_tac_template):
    """
    Parse gNB, TAC, and K8 ranges from the template.

    Args:
    - gnb_tac_template (DataFrame): gNB and TAC ranges DataFrame.

    Returns:
    - tuple: Ranges for gNB, TAC, and K8.
    """
    gnb_start, gnb_end = int(gnb_tac_template.iloc[0, 1]), int(gnb_tac_template.iloc[0, 2])
    tac_start, tac_end = int(gnb_tac_template.iloc[1, 1]), int(gnb_tac_template.iloc[1, 2])
    k8_start, k8_end = int(gnb_tac_template.iloc[2, 1]), int(gnb_tac_template.iloc[2, 2])
    return list(range(gnb_start, gnb_end + 1)), list(range(tac_start, tac_end + 1)), list(range(k8_start, k8_end + 1))


def create_lists(template, cu_bedc_template):
    """
    Create lists of values from templates.

    Args:
    - template (DataFrame): Main template DataFrame.
    - cu_bedc_template (DataFrame): CU_BEDC template DataFrame.

    Returns:
    - tuple: Lists of band and CU_BEDC values.
    """
    band_list = ["_".join(row) for row in template.values.tolist()]
    cu_bedc_list = ["_".join(row) for row in cu_bedc_template.values.tolist()]
    return band_list, cu_bedc_list


def highlight_cells(val):
    """
    Highlight cells with discrepancies.

    Args:
    - val (str): Cell value.

    Returns:
    - str: Style to apply.
    """
    if val.startswith("FALSE"):
        return 'background-color: #F67280'
    return ''


def binary_to_decimal(b_sum):
    """
    Convert binary string to decimal.

    Args:
    - b_sum (str): Binary string.

    Returns:
    - int: Decimal value.
    """
    return int(b_sum, 2)


def get_nr_cell_id(row):
    """
    Calculate NR Cell ID from gNodeB ID and Local Cell ID.

    Args:
    - row (Series): DataFrame row.

    Returns:
    - str: NR Cell ID.
    """
    if "FALSE" not in row['Custom: gNodeB_Id'] and "FALSE" not in row['Custom: Local_Cell_Id']:
        gnb_id_bin = format(int(row['Custom: gNodeB_Id']), '024b')
        local_cell_id_bin = format(int(row['Custom: Local_Cell_Id']), '012b')
        return str(binary_to_decimal(gnb_id_bin + local_cell_id_bin))
    return ''


def get_nr_cell_global_id(row):
    """
    Calculate NR Cell Global ID from NR Cell ID.

    Args:
    - row (Series): DataFrame row.

    Returns:
    - str: NR Cell Global ID.
    """
    if "FALSE" not in row['Custom: NR_Cell_Id']:
        nr_cell_id_hex = format(int(row['Custom: NR_Cell_Id']), 'x')
        return str(int("133304" + nr_cell_id_hex, 16))
    return ''


def pci_check(vals, final, dest_path):
    """
    Check for PCI conflicts and save results.

    Args:
    - vals (list): List of (Site_ID_Sector, Band, PCI) tuples.
    - final (DataFrame): Processed DataFrame.
    - dest_path (str): Destination path for the output Excel file.
    """
    conflict_list = []
    valdict = {}

    for site_id_sector, band, pci in vals:
        if site_id_sector in valdict:
            if valdict[site_id_sector] != pci:
                conflict_list.append([site_id_sector, band, valdict[site_id_sector]])
                conflict_list.append([site_id_sector, band, pci])
        else:
            valdict[site_id_sector] = pci

    with pd.ExcelWriter(dest_path) as results:
        if conflict_list:
            invalid_pci = pd.DataFrame(conflict_list, columns=['Site_Id_Sector', 'Band', 'PCI']).drop_duplicates()
            invalid_pci.to_excel(results, sheet_name="PCI_DISCREPANCY", index=False)
        final.style.applymap(highlight_cells).to_excel(results, sheet_name="Sheet1", index=False)


def get_nr_cell_name(row):
    """
    Generate NR Cell Name from row data.

    Args:
    - row (Series): DataFrame row.

    Returns:
    - str: NR Cell Name.
    """
    band_name = row['Band Name'].replace("AWS-4", "AWS4")
    if "DL" in band_name:
        band_name = band_name[:-3]
    return f"{row['Site ID']}_{row['Antenna ID']}_{band_name}"


def get_gnodeb_name(row):
    """
    Generate gNodeB Name from row data.

    Args:
    - row (Series): DataFrame row.

    Returns:
    - str: gNodeB Name.
    """
    return f"{row['Site ID'][:5]}{row['Custom: gNodeB_Id']}"


def get_local_cell_id(row):
    """
    Calculate Local Cell ID from row data.

    Args:
    - row (Series): DataFrame row.

    Returns:
    - str: Local Cell ID.
    """
    assignment_id = 50
    gnb_site_number = row['Custom: gNodeB_Site_Number']
    
    if not gnb_site_number or "FALSE" in gnb_site_number:
        return ''
    
    gnb_site_number = int(gnb_site_number)
    band_name = row['Band Name']
    antenna_id = row['Antenna ID']
    
    band_antenna_map = {
        'n26': {1: 0, 2: 1, 3: 2},
        'n29': {1: 3, 2: 4, 3: 5},
        'n71': {1: 6, 2: 7, 3: 8},
        'n66_AWS': {1: 9, 2: 10, 3: 11},
        'n70': {1: 12, 2: 13, 3: 14},
        'n66': {1: 15, 2: 16, 3: 17}
    }
    
    for key in band_antenna_map:
        if band_name.startswith(key):
            assignment_id = band_antenna_map[key].get(int(antenna_id), assignment_id)
            break
    
    return str((gnb_site_number - 1) * 21 + assignment_id)


def validate_and_transform(final, cu_bedc_list, band_list, pci_list, prach_list, tac_list, k8_list, gnb_list):
    """
    Validate and transform the final DataFrame based on templates and rules.

    Args:
    - final (DataFrame): DataFrame to validate and transform.
    - cu_bedc_list (list): List of valid CU_BEDC values.
    - band_list (list): List of valid band values.
    - pci_list (list): List of valid PCI values.
    - prach_list (list): List of valid PRACH values.
    - tac_list (list): List of valid TAC values.
    - k8_list (list): List of valid K8 values.
    - gnb_list (list): List of valid gNodeB values.

    Returns:
    - DataFrame: Validated and transformed DataFrame.
    """
    patterns = {
        'lat': re.compile(r'^\d{2}\.\d{6}$'),
        'long': re.compile(r'^-\d{2}\.\d{6}$'),
        'col': re.compile(r'^FALSE')
    }

    final['Custom: CP_Type'] = final['Custom: CP_Type'].apply(lambda x: "FALSE - The value should be 'Normal'" if x != "Normal" else x)
    final['Site_ID_CUs_Numbers'] = final['Site_ID_CUs_Numbers'].apply(lambda x: "FALSE - The value does not match with the reference template" if x not in cu_bedc_list else x)
    final['Band_DL_UL_SSB_absfreqA_bandwidth_UL_MIMO'] = final['Band_DL_UL_SSB_absfreqA_bandwidth_UL_MIMO'].apply(lambda x: "FALSE - The value does not match with the reference template" if x not in band_list else x)
    final['Custom: DL_Rank'] = final['Custom: DL_Rank'].apply(lambda x: "FALSE - The value should be '4'" if x != "4" else x)
    final['Custom: UL_Rank'] = final['Custom: UL_Rank'].apply(lambda x: "FALSE - The value should be '2'" if x != "2" else x)
    final['Custom: DL_MIMO'] = final['Custom: DL_MIMO'].apply(lambda x: "FALSE - The value should be '4'" if x != "4" else x)
    final['Custom: UL_MIMO'] = final['Custom: UL_MIMO'].apply(lambda x: "FALSE - The value should be '2'" if x != "2" else x)
    final['Custom: Physical_Cell_ID'] = final['Custom: Physical_Cell_ID'].apply(lambda x: "FALSE - The value is not within the valid range" if x not in pci_list else x)
    final['Custom: PRACH_Config_Index'] = final['Custom: PRACH_Config_Index'].apply(lambda x: "FALSE - The value is not within the valid range" if x not in prach_list else x)
    final['Custom: TAC'] = final['Custom: TAC'].apply(lambda x: "FALSE - The value is not within the valid range" if x not in tac_list else x)
    final['Custom: MMEGI'] = final['Custom: MMEGI'].apply(lambda x: "FALSE - The value should be '10'" if x != "10" else x)
    final['Custom: MME Pool'] = final['Custom: MME Pool'].apply(lambda x: "FALSE - The value should be 'MME_Production'" if x != "MME_Production" else x)
    final['Custom: UpStream_CE'] = final['Custom: UpStream_CE'].apply(lambda x: "FALSE - The value should be 'TRUE'" if x != "TRUE" else x)
    final['Custom: DownStream_CE'] = final['Custom: DownStream_CE'].apply(lambda x: "FALSE - The value should be 'TRUE'" if x != "TRUE" else x)
    final['Custom: gNodeB_Id'] = final['Custom: gNodeB_Id'].apply(lambda x: f"FALSE - The gNB ID should be a unique non-empty value in the range[1-{len(gnb_list)}]. Current value is: {x}" if x not in gnb_list else x)
    final['Custom: NR_Cell_Id'] = final['Custom: NR_Cell_Id'].apply(lambda x: "FALSE - The value should be unique non-empty integer value. Current value is empty" if not x.isdigit() else x)
    final['Custom: Cell_Identity'] = final.apply(get_nr_cell_id, axis=1)
    final['Custom: NR_Cell_Global_Identity'] = final.apply(get_nr_cell_global_id, axis=1)
    final['Cell Name'] = final.apply(get_nr_cell_name, axis=1)
    final['Custom: gNodeB_Name'] = final.apply(get_gnodeb_name, axis=1)
    final['Custom: gNodeB_Site_Number'] = final.apply(get_local_cell_id, axis=1)

    return final


def process_excel_data(source_path, template_path, dest_path):
    """
    Process Excel data by validating and transforming it using templates.

    Args:
    - source_path (str): Path to the source Excel file.
    - template_path (str): Path to the template Excel file.
    - dest_path (str): Path to save the processed Excel file.
    """
    # Load templates
    template, cu_bedc_template, gnb_tac_template = load_templates(template_path)

    # Create value lists from templates
    band_list, cu_bedc_list = create_lists(template, cu_bedc_template)

    # Parse gNB, TAC, and K8 ranges
    gnb_list, tac_list, k8_list = parse_gnb_tac_ranges(gnb_tac_template)
    pci_list = [str(i) for i in range(0, 1008)]
    prach_list = [str(i) for i in range(0, 828)]

    # Load and validate data
    final = pd.read_excel(source_path).astype(str)
    final = validate_and_transform(final, cu_bedc_list, band_list, pci_list, prach_list, tac_list, k8_list, gnb_list)

    # Check for PCI conflicts and save results
    vals = list(zip(final['Site_ID_Sector'], final['Band Name'], final['Physical Cell ID']))
    pci_check(vals, final, dest_path)


if __name__ == "__main__":
    source_path = 'path_to_source_file.xlsx'
    template_path = 'path_to_template_file.xlsx'
    dest_path = 'path_to_output_file.xlsx'
    process_excel_data(source_path, template_path, dest_path)
