# Data Integrity Check

This script processes Excel files to validate and transform data according to predefined templates. It highlights discrepancies and ensures data conforms to specified standards.

## Features

- Loads templates for validation
- Parses and validates data ranges
- Highlights discrepancies
- Saves results to a new Excel file

## Requirements

- pandas
- numpy
- re

## Usage

1. **Install the required packages**:

    ```bash
    pip install pandas numpy
    ```

2. **Update the file paths**:
    - Edit the script to set `source_path`, `template_path`, and `dest_path` to your actual file paths.

3. **Run the script**:

    ```bash
    python data_integrity_check.py
    ```

## Sample Input and Output

### Sample Input

#### Source Excel (`source_path`):

| Site_ID_Sector | Band Name | Physical Cell ID | Custom: CP_Type | ... |
|----------------|-----------|------------------|-----------------|-----|
| 1234_1         | n66       | 101              | Normal          | ... |
| 1234_2         | n71       | 102              | Normal          | ... |
| 1234_3         | n66       | 999              | NonNormal       | ... |

#### Template Excel (`template_path`):

- **Sheet1**: Contains template data for various fields.
- **Sheet2**: Contains CU_BEDC template data.
- **Sheet3**: Contains gNB, TAC, and K8 ranges.

### Sample Output

#### Output Excel (`dest_path`):

- **Sheet1**: Validated and transformed data with highlighted discrepancies.

| Site_ID_Sector | Band Name | Physical Cell ID          | Custom: CP_Type                                        | ... |
|----------------|-----------|---------------------------|--------------------------------------------------------|-----|
| 1234_1         | n66       | 101                       | Normal                                                 | ... |
| 1234_2         | n71       | 102                       | Normal                                                 | ... |
| 1234_3         | n66       | FALSE - The value is not within the valid range | FALSE - The value should be 'Normal' (current: NonNormal) | ... |

- **PCI_DISCREPANCY**: List of any PCI conflicts found.

| Site_Id_Sector | Band | PCI |
|----------------|------|-----|
| 1234_1         | n66  | 101 |
| 1234_2         | n66  | 101 |

## Explanation

### Source Excel (`source_path`)

Shows a simplified example of the input data with columns like `Site_ID_Sector`, `Band Name`, `Physical Cell ID`, and `Custom: CP_Type`.

### Template Excel (`template_path`)

Indicates the sheets used for templates without detailing their content, as they are assumed to be predefined and known.

### Output Excel (`dest_path`)

Provides examples of validated and transformed data:

- **Sheet1**: Demonstrates how discrepancies are highlighted in the processed data.

- **PCI_DISCREPANCY**: Lists any PCI conflicts detected during the validation process.
