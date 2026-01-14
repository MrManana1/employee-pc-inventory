# Employee PC Inventory Extractor

Extracts computer hardware information from HWiNFO HTML reports and generates a clean Excel inventory sheet for office/enterprise use. The script processes HTML files containing hardware information and compiles data such as PC model, serial number, CPU, RAM, monitor details, and storage into a structured Excel spreadsheet.

## Features

- **Automated Extraction**: Parses HWiNFO-generated HTML reports to extract key hardware information
- **Department Organization**: Automatically detects department from folder structure
- **Comprehensive Data**: Extracts computer model, serial number, CPU model, RAM, monitor details (name and serial), and storage information
- **Excel Output**: Generates a formatted Excel file with serial numbers and organized columns
- **Error Handling**: Gracefully handles missing data and file processing errors

## Requirements

- Python 3.6 or higher
- pandas library
- pathlib (included in Python 3.4+)

## Installation

1. Ensure Python 3.6+ is installed on your system
2. Install the required pandas library:
   ```bash
   pip install pandas
   ```

## Usage

1. Place your HWiNFO HTML report files (.htm or .html) in the appropriate department folders within the script's directory
2. Run the script:
   ```bash
   python main.py
   ```
3. The script will process all HTML files and generate `Office_Computer_Inventory.xlsx`

## File Structure

Organize your HWiNFO reports in department folders like this:
```
employee-pc-inventory/
├── main.py
├── AR/
│   ├── Employee1.HTM
│   └── Employee2.HTM
├── Billing Production/
│   └── Employee3.HTM
└── ...
```

## Output

The generated Excel file (`Office_Computer_Inventory.xlsx`) contains the following columns:
- SR#: Sequential record number
- Assigned to: Employee name (derived from filename)
- Department: Department name (derived from folder name)
- Computer System Model: PC model (e.g., DELL OptiPlex 7040)
- Serial Number: System serial number
- CPU Model: Processor information
- Memory (RAM): RAM size and type (e.g., "16 GB DDR4")
- Monitor: Monitor model name
- Monitor SN: Monitor serial number
- Storage: Storage type and capacity (e.g., "SSD (500 GB)")

## Troubleshooting

- **Excel file is open**: Close Excel completely before running the script, as it needs to delete the old file
- **Permission errors**: Ensure you have write permissions in the directory
- **Missing data**: Some fields may show "Not found" if the information isn't available in the HTML report
- **Encoding issues**: The script handles UTF-8 encoding and ignores errors for compatibility

## Notes

- The script automatically deletes the old Excel file before creating a new one
- Employee names are derived from the HTML filenames (without extension)
- Department names are taken from the parent folder names
- Monitor information extraction looks for specific patterns in the HTML content
- RAM extraction includes type detection (DDR3/DDR4/DDR5) when available
