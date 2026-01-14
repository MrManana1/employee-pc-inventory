import os
import re
import pandas as pd
from pathlib import Path

# ===================== CONFIGURATION =====================
MAIN_FOLDER = os.getcwd()
OUTPUT_EXCEL = "Office_Computer_Inventory.xlsx"

# Delete old Excel file if it exists
if os.path.exists(OUTPUT_EXCEL):
    try:
        os.remove(OUTPUT_EXCEL)
        print(f"→ Old file '{OUTPUT_EXCEL}' deleted successfully.")
    except PermissionError:
        print("\n" + "="*70)
        print("ERROR: Cannot delete Excel file - it is open in Excel!")
        print("Please CLOSE Excel completely (check Task Manager).")
        print("Then run the script again.")
        print("="*70 + "\n")
        input("Press Enter to exit...")
        exit(1)
    except Exception as e:
        print(f"Warning: Could not delete old file: {e}")

# ===================== EXTRACTION FUNCTIONS =====================

def extract_between(text, start_marker, end_marker):
    try:
        start = text.index(start_marker) + len(start_marker)
        end = text.index(end_marker, start)
        return text[start:end].strip()
    except ValueError:
        return "Not found"


def clean_text(text):
    text = text.replace("\xa0", " ").replace("&nbsp;", " ")
    return " ".join(text.split())


def get_computer_model(text):
    if "DELL OptiPlex 7040" in text: return "DELL OptiPlex 7040"
    if "DELL OptiPlex 3050" in text: return "DELL OptiPlex 3050"
    if "DELL OptiPlex 7010" in text: return "DELL OptiPlex 7010"
    if "HP Compaq Elite 8300 SFF" in text: return "HP Compaq Elite 8300 SFF"
    
    match = re.search(r'Computer Brand Name:<TD[^>]*>([^<]+)', text)
    if match:
        model = match.group(1).strip()
        return clean_text(model)
    return "Not found"


def get_serial_number(text):
    match = re.search(r'Product\s+Serial\s+Number:\s*([A-Za-z0-9\-_]+)', text, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    match = re.search(r'(?:System\s+Serial\s+Number|Chassis\s+Serial\s+Number):\s*([A-Za-z0-9\-_]+)', text, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    match = re.search(r'Mainboard\s+Serial\s+Number:\s*([A-Za-z0-9\-_]+)', text, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    match = re.search(r'([A-Z]{3}[0-9]{5}[A-Z0-9]{0,2})', text)
    if match:
        return match.group(1).strip()

    return "Not found"


def get_cpu_model(text):
    match = re.search(r'Processor\s+Name:\s*([^<]+?)(?:\s*</TD>|CPU\s*@|\s*$)', text, re.IGNORECASE)
    if match:
        cpu = match.group(1).strip()
        cpu = re.sub(r'\s*@\s*\d+\.\d+GHz.*', '', cpu, flags=re.IGNORECASE)
        cpu = re.sub(r'\s*\(R\)|\s*\(TM\)', '', cpu)
        return clean_text(cpu)

    match = re.search(r'(Intel\s+Core\s+i[3579]-\d{4,5}[A-Z]?)', text, re.IGNORECASE)
    if match:
        return clean_text(match.group(1).strip())

    return "Not found"


def get_ram(text):
    total_match = re.search(r'Total Memory Size:<TD>(\d+\.?\d*)\s*(GBytes|GB|MB|MBytes)', text, re.IGNORECASE)
    if total_match:
        size = total_match.group(1).strip()
        unit = total_match.group(2).strip().replace('GBytes', 'GB')

        if unit.lower() in ['mb', 'mbytes']:
            try:
                size_gb = float(size) / 1024
                size = str(int(size_gb)) if size_gb.is_integer() else f"{size_gb:.1f}"
                unit = "GB"
            except:
                pass

        ram_type = ""
        text_upper = text.upper()
        if "DDR4" in text_upper: ram_type = "DDR4"
        elif "DDR3" in text_upper: ram_type = "DDR3"
        elif "DDR5" in text_upper: ram_type = "DDR5"

        result = f"{size} {unit}"
        if ram_type:
            result += f" {ram_type}"
        return result.strip()

    return "Not found"


def get_monitor(text):
    """
    First finds the 'Monitor' section, then extracts:
    - Monitor Name / Monitor Name (Manuf)
    - Serial Number
    """
    monitor_name = "Not found"
    serial = "Not found"

    # Step 1: Find the Monitor section
    monitor_section = ""
    monitor_start_match = re.search(r'id="Monitor"', text)
    if monitor_start_match:
        monitor_start = monitor_start_match.start()
        # Take everything from "Monitor" until next major section (e.g. Drives, Audio, etc.)
        next_section = re.search(r'(?i)(Drives|Audio|Network|Ports|Bus|Video|CPU|Motherboard)', text[monitor_start:])
        if next_section:
            end_pos = monitor_start + next_section.start()
            monitor_section = text[monitor_start:end_pos]
        else:
            monitor_section = text[monitor_start:]

    if monitor_section == "":
        return "Not found"

    # Step 2: Extract Monitor Name from the section
    name_patterns = [
        r'Monitor\s+Name(?:\s*\(Manuf\))?:<TD[^>]*>([^<\n\r]+?)(?:\s*</TD>|\s*$)',
        r'Monitor\s+Name\s*\(Manuf\):<TD[^>]*>([^<\n\r]+)',
        r'(?:Hewlett-Packard|HP|Lenovo|Dell|Acer|LG|Samsung)[^\n\r<]+'
    ]

    for pattern in name_patterns:
        match = re.search(pattern, monitor_section, re.IGNORECASE | re.MULTILINE)
        if match:
            monitor_name = clean_text(match.group(1).strip())
            break

    # Step 3: Extract Serial Number from the section
    serial_match = re.search(r'Serial\s+Number(?:\s*\(Manuf\))?:<TD[^>]*>([A-Za-z0-9]+)', monitor_section, re.IGNORECASE)
    if serial_match:
        serial = serial_match.group(1).strip()

    # Step 4: Return the results
    return monitor_name, serial


def get_storage(text):
    # Determine type
    if "NVMe Drives" in text:
        storage_type = "NVMe"
    elif "SSD Drive (Non-rotating)" in text:
        storage_type = "SSD"
    else:
        storage_type = "HDD"
    
    # Extract capacity
    match = re.search(r'Drive Capacity:<TD[^>]*>([^<]+)', text)
    if match:
        full_capacity = match.group(1).strip()
        gb_match = re.search(r'\(([^)]+)\)', full_capacity)
        if gb_match:
            capacity = gb_match.group(1).strip()
            return f"{storage_type} ({capacity})"
    return "Not found"


# ===================== PROCESS SINGLE FILE =====================

def process_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

        filename = Path(filepath).stem
        employee_name = filename.replace("_", " ").strip()

        department = Path(filepath).parent.name

        monitor_name, monitor_sn = get_monitor(content)

        data = {
            "Assigned to": employee_name,
            "Department": department,
            "Computer System Model": get_computer_model(content),
            "Serial Number": get_serial_number(content),
            "CPU Model": get_cpu_model(content),
            "Memory (RAM)": get_ram(content),
            "Monitor": monitor_name,
            "Monitor SN": monitor_sn,
            "Storage": get_storage(content)
        }
        return data

    except Exception as e:
        print(f"Error processing {filepath.name}: {e}")
        return None


# ===================== MAIN PROCESSING =====================

if __name__ == "__main__":
    results = []

    html_files = []
    for root, dirs, files in os.walk(MAIN_FOLDER):
        for file in files:
            if file.lower().endswith(('.htm', '.html')):
                html_files.append(Path(root) / file)

    print(f"Found {len(html_files)} HWiNFO report files...\n")

    for i, file_path in enumerate(html_files, 1):
        print(f"[{i:3d}/{len(html_files):3d}] Processing: {file_path.name}")
        data = process_file(file_path)
        if data:
            results.append(data)

    if not results:
        print("\nNo valid data extracted from any file.")
    else:
        df = pd.DataFrame(results)
        df.insert(0, 'SR#', range(1, len(df) + 1))

        columns_order = [
            'SR#', 'Assigned to', 'Department', 'Computer System Model',
            'Serial Number', 'CPU Model', 'Memory (RAM)', 'Monitor', 'Monitor SN', 'Storage'
        ]

        df = df[[col for col in columns_order if col in df.columns]]

        try:
            df.to_excel(OUTPUT_EXCEL, index=False)
            print(f"\nSuccess! New Excel file created: {OUTPUT_EXCEL}")
            print(f"Total records: {len(df)}")
        except PermissionError:
            print("\n" + "="*70)
            print("ERROR: Cannot write to Excel - file is open in Excel!")
            print("Close Excel and try again.")
            print("="*70 + "\n")
        except Exception as e:
            print("Error saving Excel:", e)