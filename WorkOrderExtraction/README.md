# Work Order Extraction Framework

This framework is designed to automatically parse and extract details from PDF Work Orders and convert them into structured Excel (`.xlsx`) files.

## Prerequisites

To use this framework, make sure you have the provided Python virtual environment activated or the required packages installed. 

The framework requires the following libraries:
- `pdfplumber`
- `openpyxl`

If you are using the virtual environment in this directory, the dependencies are already installed.

## Usage

You can run the script via the command line and point it to any PDF work order you want to parse.

### 1. Activate the Virtual Environment
Before running the script, activate the local environment so Python knows where to find the libraries:

```bash
# On Mac/Linux:
source venv/bin/activate
```

### 2. Run the Script
Use the `parse_work_order.py` script by passing the path of the PDF you want to parse:

```bash
python parse_work_order.py "path/to/your/work_order.pdf" --output "Desired_Output_Name.xlsx"
```

**Options:**
- `pdf_path`: (Required) The path to the input Work Order PDF file.
- `--output` or `-o`: (Optional) The name of the output Excel file. If you don't specify this, it defaults to `WorkOrder_Summary.xlsx`.

### Examples

**Default Usage:**
If you run the script without any arguments, it defaults to checking for `DIGITCOM_630344375.pdf` and outputs `WorkOrder_Summary.xlsx`:
```bash
python parse_work_order.py
```

**Parsing a New File:**
```bash
python parse_work_order.py my_new_work_order.pdf -o my_new_summary.xlsx
```

## How It Works

When you run the script, it automatically:
1. **Extracts Work Order Details:** Creates a `Work Order Details` sheet containing general information (e.g., Vendor, WO Number, Date, WO Period, Base Value, Taxes).
2. **Extracts Line Items:** Creates a `Parsed Work Order` sheet extracting every individual item per site.
3. **Formats the Output:** 
   - Merges identical rows visually (Site Index, Site ID, and Base Value) so duplicate site data is grouped nicely.
   - Highlights the Total Amount column.
   - Automatically auto-fits the column widths.
