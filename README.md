# Excel ERP Merge Automation

## Overview

This project automates the consolidation of ERP results from multiple Excel `.xlsm` project files into a single master output file. It uses a formula-driven **main ERP workbook** to recalculate results for each project file and merges everything into one clean, traceable Excel output.

**Key Features:**

* **Centralized Configuration:** Manage paths via YAML.
* **Input Validation:** Ensures all required sheets and files exist before running.
* **Logging:** Full audit trail for troubleshooting.
* **Optional Executable:** Can be packaged for users without Python installed.

---

## Folder Structure

```text
C:\Users\file
│
├─ 25999 V2.xlsm           # Main ERP workbook (contains formulas)
├─ config.yaml             # Configuration file
├─ erp_merge.py            # Python script
├─ erp_merge.log           # Log file (auto-generated)
├─ Master_ERP_Output.xlsx  # Final merged output (auto-generated)
│
└─ Project files
   ├─ project1.xlsm
   ├─ project2.xlsm
   └─ projectN.xlsm

```

---

## Requirements

### Software

* **Windows OS**
* **Microsoft Excel** (Desktop version required for `xlwings` to trigger recalculations)

### Python

* **Python 3.9+**

### Python Packages

```bash
pip install xlwings pandas openpyxl pyyaml pyinstaller

```

---

## Configuration (`config.yaml`)

All paths and filenames are controlled via a configuration file, allowing for easy updates without modifying the source code.

```yaml
base_path: "C:\\Users\\file"
main_file: "25999 V2.xlsm"
project_folder: "Project files"
output_file: "Master_ERP_Output.xlsx"
log_file: "erp_merge.log"

```

---

## Logging

The script automatically generates `erp_merge.log` to track the automation process.

**Logged information includes:**

* Script start and completion timestamps.
* Specific filenames being processed.
* Validation warnings (e.g., missing sheets).
* Errors and full stack traces for troubleshooting.

---

## Input Validation

Before processing begins, the script performs a "pre-flight" check:

1. **Path Checks:** Verifies the main workbook and project folders exist.
2. **File Count:** Ensures at least one `.xlsm` file is present in the project folder.
3. **Schema Checks:** * **Project files** must contain a `Data` worksheet.
* **Main workbook** must contain both `Data` and `ERP` worksheets.



> [!IMPORTANT]
> If any validation step fails, the script will terminate immediately and log the specific reason to prevent corrupted data output.

