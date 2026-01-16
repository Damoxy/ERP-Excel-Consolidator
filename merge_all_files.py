import xlwings as xw
import pandas as pd
import os
import glob
import datetime

# === CONFIGURATION ===
base_path = r"C:\Users\file"
main_file = os.path.join(base_path, "25999 V2.xlsm")
project_folder = os.path.join(base_path, "Project files")
output_file = os.path.join(base_path, "Master_ERP_Output.xlsx")

# Automatically detect all .xlsm files in the project folder
data_files = glob.glob(os.path.join(project_folder, "*.xlsm"))
if not data_files:
    raise FileNotFoundError(f"No .xlsm files found in {project_folder}")

merged_erp = []

# Use context manager to safely open Excel
with xw.App(visible=False) as app:
    # Open main workbook once
    wb_main = app.books.open(main_file)

    for f in data_files:
        print(f"Processing: {os.path.basename(f)}")

        # Read Data sheet from source file
        df_data = pd.read_excel(f, sheet_name="Data")

        # Convert datetime.time columns to string (Pandas 2.x safe)
        for col in df_data.columns:
            if df_data[col].apply(lambda x: isinstance(x, datetime.time)).any():
                df_data[col] = df_data[col].apply(lambda x: x.strftime("%H:%M:%S") if isinstance(x, datetime.time) else x)

        # Refresh sheet reference to avoid COM disconnect
        ws_data = wb_main.sheets["Data"]
        ws_erp = wb_main.sheets["ERP"]

        # Ensure sheet is visible and unprotected
        ws_data.visible = True
        try:
            ws_data.api.Unprotect()
        except:
            pass  # ignore if not protected

        # Clear sheet and unmerge any merged cells
        ws_data.clear()
        try:
            merged_cells = ws_data.api.UsedRange.MergeCells
            if merged_cells:
                ws_data.api.UnMerge()
        except:
            pass  # ignore if no merged cells

        # Write headers + data
        ws_data.range("A1").value = [df_data.columns.tolist()] + df_data.values.tolist()

        # Recalculate all formulas in ERP sheet
        wb_main.app.calculate()

        # Read updated ERP sheet
        df_erp = ws_erp.range("A1").expand().options(pd.DataFrame, header=1, index=False).value

        # Add Source_File column
        df_erp.insert(0, "Source_File", os.path.basename(f))

        merged_erp.append(df_erp)

    # Merge all ERP results once and save
    df_final = pd.concat(merged_erp, ignore_index=True)
    df_final.to_excel(output_file, index=False)

print(f"✅ Done! Master ERP results saved to: {output_file}")
