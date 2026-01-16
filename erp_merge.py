import xlwings as xw
import pandas as pd
import os
import glob
import datetime
import logging
import yaml
import sys

# =========================
# Load configuration
# =========================
try:
    with open("config.yaml", "r") as f:
        config = yaml.safe_load(f)
except Exception as e:
    print("Failed to load config.yaml")
    sys.exit(1)

base_path = config["base_path"]
main_file = os.path.join(base_path, config["main_file"])
project_folder = os.path.join(base_path, config["project_folder"])
output_file = os.path.join(base_path, config["output_file"])
log_file = os.path.join(base_path, config["log_file"])

# =========================
# Logging setup
# =========================
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logging.info("=== ERP Merge Process Started ===")

# =========================
# Validation
# =========================
def validate_environment():
    if not os.path.exists(main_file):
        logging.error(f"Main ERP file not found: {main_file}")
        sys.exit(1)

    if not os.path.exists(project_folder):
        logging.error(f"Project folder not found: {project_folder}")
        sys.exit(1)

    files = glob.glob(os.path.join(project_folder, "*.xlsm"))
    if not files:
        logging.error("No .xlsm files found in project folder")
        sys.exit(1)

    return files

data_files = validate_environment()
merged_erp = []

# =========================
# Excel Processing
# =========================
try:
    with xw.App(visible=False) as app:
        wb_main = app.books.open(main_file)

        if "Data" not in [s.name for s in wb_main.sheets] or "ERP" not in [s.name for s in wb_main.sheets]:
            logging.error("Main workbook missing Data or ERP sheet")
            sys.exit(1)

        ws_data = wb_main.sheets["Data"]
        ws_erp = wb_main.sheets["ERP"]

        for f in data_files:
            logging.info(f"Processing file: {os.path.basename(f)}")

            try:
                df_data = pd.read_excel(f, sheet_name="Data")
            except Exception:
                logging.error(f"Missing Data sheet in {f}")
                continue

            # Convert datetime.time columns safely
            for col in df_data.columns:
                if df_data[col].apply(lambda x: isinstance(x, datetime.time)).any():
                    df_data[col] = df_data[col].apply(
                        lambda x: x.strftime("%H:%M:%S") if isinstance(x, datetime.time) else x
                    )

            # Prepare Data sheet
            ws_data.visible = True
            try:
                ws_data.api.Unprotect()
            except Exception:
                pass

            ws_data.clear()
            try:
                if ws_data.api.UsedRange.MergeCells:
                    ws_data.api.UnMerge()
            except Exception:
                pass

            # Write data
            ws_data.range("A1").value = [df_data.columns.tolist()] + df_data.values.tolist()

            # Recalculate ERP
            wb_main.app.calculate()

            # Read ERP output
            df_erp = ws_erp.range("A1").expand().options(
                pd.DataFrame, header=1, index=False
            ).value

            df_erp.insert(0, "Source_File", os.path.basename(f))
            merged_erp.append(df_erp)

        if not merged_erp:
            logging.error("No ERP results were generated")
            sys.exit(1)

        df_final = pd.concat(merged_erp, ignore_index=True)
        df_final.to_excel(output_file, index=False)

        logging.info(f"Master ERP output saved to {output_file}")

except Exception as e:
    logging.exception("Fatal error during ERP merge process")
    sys.exit(1)

logging.info("=== ERP Merge Process Completed Successfully ===")
print(f"Done! Master ERP results saved to: {output_file}")
