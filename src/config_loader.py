from pathlib import Path
import pandas as pd

CONFIG_PATH = Path("config/config.xlsx")

DEFAULTS = [
    ("input_directory", "input"),
    ("sheet_name", "Remittance Report"),
    ("output_file", "output/valid_remittance.xlsx"),
    ("template_file", "Output_Template/OutputFile_Temp2.xlsx"),
    ("master_table_file", "MasterDataTable/MasterDatatable_1.xlsx"),
    ("log_file", "logs/app.log"),
    ("log_level", "INFO"),
]

# Order of the numeric columns on each debtor row in the Remittance Report
REPORT_COLUMNS = [
    "Payments Received", "Direct Payments", "Principal", "Interest", "Fees",
    "Costs", "Other", "Costs Expended", "Client Proceeds", "Fees Due",
]


def _create_default_config():
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(DEFAULTS, columns=["key", "value"]).to_excel(CONFIG_PATH, index=False)


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        _create_default_config()
    df = pd.read_excel(CONFIG_PATH, dtype=str)
    config = dict(zip(df["key"].str.strip(), df["value"].str.strip()))
    for key, value in DEFAULTS:          # fill anything missing
        config.setdefault(key, value)
    config["report_columns"] = REPORT_COLUMNS
    return config