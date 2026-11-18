import random
import re
import logging as log
import pandas as pd
import numpy as np
import unicodedata
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

log.basicConfig(level=log.INFO, format="%(levelname)s: %(message)s")

RAW_DIR = Path.cwd().joinpath("NBA_rosters")
OUTPUT_DIR = Path.cwd().joinpath("Output_files")
OUTPUT_DIR.mkdir(exist_ok=True)
MasterFile = OUTPUT_DIR.joinpath("NBA_Roster_20262027.xlsx")

suffixes = {"jr.", "jr", "sr.", "sr", "ii", "iii", "iv", "(tw)"}

def format_email_column(row: pd.Series) -> str| None:
    #Formats player's emails to only lower case letters
    if row is None:
        return None
    parts = [re.sub(r"[^a-z]", "", p.lower()) for p in row["clean_parts"]]
    team = str(row["Team"]).lower().replace(" ", "")

    username = "_".join(parts)
    return f"{username}@{team}.com"

def strip_suffixes(parts: list[str]) -> list[str]:
    # Keep removing unnecessary suffixes for emails
    if parts is not None:
        while len(parts) > 1 and parts[-1].lower() in suffixes:
            parts = parts[:-1]
        return parts
        


def to_ascii(text: str) -> str:
    # Function to convert the encoding format of the text to ASCII
    if text is None or not isinstance(text, str):
        return ""
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    

def excel_styler(writer: pd.ExcelWriter):

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid" )
    header_font = Font(color="FFFFFF", bold=True, size=11)

    for sheetname, excel_ws in writer.sheets.items():
        excel_ws.freeze_panes = "A2"
        #Style headers
        for cell in excel_ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal= "center", vertical= "center")

        #Autofit column widths with padding
        for col in excel_ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max((len(str(cell.value or "")) for cell in col), default=12)
            excel_ws.column_dimensions[col_letter].width = max_len + 3



def load_and_standardise(folder: Path) -> pd.DataFrame|None:
    raw_roster_list = []
    for filepath in sorted(folder.glob("*.xls")):
        #Skip junk files in folder
        if filepath.name.startswith("~$"):
            continue

        # Determine if the files are Excel files or HTML files disguised with an .xls extension
        # and create dataframes based on each scenario
        try:
            df = pd.read_excel(filepath, dtype=str)
        except:
            tables = pd.read_html(filepath, encoding="utf-8")
            df = tables[0].astype(str)
            
        df.columns = [c.strip().lower() for c in df.columns]
        df["Team"] = filepath.stem.removeprefix("sportsref_download_")
        raw_roster_list.append(df)
        log.info(f"Loaded {len(df)} rows from {filepath.name}")

    if not raw_roster_list or len(raw_roster_list) == 0:
        raise FileNotFoundError(f"No Excel files found in NBA_Rosters")

    #Create a master dataframe from each team roster
    return pd.concat(raw_roster_list, ignore_index=True)
    

def main():
    raw_df = load_and_standardise(RAW_DIR)
    if raw_df is not None:
        cleaned_df = raw_df.drop(columns=['no.', 'pos', 'ht', 'wt', 'birth date', 'birth', 'exp',
       'college'])
        ascii_names = cleaned_df["player"].apply(to_ascii) #Change encoding to ascii
        ascii_names = ascii_names.astype(str).str.strip().str.split() # Split the player names
        
        clean_parts = ascii_names.apply(strip_suffixes)
                
        cleaned_df["clean_parts"] = clean_parts
        cleaned_df["email"] = cleaned_df.apply(format_email_column, axis=1)

        final_df = cleaned_df.drop(columns=["Team", "clean_parts"])

        with pd.ExcelWriter(MasterFile, engine="openpyxl") as writer:
            final_df.to_excel(writer, sheet_name="Player_emails", index=False)
            excel_styler(writer=writer)

        log.info(f"Success! Saved formatted Excel workbook to: {MasterFile}")


if __name__ == "__main__":
    main()