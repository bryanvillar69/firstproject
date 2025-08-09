import pandas as pd
import numpy as py
import glob

from datetime import datetime
from pathlib import Path


pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", None)

def day():
    current_day = datetime.now()
    return int(current_day.strftime("%d"))

def month():
    current_month = datetime.now()
    return current_month.strftime("%B").upper()

def year():
    current_year = datetime.now().year
    return current_year
    

current_day = day()
current_month = month()
current_year = year()

def convert_to_string(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(str)
    return supplier_name

def convert_to_float(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(float)
    return supplier_name

def convert_to_integer(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(int)
    return supplier_name

# RENAME COLUMNS
def rename_columns(supplier_name, old_name, new_name):
    supplier_name.rename(columns={old_name: new_name}, inplace=True)
    return supplier_name


# PATH FOR PARTIAL FORMS
partial_form_xlsx = f"C:\\Users\\Dell\\Desktop\\ebingo Reports\\PARTIAL FORMS\\{current_year}\\PARTIAL FORMS {current_month} {current_year}.xlsx"

# READ PARTIAL FORM
read_partial_form = pd.ExcelFile(partial_form_xlsx)

# SELECT CURRENT SHEET OF EXCEL
current_sheet = pd.read_excel(partial_form_xlsx, sheet_name=read_partial_form.sheet_names[-1], header=None)

# SELECT IMPORTANT COLUMNS IN PARTIAL FORMS
dingo_grand_total = current_sheet.iloc[107:160, [11]].reset_index(drop=True)


# TO SELECT DINGO FILE IN FLASH DRIVE
dingo_folder = Path("F:\\EBINGO BAGUIO\\")
dingo_pattern = "**/ReportAccountingPartial-*.xls"

dingo_file_match = list(dingo_folder.glob(dingo_pattern))

if dingo_file_match:
    dingo_folder = dingo_file_match[0]
    print("DINGO File is working.")
    dingo = pd.read_excel(dingo_folder)
else:
    print("File not found.")


# RENAME DINGO_GRAND_TOTAL
rename_columns(dingo_grand_total, 11, "DINGO GRAND TOTAL")

# SELECT IMPORTANT COLUMNS IN DINGO PARTIAL AND FIX INDEX
dingo = dingo.iloc[14:, [4, 8]].reset_index(drop=True)

# RENAME COLUMNS (CALL)
rename_columns(dingo, "Unnamed: 4", "DINGO VAULT")
rename_columns(dingo, "Unnamed: 8", "DINGO CASH IN")

# CLEAN UP COLUMNS. REMOVE STRINGS, NAN AND OTHERS
dingo = dingo[dingo["DINGO VAULT"].apply(lambda numbers: str(numbers).replace(",", "").replace(".","").isnumeric())]
dingo = dingo[dingo["DINGO CASH IN"].apply(lambda numbers: str(numbers).replace(",", "").replace(".","").isnumeric())]

dingo["DINGO CASH IN"] = dingo["DINGO CASH IN"].str.replace(",","", regex=False)

# CONVERT DINGO CASH IN TO INTEGER
convert_to_float(dingo, "DINGO CASH IN")
convert_to_integer(dingo, "DINGO CASH IN")
 
# SORT LOWEST TO HIGHEST BASED FROM DINGO VAULT
dingo = dingo.sort_values(by="DINGO VAULT", ascending=True).reset_index(drop=True)

# REMOVE DUPLICATE NUMBERS
dingo = dingo.drop_duplicates(subset="DINGO VAULT")

# DINGO VAULT SHOULD ALWAYS BE 1 - 53. ITS LIKE CREATING A NEW WORKSHEET, COLUMNED NAME DINGO VAULT AND CONTENTS ARE RANGE(1, 54)
complete_vlt = range(1, 54)
complete = pd.DataFrame({"DINGO VAULT": complete_vlt})

# MERGE BOTH "COMPLETE" AND "DINGO"
dingo_full = complete.merge(dingo, on="DINGO VAULT", how="left")

# SUBTRACT DINGO CASH IN TO DINGO GRAND TOTAL
dingo_full["DINGO CASH IN"] = dingo_full["DINGO CASH IN"] - dingo_grand_total["DINGO GRAND TOTAL"]

# CONVERT TO STRING AND REMOVE NAN STRINGS, COMMA AND CONVERT TO FLOAT AND TO INTEGER
convert_to_string(dingo_full, "DINGO CASH IN")
dingo_full = dingo_full[dingo_full["DINGO CASH IN"].apply(lambda numbers: str(numbers).replace(",", "").replace(".","").isnumeric())]
dingo_full["DINGO CASH IN"] = dingo_full["DINGO CASH IN"].str.replace(",","", regex=False)
convert_to_float(dingo_full, "DINGO CASH IN")
convert_to_integer(dingo_full, "DINGO CASH IN")

# GET ONLY ROWS WITH 2000 AND ABOVE BASED FROM DINGO CASH IN
dingo_full = dingo_full[dingo_full["DINGO CASH IN"] > 2000]

# INSERT DINGO SPACE AND DINGO POUCH COLUMNS 
dingo_full["DINGO SPACE"] = ""
dingo_full["DINGO POUCH"] = range(1, len(dingo_full["DINGO VAULT"]) +1)

# RENAME AGAIN FOR FINAL
rename_columns(dingo_full, "DINGO VAULT", "")
rename_columns(dingo_full, "DINGO CASH IN", "")
rename_columns(dingo_full, "DINGO SPACE", "DINGO")
rename_columns(dingo_full, "DINGO POUCH", "")

# TO SAVE EXCEL FILE
try:
    dingo_full.to_csv("c:\\Users\\Dell\\Desktop\\BRYAN FILES\\BRYAN PROGRAM\\DINGO_partial.csv", index=False)
    print("Congratulations! Your partial has been generated successfully.")
except FileNotFoundError:
    print("Kindly check the spelling, capitalization, and spacing of your folders destination.")
except Exception as e:
    print("Sorry, your file has not been generated.")
    print("error:", e)


# to make this to an EXE file
# pyinstaller --onefile --windowed --icon=DINGO_ONLY_ICON-01.ico dingo_only.py
