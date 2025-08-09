import pandas as pd
import numpy as py
import glob
import os
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

# PATH FOR PARTIAL FORMS
partial_form_xlsx = f"C:\\Users\\Dell\\Desktop\\ebingo Reports\\PARTIAL FORMS\\{current_year}\\PARTIAL FORMS {current_month} {current_year}.xlsx"

# READ PARTIAL FORM EXCEL FILE
read_partial_form = pd.ExcelFile(partial_form_xlsx)

# SELECT CURRENT SHEET OF PARTIAL FORM
current_sheet_partial_form = pd.read_excel(partial_form_xlsx, sheet_name=read_partial_form.sheet_names[-1], header=None)

# GET GRAND TOTAL OF FBM FROM PARTIAL FROMS
grand_total = current_sheet_partial_form.iloc[4:100, [13]].reset_index(drop=True)

# RENAME GRAND TOTAL FROM 13 TO "Grand Total"
grand_total.rename(columns={13:"FBM GRAND TOTAL"}, inplace=True)


# PATH OF FBM FILE FROM FLASH DRIVE 
fbm_folder_path = Path(r"F:\\Server Reports\\Sessions\\")
fbm_pattern = "**/ReportAccountingPartial-*.xls"

# COMBINE fbm_folder_path and fbm_pattern
fbm_file_match = list(fbm_folder_path.glob(fbm_pattern))


if fbm_file_match:
    fbm_folder_path = fbm_file_match[0]
    print("FBM File is working")
    fbm = pd.read_excel(fbm_folder_path)
else:
    print("File not found")


def rename_column(supplier_name, old_name, new_name):
    supplier_name.rename(columns={old_name:new_name}, inplace=True)
    return supplier_name

def convert_to_string(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(str)
    return supplier_name

def convert_to_float(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(float)
    return supplier_name

def convert_to_integer(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(int)
    return supplier_name

def to_complete_vault_number():
    complete = range(1, 101)


# TO RENAME COLUMNS
rename_column(fbm, "Unnamed: 8", "FBM VAULT")
rename_column(fbm, "Unnamed: 9", "FBM CASH IN")

# COLUMN WE ONLY NEED
fbm = fbm[["FBM VAULT", "FBM CASH IN"]]

# CONVERT FBM VAULT TO STRING
convert_to_string(fbm, "FBM VAULT")

# THIS CLEANS UP WHOLE SPREAD SHEET. REMOVES NAN, AND OTHER STRINGS. USE BOTH OF THAT LINE
fbm = fbm[fbm["FBM VAULT"].apply(lambda x: x.replace(",", "").replace(".", "").isnumeric())]
fbm["FBM CASH IN"] = fbm["FBM CASH IN"].str.replace(",", "", regex=False)

# CONVERT FBM VAULT TO FLOAT AND TO INTEGER
convert_to_float(fbm, "FBM VAULT")
convert_to_integer(fbm, "FBM VAULT")

# CONVERT FBM CASH IN TO STRINGM TO FLOAT AND TO INTEGER
convert_to_string(fbm, "FBM CASH IN")
convert_to_float(fbm, "FBM CASH IN")
#convert_to_integer(fbm, "FBM CASH IN")

#  SORT VALUE BY FBM VAULT (1 - 96)
fbm = fbm.sort_values(by="FBM VAULT", ascending=True)


# TO REMOVE DUPLICATES
fbm = fbm.drop_duplicates(subset="FBM VAULT")

# TO COMPLETE VLT
complete_vlt = list(range(1,97))
complete = pd.DataFrame({"FBM VAULT": complete_vlt })

# TO MERGE COMPLETE TO FBM
fbm_full = complete.merge(fbm, on="FBM VAULT", how="left")



# SUBTRACT DINGO CASH IN TO DINGO GRAND TOTAL
fbm_full["FBM CASH IN"] = fbm_full["FBM CASH IN"] - grand_total["FBM GRAND TOTAL"]

# CONVERT TO STRING AND REMOVE NAN STRINGS, COMMA AND CONVERT TO FLOAT AND TO INTEGER
convert_to_string(fbm_full, "FBM CASH IN")
dingo_full = fbm_full[fbm_full["FBM CASH IN"].apply(lambda numbers: str(numbers).replace(",", "").replace(".","").isnumeric())]
dingo_full["FBM CASH IN"] = fbm_full["FBM CASH IN"].str.replace(",","", regex=False)
convert_to_float(fbm_full, "FBM CASH IN")
convert_to_integer(fbm_full, "FBM CASH IN")

# TO SORT HIGHEST TO LOWEST BASED FROM FBM CASH IN
fbm_full = fbm_full.sort_values(by="FBM CASH IN", ascending=False)

# SELECT FBM CASH IN ROWS WITH 2000 AND ABOVE ONLY
fbm_full = fbm_full[fbm_full["FBM CASH IN"] > 2000] 


# SORT TO LOWEST TO HIGHEST BASED FROM FBM VAULT
fbm_full = fbm_full.sort_values(by="FBM VAULT", ascending=True)

# FIX INDEX TO 1-96
fbm_full = fbm_full.reset_index(drop=True)


# ADD SPACE AND POUCH COLUMNS
fbm_full["FBM SPACE"] = ""
fbm_full["FBM POUCH"] = range(1, len(fbm_full["FBM VAULT"]) + 1)


# CONVERT FBM CASH IN TO INTEGER AGAIN
convert_to_integer(fbm_full, "FBM CASH IN")

# CHANGE COLUMN NAMES AGAIN.
rename_column(fbm_full, "FBM VAULT", "")
rename_column(fbm_full, "FBM CASH IN", "")
rename_column(fbm_full, "FBM SPACE", "FBM")
rename_column(fbm_full, "FBM POUCH", "")


print(fbm_full)
# TO SAVE EXCEL FILE
try:
    fbm_full.to_csv("c:\\Users\\Dell\\Desktop\\BRYAN FILES\\BRYAN PROGRAM\\FBM_partial.csv", index=False)
    print("Congratulations! Your partial has been generated successfully.")
except FileNotFoundError:
    print("Kindly check the spelling, capitalization, and spacing of your folders destination.")
except Exception as e:
    print("Sorry, your file has not been generated.")
    print("error:", e)


# to make this to an EXE file
# pyinstaller --onefile --windowed --icon=FBM_ONLY_ICON-01.ico fbm_only.py





