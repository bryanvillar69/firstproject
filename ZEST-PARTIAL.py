import pandas as pd
import numpy as py
import os

from datetime import datetime
from pathlib import Path

# TO GET CURRENT MONTH, DAY AND YEAR FUNCTION

def current_day():
    day = datetime.now()
    return day.strftime("%d")

def current_month():
    month = datetime.now()
    return month.strftime("%B").upper()

def current_year():
    year = datetime.now().year
    return year

# NO. OF ROWS AND COLUMNS
def no_rows(name_of_excel):
    name_of_excel.shape[0]
    return name_of_excel

def no_cols(name_of_excel):
    name_of_excel.shape[1]
    return name_of_excel

# CONVERTION TO STRING, FLOAT, INTEGERS
def convert_to_string(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(str)
    return supplier_name

def convert_to_float(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(float)
    return supplier_name

def convert_to_integer(supplier_name, column_name):
    supplier_name[column_name] = supplier_name[column_name].astype(int)
    return supplier_name

# RENAMING COLUMNS
def rename(supplier_name,old_name, new_name):
    supplier_name.rename(columns={old_name:new_name}, inplace=True)
    return supplier_name



# rows = no_rows(zest_grand_total)
# columns = no_cols(zest_grand_total)
# print(rows, columns)

# PARTIAL FORMS LINK
partial_forms_link = f"C:\\Users\\Dell\\Desktop\\ebingo Reports\\PARTIAL FORMS\\{current_year()}\\PARTIAL FORMS {current_month()} {current_year()}.xlsx"

# TO READ EXCEL FILE 
read_excel = pd.ExcelFile(partial_forms_link)

# SELECT CURRENT SHEET
current_sheet = pd.read_excel(partial_forms_link, sheet_name=read_excel.sheet_names[-1], header=None)

# SELECT IMPORTANT COLUMMNS AND RESET INDEX
zest_grand_total = current_sheet.iloc[195:233, [11]].reset_index(drop=True)


# SELECT ZEST FROM FLASH DRIVE
zest_flash_drive_link = Path("F:\\")
zest_file_pattern = "T2_RAF_E_BINGO_BAGUIO_Terminal_Summary_*.csv"

zest_path = list(zest_flash_drive_link.glob(zest_file_pattern))

if zest_path:
    zest_flash_drive_link = zest_path[0]
    zest = pd.read_csv(zest_flash_drive_link)
    print("File is Working.")
else:
    print("File not found. Check your file path.")

# renaming 
rename(zest_grand_total, 11, "ZEST GRAND TOTAL")



# SELECT IMPORTANT COLUMNS
zest = zest[["Term ID", "Cash In"]]

#CONVERT TERM ID ONLY TO STRING 
convert_to_string(zest,"Term ID")

# REMOVE NANS, STRINGS, COMMAS AND PERIODS.
zest = zest[zest["Term ID"].apply(lambda number: number.replace(",","").replace(".","").isnumeric())]
zest = zest[zest["Cash In"].apply(lambda number: number.replace(",","").replace(".","").isnumeric())]

zest["Term ID"] = zest["Term ID"].str.replace(",","", regex=False)
zest["Cash In"] = zest["Cash In"].str.replace(",","", regex=False)


# CONVERT TO FLOAT AND TO INTEGER
convert_to_float(zest,"Term ID")
convert_to_float(zest,"Cash In")

convert_to_integer(zest,"Term ID")
convert_to_integer(zest,"Cash In")

# ZEST TERM ID MUST BE 1 - 38
complete_vlt = range(1, 39)
complete = pd.DataFrame({"Term ID": complete_vlt})

zest_full = complete.merge(zest, on="Term ID", how="left")


# SUBTRACT CASH IN FROM GRAND TOTAL:
zest_full["Cash In"] = zest_full["Cash In"] - zest_grand_total["ZEST GRAND TOTAL"]

#CONVERT CASH IN ONLY TO STRING 
convert_to_string(zest_full,"Cash In")

# REMOVE NANS, STRINGS, COMMAS AND PERIODS AGAIN
zest_full = zest_full[zest_full["Cash In"].apply(lambda number: number.replace(",","").replace(".","").isnumeric())]
zest_full["Cash In"] = zest_full["Cash In"].str.replace(",","", regex=False)


# CONVERT CASH IN ONLY TO FLOAT AND THEN INTEGER
convert_to_float(zest_full,"Cash In")
convert_to_integer(zest_full,"Cash In")

# SELECT ONLY ROWS WITH 2000 ABOVE
zest_full = zest_full[zest_full["Cash In"] > 2000]

# ADD SPACE AND POUCH COLUMNS
zest_full["ZEST"] = ""
zest_full[""] = range(1, len(zest_full["Term ID"]) +1)

# RENAME COLUMNS FOR FINAL
rename(zest_full, "Term ID", "")
rename(zest_full, "Cash In", "")

# TO GENERATE A CSV FILE.
try:
    zest_full.to_csv(f"C:\\Users\\Dell\\Desktop\\BRYAN FILES\\BRYAN PROGRAM\\ZEST_partial.csv", index=False)
    print("You have successfully generated Zest.")
except FileNotFoundError:
    print("Kindly check the spelling, capitalization, and spacing of your folders destination.")
except Exception as e:
    print("Sorry, your file has not been generated.")
    print("error:", e)


# PY TO EXE
# pyinstaller --onefile --windowed --icon=ZEST_ONLY_ICON-01.ico zest_only.py


