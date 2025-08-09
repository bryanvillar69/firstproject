import pandas as pd
import glob
import os
import numpy as np
from pathlib import Path
from datetime import datetime

pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", None)

#dates for partial form link

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


#link to partial form
partial_form_xlsx = f"C:\\Users\\Dell\\Desktop\\ebingo Reports\\PARTIAL FORMS\\{current_year}\\PARTIAL FORMS {current_month} {current_year}.xlsx"
#partial_form_xlsx = f"C:\\Users\\Dell\\Desktop\\ebingo Reports\\PARTIAL FORMS\\2025\\PARTIAL FORMS JUNE 2025.xlsx"

# to test partial form is working using link
try:
    with open(partial_form_xlsx, "r") as file:
        print("Partial_form is working.")
except Exception:
    print("Partial_form file path is wrong.")

#to read partial form
reading_partial_form = pd.ExcelFile(partial_form_xlsx)

#selecting current sheet of partial_forms
partial_forms = pd.read_excel(partial_form_xlsx, sheet_name=reading_partial_form.sheet_names[-1], header=None)

# to open partial form and read GRAND TOTAL FOR FBM, DINGO, AND ZEST and reset their index
def grand_total(total, start_row, end_row, column_name):
    result = total.iloc[start_row: end_row, [column_name]].reset_index(drop=True).astype(int) # slice data, fix indexes and convert it to integers
    return result

def change_column_name_partial(partial_forms):
    partial_forms.columns = ["Grand Total"]
    return partial_forms


# converted to function
#fbm_grand_totals = partial_forms.iloc[4:100, [13]]
#dingo_grand_total = partial_forms.iloc[107:160, [11]]
#zest_grand_total = partial_forms.iloc[195:233, [11]]


fbm_grand_total = grand_total(partial_forms, 4, 100, 13)
dingo_grand_total = grand_total(partial_forms, 107, 160, 11)
zest_grand_total = grand_total(partial_forms, 195, 233, 11)

fbm_grand_total = change_column_name_partial(fbm_grand_total)


# number of rows and columns
def number_of_rows(num_rows):
    return num_rows.shape[0]

def number_of_columns(num_cols):
    return num_cols.shape[1]


fbm_total_rows = number_of_rows(fbm_grand_total)
fbm_total_columns = number_of_columns(fbm_grand_total)

dingo_total_rows = number_of_rows(dingo_grand_total)
dingo_total_columns = number_of_columns(dingo_grand_total)

zest_total_rows = number_of_rows(zest_grand_total)
zest_total_columns = number_of_columns(zest_grand_total)

#print(f"total rows: {zest_total_rows}, total columns: {zest_total_columns}")

#print(fbm_grand_total.dtypes)

# to open FBM file in flash drive:
fbm_folder = Path(r"F:\\Server Reports\\Sessions\\")

# sample file path of FBM:  F:\\Server Reports\\Sessions\\6_30_2025_8_44_14_AM\\Game Reports\\ReportAccountingPartial-20250630084414.xls"
fbm_file_match = list(fbm_folder.glob("**/ReportAccountingPartial-*.xls"))

if fbm_file_match:
    fbm_file_path = fbm_file_match[0]
    print("FBM FILE IS WORKING! Found:", fbm_file_path)

    fbm = pd.read_excel(fbm_file_path)
else:
    print("No FBM file found.")

# to open DINGO file in flash drive:
dingo_folder = Path(r"F:\\EBINGO BAGUIO")
dingo_pattern = "ReportAccountingPartial-*.xls"

dingo_file_match = list(dingo_folder.glob(dingo_pattern))

if dingo_file_match:
    dingo_file_path = dingo_file_match[0]
    print("DINGO FILE IS WORKING! Found:", dingo_file_path)

    dingo = pd.read_excel(dingo_file_path)
    #print(dingo.head())
else:
    print("No Dingo file found.")

"""
# to open ZEST file in flash drive:
zest_folder = Path(r"F:\\")
zest_pattern = "T2_RAF_E_BINGO_BAGUIO_Terminal_Summary_*.csv"

zest_file_match = list(zest_folder.glob(zest_pattern))

if zest_file_match:
    zest_file_path = zest_file_match[0]
    print("ZEST FILE IS WORKING! Found", zest_file_path)

    zest = pd.read_csv(zest_file_path)
    print(zest.head())
else:
    print("No Zest file found.")
"""

# convert single column.
def convert_to_int(partial, column_name):
    partial[column_name] = partial[column_name].astype(int)
    return partial

def convert_to_str(partial, column_name):
    partial[column_name] = partial[column_name].astype(str)
    return partial

def convert_to_float(partial, column_name):
    partial[column_name] = partial[column_name].astype(float)
    return partial

# select columns function:
def select_column(selected_columns, column_one, column_two):
    columnns = selected_columns.iloc[:, [column_one,column_two]]
    return columnns

#change column name:
def change_column_name(supplier_name, column_one, column_two):
    supplier_name.columns = [column_one, column_two]
    return supplier_name

# to remove strings, spaces and leave only numbers
def to_clean_column_one(supplier_name, column1):
    #convert to string
    supplier_name[column1] =supplier_name[column1].astype(str)

    # keeps rows that are numbers only.
    supplier_name =  supplier_name[supplier_name[column1].apply(lambda x: x.replace(",","").replace(".","").isnumeric())]

    #convert to int
    supplier_name[column1] = supplier_name[column1].astype(int)
    return supplier_name

def to_clean_column_two(supplier_name, column2):
    #convert to string
    supplier_name[column2] = supplier_name[column2].astype(str)

    # keeps rows that are numbers only. 
    supplier_name =  supplier_name[supplier_name[column2].apply(lambda x: x.replace(",","").replace(".","").isnumeric())]

    # remove commas and periods and convert to float. regex means removes every character like dots, commas and etc.
    supplier_name[column2] = supplier_name[column2].str.replace(",","", regex=False) 
    supplier_name[column2] = supplier_name[column2].astype(float)
    supplier_name[column2] = supplier_name[column2].astype(int)

    return supplier_name

#sortby cash in from highest number to lowest and select numbers 2200 and above
def sort_by_column_two(supplier_name):
    column = supplier_name.columns[1]

    supplier_name = supplier_name.sort_values(by=column, ascending=False)
    return supplier_name

# sort column one "VLT" from (fbm 1 -96), (dingo 1 -53), (zest 1-38) (total number of machines)
def sort_by_column_one(supplier_name):
    column = supplier_name.columns[0]
    supplier_name = supplier_name.sort_values(by=column, ascending=True)
    return supplier_name

# fix index to (fbm 1 -96), (dingo 1 -53), (zest 1-38) based from "VLT" of supplier
def fix_index(supplier_name):
    supplier_name = supplier_name.reset_index(drop=True)
    supplier_name.index += 1
    return supplier_name

# combined fbm partial data from flash drive and grand total from partial forms.
def combine_grand_total_fbm(supplier_name, partial_forms):
     combined_total = pd.concat([supplier_name, partial_forms], axis=1)
     return combined_total

# create new column beside grand total and subtract  fbm partial from grandtotal
def subtract_grand_total(combined, column_one, column_two):
    combined["Subtracted"] = combined[column_one] - combined[column_two]
    return combined

def dingo_subtract_grand_total(combined, column_one, column_two):
    combined["Dingo Subtracted"] = combined[column_one] - combined[column_two]
    return combined

# remove cash in and grand total and sort value of column "Subtracted" highest to lowest
def sorting(supplier_name, column_vlt, column_subtracted):
    supplier_name = supplier_name[[column_vlt, column_subtracted]]
    supplier_name = supplier_name.sort_values(by=column_subtracted, ascending=False)
    supplier_name = supplier_name[supplier_name[column_subtracted]> 2300]
    return supplier_name

def rename_column(merged_partial, column_one, column_two, column_three, column_four, 
                               column_five, column_six, column_seven, column_eight):
    
    merged_partial.rename(columns={column_one: "FBM VLT"}, inplace=True)
    merged_partial.rename(columns={column_two: "FBM CASH IN"}, inplace=True)
    merged_partial.rename(columns={column_three: "FBM SPACE"}, inplace=True)
    merged_partial.rename(columns={column_four: "FBM POUCH"}, inplace=True)
    merged_partial.rename(columns={column_five: "DINGO VLT"}, inplace=True)
    merged_partial.rename(columns={column_six: "DINGO CASH IN"}, inplace=True)
    merged_partial.rename(columns={column_seven: "DINGO SPACE"}, inplace=True)
    merged_partial.rename(columns={column_eight: "DINGO POUCH"}, inplace=True)

def rename_column_final(merged_partial):
    
    merged_partial.rename(columns={"FBM VLT" : ""}, inplace=True)
    merged_partial.rename(columns={"FBM CASH IN": ""}, inplace=True)
    merged_partial.rename(columns={"FBM SPACE": "FBM"}, inplace=True)
    merged_partial.rename(columns={"FBM POUCH": ""}, inplace=True)
    merged_partial.rename(columns={"DINGO VLT": ""}, inplace=True)
    merged_partial.rename(columns={"Dingo Subtracted": ""}, inplace=True)
    merged_partial.rename(columns={"DINGO SPACE": "DINGO"}, inplace=True)
    merged_partial.rename(columns={"DINGO POUCH": ""}, inplace=True)    

def add_column(merged_partial, column_num, column_name, column_details):
    merged_partial.insert(column_num, column_name, column_details)

    return merged_partial

def convert_dingo_blankspaces(merged_partial):
    dingo_vlt = merged_partial.columns[4]
    dingo_cash_in = merged_partial.columns[5]
    dingo_pouch = merged_partial.columns[7]

    dingo_zero_rows = (merged_partial[dingo_vlt] == "") & (merged_partial[dingo_cash_in] == "")
    #partial.loc[dingo_zero_rows, [dingo_vlt, dingo_cash_in, dingo_pouch]] = ""
    merged_partial.loc[dingo_zero_rows, [dingo_vlt, dingo_cash_in, dingo_pouch]] = "" # np.nan can be converted to blank

    return merged_partial

# function calls of FBM:
fbm = select_column(fbm ,8 ,9)
fbm = change_column_name(fbm, "FBM VLT", "FBM CASH IN")
fbm = to_clean_column_one(fbm, fbm.columns[0])
fbm = to_clean_column_two(fbm, fbm.columns[1])
fbm = sort_by_column_two(fbm)
fbm = sort_by_column_one(fbm)
fbm = fix_index(fbm)

# INSERT: REMOVE DUPLICATE AND VAULT IS ALWAYS 1 - 96
fbm = fbm.drop_duplicates(subset="FBM VLT")
complete_fbm_vlt = range(1, 97)
complete_fbm = pd.DataFrame({"FBM VLT": complete_fbm_vlt})
fbm = complete_fbm.merge(fbm, on="FBM VLT", how="left")

fbm_partial = combine_grand_total_fbm(fbm, fbm_grand_total)
fbm_partial = subtract_grand_total(fbm_partial,fbm_partial.columns[1], fbm_partial.columns[2])
fbm_partial = sorting(fbm_partial, fbm_partial.columns[0], fbm_partial.columns[3])
fbm_partial = sort_by_column_one(fbm_partial)
fbm_partial = fix_index(fbm_partial)

#function call of DINGO:
dingo = select_column(dingo, 4, 8)
dingo = change_column_name(dingo, "DINGO VLT", "DINGO CASH IN")
dingo = to_clean_column_one(dingo, dingo.columns[0])
dingo = to_clean_column_two(dingo, dingo.columns[1])
dingo = sort_by_column_two(dingo)
dingo = sort_by_column_one(dingo)
dingo = fix_index(dingo)

# INSERT REMOVE DUPLICATE AND VAULT IS ALWAYS 1 - 53
dingo = dingo.drop_duplicates(subset="DINGO VLT")
complete_dingo_vlt = range(1, 54)
complete_dingo = pd.DataFrame({"DINGO VLT": complete_dingo_vlt})
dingo = complete_dingo.merge(dingo, on="DINGO VLT", how="left")

dingo_partial = combine_grand_total_fbm(dingo, dingo_grand_total)
dingo_partial = dingo_subtract_grand_total(dingo_partial,dingo_partial.columns[1], dingo_partial.columns[2])
dingo_partial = sorting(dingo_partial, dingo_partial.columns[0], dingo_partial.columns[3])
dingo_partial = sort_by_column_one(dingo_partial)
dingo_partial = fix_index(dingo_partial)
#dingo_partial = add_final_two_columns(dingo_partial, "Space", "Pouch")


# COMBINE DATA OF fbm_partial & dingo_partial
merged_partial = pd.concat([fbm_partial, dingo_partial], axis=1)
rename_column(merged_partial, "VLT", "Subtracted", "SPACE FBM", "POUCH FBM", "DINGO VLT", "Subtracted", "SPACE DINGO", "POUCH DINGO")

# ADD COLUMNS SPACE AND POUCH FOR EACH SUPPLIER
merged_partial = add_column(merged_partial, 2, "FBM SPACE", "")
merged_partial = add_column(merged_partial, 3, "FBM POUCH", 0)

merged_partial = add_column(merged_partial, 6, "DINGO SPACE", "")
merged_partial = add_column(merged_partial, 7, "DINGO POUCH", 0)

merged_partial.replace(np.nan, "", inplace=True)


# FIND MAX RANGE OF FBM VLT FOR FBM POUCH
merged_partial["FBM VLT"] = pd.to_numeric(merged_partial["FBM VLT"], errors="coerce")
merged_partial["FBM CASH IN"] = pd.to_numeric(merged_partial["FBM CASH IN"], errors="coerce")

# FBM POUCH NUMBERING
merged_partial["FBM POUCH"] = range(1, len(merged_partial["FBM VLT"]) + 1)

# THOS ROWS WITH NaN VALUES WILL CONVERT OTHER ROWS TO NaN.
merged_partial.loc[merged_partial["FBM VLT"].isna(), "FBM POUCH"] = np.nan

# FIND MAX RANGE FOR FBM POUCH EXCLUDING NaN. Numbers only
fbm_max_range = int(merged_partial["FBM POUCH"].dropna().max())
number_start = fbm_max_range + 1


# DINGO POUCH NUMBERING - #RANGE(1, 10)
merged_partial["DINGO POUCH"] = range(number_start, len(merged_partial) + number_start)

# REMOVE ZEROS FROM DINGO
convert_dingo_blankspaces(merged_partial)

# RENAME COLUMNS FINAL
rename_column_final(merged_partial)


# TO MAKE THIS PROGRAM USABLE
try:
    merged_partial.to_csv("c:\\Users\\Dell\\Desktop\\BRYAN FILES\\BRYAN PROGRAM\\FBM-DINGO.csv", index=False)
    print("Congratulations! Your partial has been generated successfully.")
except FileNotFoundError:
    print("Kindly check the spelling, capitalization, and spacing of your folders destination.")
except Exception as e:
    print("Sorry, your file has not been generated.")
    print("error:", e)

# how to make your program a .exe
#pyinstaller --onefile --windowed --icon=FBM-DINGO.ico FBM-DINGO.py
