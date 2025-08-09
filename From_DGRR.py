import os 
import glob
import calendar
import pandas as pd
import numpy as py
import tkinter as tk
import tkcalendar

from tkcalendar import Calendar
from datetime import datetime
from datetime import datetime
from pathlib import Path


pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", None)

"""
def get_month_year():
    date = select_Date.get_date()
    label.config(text=f"Selected: {date}")
    return date
"""

selected_date = None

def get_month_year():
    global selected_date
    selected_date = select_Date.get_date()
    root.destroy()
 
root = tk.Tk()  
root.title("Machine Daily Sales")
root.geometry("400x350")

label_prompt = tk.Label(root, text="Select a Month and Year.")
label_prompt.pack(pady=(20,5))

# DATE ENTRY WIDGET
select_Date = Calendar(root, width=16, background="black", foreground="white", borderwidth=2)
select_Date.pack(pady=5)

btn = tk.Button(root, text="Get Date", command=get_month_year)
btn.pack()

label = tk.Label(root, text="")
label.pack(pady=10)

root.mainloop()

# TO MAKE THE DATE FROM TK.CALENDAR USABLE
if selected_date:
    date_object = datetime.strptime(selected_date, "%m/%d/%y")
    selected_month = date_object.strftime("%B").upper()
    selected_year = date_object.year

# CURRENT DATE
def current_day():
    day = datetime.now().day
    return day

def current_month():
    month = datetime.now()
    return month.strftime("%B").upper()

def current_year():
    year = datetime.now().year
    return year

def to_get_months():
    while True:
            dgrr_month_one = input("Enter the month you want to access: ").upper()

            match dgrr_month_one:
                case "JANUARY"|"FEBRUARY"|"MARCH"|"APRIL"|"MAY"|"JUNE"|"JULY"|"AUGUST"|"SEPTEMBER"|"OCTOBER"|"NOVEMBER"|"DECEMBER":
                    return dgrr_month_one
                case _:
                    print ("Please enter a valid month.")

def to_get_year():
     while True:
        dgrr_year_one = input("Enter the year you want to access: ")
        if not dgrr_year_one.isdigit():
            print("Enter a valid number.")
        elif len(dgrr_year_one) != 4:
            print("Enter 4 digit numbers only.")
        else:
            return int(dgrr_year_one)
        
# TO GET AN ENDING DATE FOR A MONTH
"""while True:
    end_date = input("Enter your Ending date: ") 
    if not end_date.isnumeric():
        print("End date must be a valid day.")
    else:
        end_date = int(end_date)
        if end_date <= 0: 
            print("End date must be greater than 0.")
        elif end_date > 31: 
            print("End date must not exceed to 31.")
        else:
            break"""

end_date = 31

# TO GET A SPECIFIC DATE PER DAY. to_get_datep[0]
to_get_date = list(range(1, end_date + 1))

"""def to_get_days_in_month(month: str, year: int) -> int:    
    month_index = list(calendar.month_name).index(month.capitalize())
        

    def get_days_in_month(month: str, year: int) -> int:
    #Get the number of days in the specified month and year.
    import calendar
    month_index = list(calendar.month_name).index(month.capitalize())
    return calendar.monthrange(year, month_index)[1]

# Get inputs
selected_month = to_get_month()
selected_year = to_get_year()

# Get number of days in the selected month/year
num_days = get_days_in_month(selected_month, selected_year)

# Generate file paths for all days
file_paths = []
for day in range(1, num_days + 1):
    day_str = str(day).zfill(2)
    path = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {day_str}.xlsx'
    file_paths.append(path)

# Example: print all file paths
for path in file_paths:
    print(path)"""


#selected_month = to_get_months()
#selected_year = to_get_year()

selected_month = "MAY"
selected_year = 2025

#indexed_months = list(selected_month).index()

for month in range(1,13):
    lastday = calendar.monthrange(selected_year, month)[1]
    #print(f"last day of{calendar.month_name[month]} {selected_year} is {lastday}")
    #print(lastday)


# END OF MONTH DGRR LINK
#dgrr_link = f"C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{year}\{month} {year}\NEW daily sales {month}"

dgrr_link_1 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[0]}.xlsx'
dgrr_link_2 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[1]}.xlsx'
dgrr_link_3 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[2]}.xlsx'
dgrr_link_4 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[3]}.xlsx'
dgrr_link_5 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[4]}.xlsx'
dgrr_link_6 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[5]}.xlsx'
dgrr_link_7 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[6]}.xlsx'
dgrr_link_8 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[7]}.xlsx'
dgrr_link_9 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[8]}.xlsx'
dgrr_link_10 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[9]}.xlsx'

dgrr_link_11 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[10]}.xlsx'
dgrr_link_12 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[11]}.xlsx'
dgrr_link_13 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[12]}.xlsx'
dgrr_link_14 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[13]}.xlsx'
dgrr_link_15 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[14]}.xlsx'
dgrr_link_16 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[15]}.xlsx'
dgrr_link_17 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[16]}.xlsx'
dgrr_link_18 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[17]}.xlsx'
dgrr_link_19 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[18]}.xlsx'
dgrr_link_20 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[19]}.xlsx'

dgrr_link_21 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[20]}.xlsx'
dgrr_link_22 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[21]}.xlsx'
dgrr_link_23 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[22]}.xlsx'
dgrr_link_24 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[23]}.xlsx'
dgrr_link_25 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[24]}.xlsx'
dgrr_link_26 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[25]}.xlsx'
dgrr_link_27 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[26]}.xlsx'
dgrr_link_28 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[27]}.xlsx'
dgrr_link_29 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[28]}.xlsx'
dgrr_link_30 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[29]}.xlsx'

dgrr_link_31 = f'C:\\Users\\Dell\\Desktop\\ebingo Reports\\NEW DGRR\\{selected_year}\\{selected_month} {selected_year}\\NEW daily sales {selected_month} {to_get_date[30]}.xlsx'

links = [dgrr_link_1, dgrr_link_2, dgrr_link_3, dgrr_link_4, dgrr_link_5, dgrr_link_6, dgrr_link_7, dgrr_link_8, dgrr_link_9, dgrr_link_10, dgrr_link_11, dgrr_link_12, dgrr_link_13, dgrr_link_14, dgrr_link_15,dgrr_link_16, dgrr_link_17, dgrr_link_18, dgrr_link_19, dgrr_link_20, dgrr_link_21, dgrr_link_22, dgrr_link_23, dgrr_link_24, dgrr_link_25,dgrr_link_26, dgrr_link_27, dgrr_link_28, dgrr_link_29, dgrr_link_30, dgrr_link_31]


# TO TEST LINK
try:
    dgrr = Path(dgrr_link_28)
    if dgrr.exists():
        read_dgrr_file = pd.ExcelFile(dgrr_link_28)
        print("Your file is working.")
    else:
        print("File not found.")
except IndexError:
    print("Index Error")
except Exception as e:
    print(f"Error:", e)

def pgi_daily(dggr_link_date):
    # TO READ FBM DSR FROM DGRR
    pgi_dsr = pd.read_excel(dggr_link_date, sheet_name="DSR PGI", header=None)
    
    # SELECT FBM VAULT NUMBER AND TOTAL IN, AND RENAME COLUMN
    supplier_name = pgi_dsr[[0, 1, 6]]
    supplier_name = supplier_name.iloc[10:20].reset_index(drop=True)
    supplier_name.rename(columns={0:"PGI VLT", 1:"SUPPLIER GAME", 6:"PGI TOTAL IN"}, inplace=True)
    return supplier_name

def zest_daily(dggr_link_date):
    # TO READ FBM DSR FROM DGRR
    zest_dsr = pd.read_excel(dggr_link_date, sheet_name="DSR ZEST", header=None)
    
    # SELECT FBM VAULT NUMBER AND TOTAL IN, AND RENAME COLUMN
    supplier_name = zest_dsr[[0, 1, 6]]
    supplier_name = supplier_name.iloc[10:48].reset_index(drop=True)
    supplier_name.rename(columns={0:"ZEST VLT", 1:"SUPPLIER GAME", 6:"ZEST TOTAL IN"}, inplace=True)
    return supplier_name

def ortiz_daily(dggr_link_date):
    # TO READ FBM DSR FROM DGRR
    ortiz_dsr = pd.read_excel(dggr_link_date, sheet_name="DSR ORTIZ", header=None)
    
    # SELECT FBM VAULT NUMBER AND TOTAL IN, AND RENAME COLUMN
    supplier_name = ortiz_dsr[[0, 1, 6]]
    supplier_name = supplier_name.iloc[10:30].reset_index(drop=True)
    supplier_name.rename(columns={0:"ORTIZ VLT", 1:"SUPPLIER GAME", 6:"ORTIZ TOTAL IN"}, inplace=True)
    return supplier_name

def dingo_daily(dggr_link_date):
    # TO READ FBM DSR FROM DGRR
    dingo_dsr = pd.read_excel(dggr_link_date, sheet_name="DSR DINGO", header=None)
    
    # SELECT FBM VAULT NUMBER AND TOTAL IN, AND RENAME COLUMN
    supplier_name = dingo_dsr[[0, 1, 6]]
    supplier_name = supplier_name.iloc[10:63].reset_index(drop=True)
    supplier_name.rename(columns={0:"DINGO VLT", 1:"SUPPLIER GAME", 6:"DINGO TOTAL IN"}, inplace=True)
    return supplier_name


def fbm_daily(dggr_link_date):
    
    # TO READ FBM DSR FROM DGRR
    fbm_dsr = pd.read_excel(dggr_link_date, sheet_name="DSR FBM", header=None)
    
    # SELECT FBM VAULT NUMBER AND TOTAL IN, AND RENAME COLUMN
    supplier_name = fbm_dsr[[0, 1, 6]]
    supplier_name = supplier_name.iloc[10:106].reset_index(drop=True)
    supplier_name.rename(columns={0:"FBM VLT", 1:"SUPPLIER GAME", 6:"FBM TOTAL IN"}, inplace=True)
    return supplier_name
