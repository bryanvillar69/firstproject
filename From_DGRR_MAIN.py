import pandas as pd
import numpy as py
from datetime import datetime
from num2words import num2words

from dgrr import selected_month, selected_year
from dgrr import fbm_daily, dingo_daily, ortiz_daily, zest_daily, pgi_daily
from dgrr import dgrr_link_1, dgrr_link_2, dgrr_link_3, dgrr_link_4, dgrr_link_5, dgrr_link_6, dgrr_link_7, dgrr_link_8, dgrr_link_9, dgrr_link_10, dgrr_link_11, dgrr_link_12, dgrr_link_13, dgrr_link_14, dgrr_link_15,dgrr_link_16, dgrr_link_17, dgrr_link_18, dgrr_link_19, dgrr_link_20, dgrr_link_21, dgrr_link_22, dgrr_link_23, dgrr_link_24, dgrr_link_25,dgrr_link_26, dgrr_link_27, dgrr_link_28, dgrr_link_29, dgrr_link_30, dgrr_link_31

# NUMBER OF ROWS FOR EACH SUPPLIER
fbm_rows = list(range(1, 97))
dingo_rows = list(range(1, 54))
ortiz_rows = list(range(1, 21))
zest_rows = list(range(1, 39))
perception_rows = list(range(1, 11))

# CREATE A 31 COLUMNS REPRESENTING DAYS NAMED (ONE, TWO.... THIRTY-ONE)
column_of_days = [f"{num2words(i, to='cardinal').upper()}" for i in range(1, 32)]

fbm = pd.DataFrame(columns=column_of_days, index=range(96))
fbm.insert(0, "VLT", fbm_rows)
fbm.insert(1, "SUPPLIER GAME", "")

dingo = pd.DataFrame(columns=column_of_days, index=range(53))
dingo.insert(0, "VLT", dingo_rows)
dingo.insert(1, "SUPPLIER GAME", "")

ortiz = pd.DataFrame(columns=column_of_days, index=range(20))
ortiz.insert(0, "VLT", ortiz_rows)
ortiz.insert(1, "SUPPLIER GAME", "")

zest = pd.DataFrame(columns=column_of_days, index=range(38))
zest.insert(0, "VLT", zest_rows)
zest.insert(1, "SUPPLIER GAME", "")

pgi = pd.DataFrame(columns=column_of_days, index=range(10))
pgi.insert(0, "VLT", perception_rows)
pgi.insert(1, "SUPPLIER GAME", "")

def str_convert(dataframe, column_name):
    dataframe[column_name] = dataframe[column_name].astype(str)
    return dataframe

def float_convert(dataframe, column_name):
    dataframe[column_name] = dataframe[column_name].astype(str)
    return dataframe

def int_convert(dataframe, column_name):
    dataframe[column_name] = dataframe[column_name].astype(str)
    return dataframe

def fbm_daily_sales(column_name, dggr_link):
    fbm_one = fbm_daily(dggr_link)

    #inserting columns from fbm_daily to dgrr_spreadsheet. 
    fbm["VLT"] = fbm_one.iloc[:, 0].values
    fbm["SUPPLIER GAME"] = fbm_one.iloc[:, 1].values
    fbm[column_name] = fbm_one.iloc[:, 2].values
    

def dingo_daily_sales(column_name, dggr_link):
    dingo_one = dingo_daily(dggr_link)

    #inserting columns from dingo_daily to dgrr_spreadsheet. 
    dingo["VLT"] = dingo_one.iloc[:, 0].values
    dingo["SUPPLIER GAME"] = dingo_one.iloc[:, 1].values
    dingo[column_name] = dingo_one.iloc[:, 2].values

def ortiz_daily_sales(column_name, dggr_link):
    ortiz_one = ortiz_daily(dggr_link)

    #inserting columns from ortiz_daily to dgrr_spreadsheet. 
    ortiz["VLT"] = ortiz_one.iloc[:, 0].values
    ortiz["SUPPLIER GAME"] = ortiz_one.iloc[:, 1].values
    ortiz[column_name] = ortiz_one.iloc[:, 2].values

def zest_daily_sales(column_name, dggr_link):
    zest_one = zest_daily(dggr_link)

    #inserting columns from zest_daily to dgrr_spreadsheet. 
    zest["VLT"] = zest_one.iloc[:, 0].values
    zest["SUPPLIER GAME"] = zest_one.iloc[:, 1].values
    zest[column_name] = zest_one.iloc[:, 2].values

def pgi_daily_sales(column_name, dggr_link):
    pgi_one = pgi_daily(dggr_link)

    #inserting columns from pgi_daily to dgrr_spreadsheet. 
    pgi["VLT"] = pgi_one.iloc[:, 0].values
    pgi["SUPPLIER GAME"] = pgi_one.iloc[:, 1].values
    pgi[column_name] = pgi_one.iloc[:, 2].values

# TO CHECK WETHER THE FILE HAS THE DATA IF NOT THE COLUMS AND ROWS BECOMES ZERO(0) 
def safe_call(supplier_daily_sales, day, link):
    try:
        supplier_daily_sales(day, link)
    except Exception as e:
        print("Error:", e)


# CALL FUNCTIONS
if dgrr_link_1: safe_call(fbm_daily_sales, "ONE", dgrr_link_1)
if dgrr_link_2: safe_call(fbm_daily_sales, "TWO", dgrr_link_2)
if dgrr_link_3: safe_call(fbm_daily_sales, "THREE", dgrr_link_3)
if dgrr_link_4: safe_call(fbm_daily_sales, "FOUR", dgrr_link_4)
if dgrr_link_5: safe_call(fbm_daily_sales, "FIVE", dgrr_link_5)
if dgrr_link_6: safe_call(fbm_daily_sales, "SIX", dgrr_link_6)
if dgrr_link_7: safe_call(fbm_daily_sales, "SEVEN", dgrr_link_7)
if dgrr_link_8: safe_call(fbm_daily_sales, "EIGHT", dgrr_link_8)
if dgrr_link_9: safe_call(fbm_daily_sales, "NINE", dgrr_link_9)
if dgrr_link_10: safe_call(fbm_daily_sales, "TEN", dgrr_link_10)
if dgrr_link_11: safe_call(fbm_daily_sales, "ELEVEN", dgrr_link_11)
if dgrr_link_12: safe_call(fbm_daily_sales, "TWELVE", dgrr_link_12)
if dgrr_link_13: safe_call(fbm_daily_sales, "THIRTEEN", dgrr_link_13)
if dgrr_link_14: safe_call(fbm_daily_sales, "FOURTEEN", dgrr_link_14)
if dgrr_link_15: safe_call(fbm_daily_sales, "FIFTEEN", dgrr_link_15)
if dgrr_link_16: safe_call(fbm_daily_sales, "SIXTEEN", dgrr_link_16)
if dgrr_link_17: safe_call(fbm_daily_sales, "SEVENTEEN", dgrr_link_17)
if dgrr_link_18: safe_call(fbm_daily_sales, "EIGHTEEN", dgrr_link_18)
if dgrr_link_19: safe_call(fbm_daily_sales, "NINETEEN", dgrr_link_19)
if dgrr_link_20: safe_call(fbm_daily_sales, "TWENTY", dgrr_link_20)
if dgrr_link_21: safe_call(fbm_daily_sales, "TWENTY-ONE", dgrr_link_21)
if dgrr_link_22: safe_call(fbm_daily_sales, "TWENTY-TWO", dgrr_link_22)
if dgrr_link_23: safe_call(fbm_daily_sales, "TWENTY-THREE", dgrr_link_23)
if dgrr_link_24: safe_call(fbm_daily_sales, "TWENTY-FOUR", dgrr_link_24)
if dgrr_link_25: safe_call(fbm_daily_sales, "TWENTY-FIVE", dgrr_link_25)
if dgrr_link_26: safe_call(fbm_daily_sales, "TWENTY-SIX", dgrr_link_26)
if dgrr_link_27: safe_call(fbm_daily_sales, "TWENTY-SEVEN", dgrr_link_27)
if dgrr_link_28: safe_call(fbm_daily_sales, "TWENTY-EIGHT", dgrr_link_28)
if dgrr_link_29: safe_call(fbm_daily_sales, "TWENTY-NINE", dgrr_link_29)
if dgrr_link_30: safe_call(fbm_daily_sales, "THIRTY", dgrr_link_30)
if dgrr_link_30: safe_call(fbm_daily_sales, "THIRTY", dgrr_link_31)

if dgrr_link_1: safe_call(dingo_daily_sales, "ONE", dgrr_link_1)
if dgrr_link_2: safe_call(dingo_daily_sales, "TWO", dgrr_link_2)
if dgrr_link_3: safe_call(dingo_daily_sales, "THREE", dgrr_link_3)
if dgrr_link_4: safe_call(dingo_daily_sales, "FOUR", dgrr_link_4)
if dgrr_link_5: safe_call(dingo_daily_sales, "FIVE", dgrr_link_5)
if dgrr_link_6: safe_call(dingo_daily_sales, "SIX", dgrr_link_6)
if dgrr_link_7: safe_call(dingo_daily_sales, "SEVEN", dgrr_link_7)
if dgrr_link_8: safe_call(dingo_daily_sales, "EIGHT", dgrr_link_8)
if dgrr_link_9: safe_call(dingo_daily_sales, "NINE", dgrr_link_9)
if dgrr_link_10: safe_call(dingo_daily_sales, "TEN", dgrr_link_10)
if dgrr_link_11: safe_call(dingo_daily_sales, "ELEVEN", dgrr_link_11)
if dgrr_link_12: safe_call(dingo_daily_sales, "TWELVE", dgrr_link_12)
if dgrr_link_13: safe_call(dingo_daily_sales, "THIRTEEN", dgrr_link_13)
if dgrr_link_14: safe_call(dingo_daily_sales, "FOURTEEN", dgrr_link_14)
if dgrr_link_15: safe_call(dingo_daily_sales, "FIFTEEN", dgrr_link_15)
if dgrr_link_16: safe_call(dingo_daily_sales, "SIXTEEN", dgrr_link_16)
if dgrr_link_17: safe_call(dingo_daily_sales, "SEVENTEEN", dgrr_link_17)
if dgrr_link_18: safe_call(dingo_daily_sales, "EIGHTEEN", dgrr_link_18)
if dgrr_link_19: safe_call(dingo_daily_sales, "NINETEEN", dgrr_link_19)
if dgrr_link_20: safe_call(dingo_daily_sales, "TWENTY", dgrr_link_20)
if dgrr_link_21: safe_call(dingo_daily_sales, "TWENTY-ONE", dgrr_link_21)
if dgrr_link_22: safe_call(dingo_daily_sales, "TWENTY-TWO", dgrr_link_22)
if dgrr_link_23: safe_call(dingo_daily_sales, "TWENTY-THREE", dgrr_link_23)
if dgrr_link_24: safe_call(dingo_daily_sales, "TWENTY-FOUR", dgrr_link_24)
if dgrr_link_25: safe_call(dingo_daily_sales, "TWENTY-FIVE", dgrr_link_25)
if dgrr_link_26: safe_call(dingo_daily_sales, "TWENTY-SIX", dgrr_link_26)
if dgrr_link_27: safe_call(dingo_daily_sales, "TWENTY-SEVEN", dgrr_link_27)
if dgrr_link_28: safe_call(dingo_daily_sales, "TWENTY-EIGHT", dgrr_link_28)
if dgrr_link_29: safe_call(dingo_daily_sales, "TWENTY-NINE", dgrr_link_29)
if dgrr_link_30: safe_call(dingo_daily_sales, "THIRTY", dgrr_link_30)
if dgrr_link_31: safe_call(dingo_daily_sales, "THIRTY-ONE", dgrr_link_31)

if dgrr_link_1: safe_call(ortiz_daily_sales, "ONE", dgrr_link_1)
if dgrr_link_2: safe_call(ortiz_daily_sales, "TWO", dgrr_link_2)
if dgrr_link_3: safe_call(ortiz_daily_sales, "THREE", dgrr_link_3)
if dgrr_link_4: safe_call(ortiz_daily_sales, "FOUR", dgrr_link_4)
if dgrr_link_5: safe_call(ortiz_daily_sales, "FIVE", dgrr_link_5)
if dgrr_link_6: safe_call(ortiz_daily_sales, "SIX", dgrr_link_6)
if dgrr_link_7: safe_call(ortiz_daily_sales, "SEVEN", dgrr_link_7)
if dgrr_link_8: safe_call(ortiz_daily_sales, "EIGHT", dgrr_link_8)
if dgrr_link_9: safe_call(ortiz_daily_sales, "NINE", dgrr_link_9)
if dgrr_link_10: safe_call(ortiz_daily_sales, "TEN", dgrr_link_10)
if dgrr_link_11: safe_call(ortiz_daily_sales, "ELEVEN", dgrr_link_11)
if dgrr_link_12: safe_call(ortiz_daily_sales, "TWELVE", dgrr_link_12)
if dgrr_link_13: safe_call(ortiz_daily_sales, "THIRTEEN", dgrr_link_13)
if dgrr_link_14: safe_call(ortiz_daily_sales, "FOURTEEN", dgrr_link_14)
if dgrr_link_15: safe_call(ortiz_daily_sales, "FIFTEEN", dgrr_link_15)
if dgrr_link_16: safe_call(ortiz_daily_sales, "SIXTEEN", dgrr_link_16)
if dgrr_link_17: safe_call(ortiz_daily_sales, "SEVENTEEN", dgrr_link_17)
if dgrr_link_18: safe_call(ortiz_daily_sales, "EIGHTEEN", dgrr_link_18)
if dgrr_link_19: safe_call(ortiz_daily_sales, "NINETEEN", dgrr_link_19)
if dgrr_link_20: safe_call(ortiz_daily_sales, "TWENTY", dgrr_link_20)
if dgrr_link_21: safe_call(ortiz_daily_sales, "TWENTY-ONE", dgrr_link_21)
if dgrr_link_22: safe_call(ortiz_daily_sales, "TWENTY-TWO", dgrr_link_22)
if dgrr_link_23: safe_call(ortiz_daily_sales, "TWENTY-THREE", dgrr_link_23)
if dgrr_link_24: safe_call(ortiz_daily_sales, "TWENTY-FOUR", dgrr_link_24)
if dgrr_link_25: safe_call(ortiz_daily_sales, "TWENTY-FIVE", dgrr_link_25)
if dgrr_link_26: safe_call(ortiz_daily_sales, "TWENTY-SIX", dgrr_link_26)
if dgrr_link_27: safe_call(ortiz_daily_sales, "TWENTY-SEVEN", dgrr_link_27)
if dgrr_link_28: safe_call(ortiz_daily_sales, "TWENTY-EIGHT", dgrr_link_28)
if dgrr_link_29: safe_call(ortiz_daily_sales, "TWENTY-NINE", dgrr_link_29)
if dgrr_link_30: safe_call(ortiz_daily_sales, "THIRTY", dgrr_link_30)
if dgrr_link_31: safe_call(ortiz_daily_sales, "THIRTY-ONE", dgrr_link_31)

if dgrr_link_1: safe_call(zest_daily_sales, "ONE", dgrr_link_1)
if dgrr_link_2: safe_call(zest_daily_sales, "TWO", dgrr_link_2)
if dgrr_link_3: safe_call(zest_daily_sales, "THREE", dgrr_link_3)
if dgrr_link_4: safe_call(zest_daily_sales, "FOUR", dgrr_link_4)
if dgrr_link_5: safe_call(zest_daily_sales, "FIVE", dgrr_link_5)
if dgrr_link_6: safe_call(zest_daily_sales, "SIX", dgrr_link_6)
if dgrr_link_7: safe_call(zest_daily_sales, "SEVEN", dgrr_link_7)
if dgrr_link_8: safe_call(zest_daily_sales, "EIGHT", dgrr_link_8)
if dgrr_link_9: safe_call(zest_daily_sales, "NINE", dgrr_link_9)
if dgrr_link_10: safe_call(zest_daily_sales, "TEN", dgrr_link_10)
if dgrr_link_11: safe_call(zest_daily_sales, "ELEVEN", dgrr_link_11)
if dgrr_link_12: safe_call(zest_daily_sales, "TWELVE", dgrr_link_12)
if dgrr_link_13: safe_call(zest_daily_sales, "THIRTEEN", dgrr_link_13)
if dgrr_link_14: safe_call(zest_daily_sales, "FOURTEEN", dgrr_link_14)
if dgrr_link_15: safe_call(zest_daily_sales, "FIFTEEN", dgrr_link_15)
if dgrr_link_16: safe_call(zest_daily_sales, "SIXTEEN", dgrr_link_16)
if dgrr_link_17: safe_call(zest_daily_sales, "SEVENTEEN", dgrr_link_17)
if dgrr_link_18: safe_call(zest_daily_sales, "EIGHTEEN", dgrr_link_18)
if dgrr_link_19: safe_call(zest_daily_sales, "NINETEEN", dgrr_link_19)
if dgrr_link_20: safe_call(zest_daily_sales, "TWENTY", dgrr_link_20)
if dgrr_link_21: safe_call(zest_daily_sales, "TWENTY-ONE", dgrr_link_21)
if dgrr_link_22: safe_call(zest_daily_sales, "TWENTY-TWO", dgrr_link_22)
if dgrr_link_23: safe_call(zest_daily_sales, "TWENTY-THREE", dgrr_link_23)
if dgrr_link_24: safe_call(zest_daily_sales, "TWENTY-FOUR", dgrr_link_24)
if dgrr_link_25: safe_call(zest_daily_sales, "TWENTY-FIVE", dgrr_link_25)
if dgrr_link_26: safe_call(zest_daily_sales, "TWENTY-SIX", dgrr_link_26)
if dgrr_link_27: safe_call(zest_daily_sales, "TWENTY-SEVEN", dgrr_link_27)
if dgrr_link_28: safe_call(zest_daily_sales, "TWENTY-EIGHT", dgrr_link_28)
if dgrr_link_29: safe_call(zest_daily_sales, "TWENTY-NINE", dgrr_link_29)
if dgrr_link_30: safe_call(zest_daily_sales, "THIRTY", dgrr_link_30)
if dgrr_link_31: safe_call(zest_daily_sales, "THIRTY-ONE", dgrr_link_31)

if dgrr_link_1: safe_call(pgi_daily_sales, "ONE", dgrr_link_1)
if dgrr_link_2: safe_call(pgi_daily_sales, "TWO", dgrr_link_2)
if dgrr_link_3: safe_call(pgi_daily_sales, "THREE", dgrr_link_3)
if dgrr_link_4: safe_call(pgi_daily_sales, "FOUR", dgrr_link_4)
if dgrr_link_5: safe_call(pgi_daily_sales, "FIVE", dgrr_link_5)
if dgrr_link_6: safe_call(pgi_daily_sales, "SIX", dgrr_link_6)
if dgrr_link_7: safe_call(pgi_daily_sales, "SEVEN", dgrr_link_7)
if dgrr_link_8: safe_call(pgi_daily_sales, "EIGHT", dgrr_link_8)
if dgrr_link_9: safe_call(pgi_daily_sales, "NINE", dgrr_link_9)
if dgrr_link_10: safe_call(pgi_daily_sales, "TEN", dgrr_link_10)
if dgrr_link_11: safe_call(pgi_daily_sales, "ELEVEN", dgrr_link_11)
if dgrr_link_12: safe_call(pgi_daily_sales, "TWELVE", dgrr_link_12)
if dgrr_link_13: safe_call(pgi_daily_sales, "THIRTEEN", dgrr_link_13)
if dgrr_link_14: safe_call(pgi_daily_sales, "FOURTEEN", dgrr_link_14)
if dgrr_link_15: safe_call(pgi_daily_sales, "FIFTEEN", dgrr_link_15)
if dgrr_link_16: safe_call(pgi_daily_sales, "SIXTEEN", dgrr_link_16)
if dgrr_link_17: safe_call(pgi_daily_sales, "SEVENTEEN", dgrr_link_17)
if dgrr_link_18: safe_call(pgi_daily_sales, "EIGHTEEN", dgrr_link_18)
if dgrr_link_19: safe_call(pgi_daily_sales, "NINETEEN", dgrr_link_19)
if dgrr_link_20: safe_call(pgi_daily_sales, "TWENTY", dgrr_link_20)
if dgrr_link_21: safe_call(pgi_daily_sales, "TWENTY-ONE", dgrr_link_21)
if dgrr_link_22: safe_call(pgi_daily_sales, "TWENTY-TWO", dgrr_link_22)
if dgrr_link_23: safe_call(pgi_daily_sales, "TWENTY-THREE", dgrr_link_23)
if dgrr_link_24: safe_call(pgi_daily_sales, "TWENTY-FOUR", dgrr_link_24)
if dgrr_link_25: safe_call(pgi_daily_sales, "TWENTY-FIVE", dgrr_link_25)
if dgrr_link_26: safe_call(pgi_daily_sales, "TWENTY-SIX", dgrr_link_26)
if dgrr_link_27: safe_call(pgi_daily_sales, "TWENTY-SEVEN", dgrr_link_27)
if dgrr_link_28: safe_call(pgi_daily_sales, "TWENTY-EIGHT", dgrr_link_28)
if dgrr_link_29: safe_call(pgi_daily_sales, "TWENTY-NINE", dgrr_link_29)
if dgrr_link_30: safe_call(pgi_daily_sales, "THIRTY", dgrr_link_30)
if dgrr_link_31: safe_call(pgi_daily_sales, "THIRTY-ONE", dgrr_link_31)


# ADD NEW ROW NAMED TOTAL UNDER VLT AND IT COMPUTES TOTAL NUMBER PER COLUMNS
def sum_rows(column_name, dataframe):
    new_row = {"VLT": "TPD"}
    for column_name in dataframe.columns[1:]:
        new_row[column_name] = dataframe[column_name].sum()
        
    return new_row

fbm_new_row = sum_rows("ONE", fbm)
dingo_new_row = sum_rows("ONE", dingo)
ortiz_new_row = sum_rows("ONE", ortiz)
zest_new_row = sum_rows("ONE", zest)
pgi_new_row = sum_rows("ONE", pgi)

fbm = pd.concat([fbm, pd.DataFrame([fbm_new_row])], ignore_index=True).reset_index(drop=True)
dingo = pd.concat([dingo, pd.DataFrame([dingo_new_row])], ignore_index=True).reset_index(drop=True)
ortiz = pd.concat([ortiz, pd.DataFrame([ortiz_new_row])], ignore_index=True).reset_index(drop=True)
zest = pd.concat([zest, pd.DataFrame([zest_new_row])], ignore_index=True).reset_index(drop=True)
pgi = pd.concat([pgi, pd.DataFrame([pgi_new_row])], ignore_index=True).reset_index(drop=True)

# ADD NEW COLUMN AND IT WILL ADD ALL ROWS
columns = ["ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN",
           "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN", "TWENTY",
           "TWENTY-ONE", "TWENTY-TWO", "TWENTY-THREE", "TWENTY-FOUR", "TWENTY-FIVE", "TWENTY-SIX", "TWENTY-SEVEN", "TWENTY-EIGHT", "TWENTY-NINE", "THIRTY",
           "THIRTY-ONE"]

fbm["TPM"] = fbm[columns].sum(axis=1)
dingo["TPM"] = dingo[columns].sum(axis=1)
ortiz["TPM"] = ortiz[columns].sum(axis=1)
zest["TPM"] = zest[columns].sum(axis=1)
pgi["TPM"] = pgi[columns].sum(axis=1)

sorted_fbm = fbm[["VLT", "SUPPLIER GAME", "TPM"]]
sorted_dingo = dingo[["VLT", "SUPPLIER GAME", "TPM"]]
sorted_ortiz = ortiz[["VLT", "SUPPLIER GAME", "TPM"]]
sorted_zest = zest[["VLT", "SUPPLIER GAME", "TPM"]]
sorted_pgi = pgi[["VLT", "SUPPLIER GAME", "TPM"]]

sorted_fbm = sorted_fbm.sort_values(by="TPM", ascending=False).reset_index(drop=True)
sorted_dingo = sorted_dingo.sort_values(by="TPM", ascending=False).reset_index(drop=True)
sorted_ortiz = sorted_ortiz.sort_values(by="TPM", ascending=False).reset_index(drop=True)
sorted_zest = sorted_zest.sort_values(by="TPM", ascending=False).reset_index(drop=True)
sorted_pgi = sorted_pgi.sort_values(by="TPM", ascending=False).reset_index(drop=True)

# SORT EVERY SUPPLIER FROM HIGHEST TO LOWEST BASED FROM THEIR CASH IN
sorted = pd.concat([sorted_fbm, sorted_dingo, sorted_ortiz, sorted_zest, sorted_pgi], axis=1)

#complete_dingo = pd.DataFrame({"DINGO VLT": complete_dingo_vlt})

sorted.columns = ["FBM VLT", "SUPPLIER GAME", "FBM TOTAL CASH IN PER MACHINE",
                  "DINGO VLT", "SUPPLIER GAME", "DINGO TOTAL CASH IN PER MACHINE",
                  "ORTIZ VLT", "SUPPLIER GAME", "ORTIZ TOTAL CASH IN PER MACHINE",
                  "ZEST VLT", "SUPPLIER GAME", "ZEST TOTAL CASH IN PER MACHINE",
                  "PGI VLT", "SUPPLIER GAME", "PGI TOTAL CASH IN PER MACHINE"]

sorted = sorted.drop(index=0)


# TO SAVE
output_path = f"c:\\Users\\Dell\\Desktop\\BRYAN FILES\\BRYAN PROGRAM\\DGRR_Daily_sales_{selected_month}_{selected_year}.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
    sorted.to_excel(writer, sheet_name="SUMMARY", index=False)
    fbm.to_excel(writer, sheet_name="FBM", index=False)
    dingo.to_excel(writer, sheet_name="DINGO", index=False)
    ortiz.to_excel(writer, sheet_name="ORTIZ", index=False)
    zest.to_excel(writer, sheet_name="ZEST", index=False)
    pgi.to_excel(writer, sheet_name="PGI", index=False)



# .PY TO .EXE
#pyinstaller --onefile --windowed --exclude-module matplotlib --icon=PER_MACHINE_LOGO.ico Daily_Sales_Per_Machine.py
#pyinstaller --onefile --exclude-module matplotlib Daily_Sales_Per_Machine.py
