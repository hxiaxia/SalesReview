import glob
import os
import pandas as pd

files = glob.glob("downloads/*.xlsx")
if files:
    latest = max(files, key=os.path.getctime)
    print("FILE:", latest)
    xl = pd.ExcelFile(latest)
    print("ALL SHEETS:", xl.sheet_names)
    for sheet in xl.sheet_names:
        df = pd.read_excel(latest, sheet_name=sheet)
        if sheet != xl.sheet_names[0]:
            print(f"Sheet {sheet} has {len(df)} rows")
