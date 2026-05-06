import pandas as pd
import os
import glob
files = glob.glob("backup/*.xlsx") + glob.glob("downloads/*.xlsx")
if files:
    latest = max(files, key=os.path.getctime)
    print("Latest file:", latest)
    df = pd.read_excel(latest)
    print("Columns:", df.columns.tolist())
    project_col = None
    for col in df.columns:
        if '项目' in str(col):
            project_col = col
            break
    if project_col:
        projects = df[project_col].dropna().unique().tolist()
        print("First 5 projects:", projects[:5])
        print("Total projects:", len(projects))
