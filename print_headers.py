import pandas as pd
import os

files = ['Bookings.xlsx', 'Staff.xlsx', 'services.xls']
with open('headers.txt', 'w', encoding='utf-8') as f_out:
    for f in files:
        if os.path.exists(f):
            try:
                df = pd.read_excel(f, nrows=0)
                f_out.write(f"--- {f} ---\n")
                f_out.write(str(list(df.columns)) + "\n")
            except Exception as e:
                f_out.write(f"Error reading {f}: {e}\n")
        else:
            f_out.write(f"File {f} not found\n")
