import pandas as pd  # type: ignore
import sys

with open("result.txt", "w", encoding='utf-8') as f:
    try:
        xl = pd.ExcelFile("MPS2603-1(생산배포용).xlsx")
        f.write("Sheets: " + ", ".join(xl.sheet_names) + "\n")
        df = xl.parse("[생산배포용]")
        f.write("Columns: " + ", ".join(map(str, df.columns.tolist())) + "\n")
        f.write(df.head(10).to_string() + "\n")
    except Exception as e:
        f.write("Error: " + str(e) + "\n")
