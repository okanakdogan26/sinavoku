import pandas as pd
excel_path = "(TG) AYT ÖZEL DERECE SINAVI (D1226) CA.xlsx"
try:
    df = pd.read_excel(excel_path)
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    print("Columns:", df.columns.tolist())
    print("First 20 rows:")
    print(df.head(20))
    print(df.head())
except Exception as e:
    print(e)
