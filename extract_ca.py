import pandas as pd
import numpy as np

def extract_answer_key():
    excel_path = "(TG) AYT ÖZEL DERECE SINAVI (D1226) CA.xlsx"
    df = pd.read_excel(excel_path, header=None)
    
    # Let's find the rows that have the test titles
    # "TÜRK DİLİ VE EDEBİYATI / SOSYAL BİLİMLER-1" is in the file.
    # Usually it's in row 1 or 2.
    
    # We can just iterate through all columns and collect answers.
    # A robust way is: For every cell that contains a number 1..46, if the next cell is '-' and the next is an answer (A-E), record it.
    # We can use a regex to find A-E.
    
    # We will build a dictionary: test_name -> list of answers.
    # The tests are TDE-Sos1 (40), Sos2 (46), Mat (40), Fen (40)
    # We will just sequentially read blocks of 40, 46, 40, 40 questions.
    
    # Find the row containing "1" in the first few columns
    start_row = None
    for idx, row in df.iterrows():
        if row.astype(str).str.contains('1').any():
            if '-' in row.astype(str).values or 'A' in row.astype(str).values or 'D' in row.astype(str).values:
                # This looks like the first row of answers
                start_row = idx
                break
                
    if start_row is None:
        start_row = 2 # fallback
        
    print(f"Answer key seems to start at row {start_row}")
    
    # Find the positions of the test headers
    tests = {
        "TDE-Sos1": "TÜRK DİLİ VE EDEBİYATI",
        "Sos2": "SOSYAL BİLİMLER-2",
        "Mat": "MATEMATİK",
        "Fen": "FEN BİLİMLERİ"
    }
    
    answer_keys = {k: {} for k in tests.keys()}
    
    for test_key, test_name in tests.items():
        # find the cell containing the test name
        found = False
        for row_idx, row in df.iterrows():
            for col_idx, cell in enumerate(row):
                if pd.notna(cell) and test_name in str(cell):
                    start_row = row_idx + 1
                    # Start column for this test
                    start_col = col_idx
                    # scan 15 columns max (5 pairs of Q-A) to capture up to 50 questions
                    for r in range(start_row, start_row + 20): 
                        if r >= len(df): break
                        for c in range(start_col, start_col + 15):
                            if c >= len(df.columns) - 2: break
                            q_val = df.iloc[r, c]
                            dash_val = df.iloc[r, c+1]
                            a_val = df.iloc[r, c+2]
                            
                            try:
                                if pd.notna(q_val) and pd.notna(a_val):
                                    q = int(float(str(q_val).strip()))
                                    a = str(a_val).strip()
                                    if a in ['A', 'B', 'C', 'D', 'E'] and (dash_val == '-' or pd.isna(dash_val)):
                                        answer_keys[test_key][q] = a
                            except ValueError:
                                pass
                    found = True
                    break
            if found: break
            
    # Convert dicts to strings
    for k, v in answer_keys.items():
        if not v:
            print(f"{k} -> Not found!")
            continue
        max_q = max(v.keys())
        # create string with spaces for missing questions
        key_str = "".join([v.get(i, " ") for i in range(1, max_q + 1)])
        print(f"{k} ({len(key_str)}): '{key_str}'")

extract_answer_key()
