import app
import openpyxl
import pandas as pd

def test():
    print("Parsing file...")
    df, totals, wb, sheet = app.parse_excel_file('0903.xlsx')
    if df is None:
        print("DF IS NONE!")
        return
    print(f"Parsed {len(df)} rows. Columns: {df.columns.tolist()}")
    print(f"Totals: {totals}")
    try:
        results, diffs = app.analyze_data(df, totals)
        print("Analyze successful!")
        print(f"Total Diffs: {len(diffs)}")
    except Exception as e:
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test()
