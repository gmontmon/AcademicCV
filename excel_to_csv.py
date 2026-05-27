import pandas as pd

EXCEL_FILE = "Academic_CV_Structure_Template.xlsx"

print("\nLoading Excel workbook...\n")

excel_file = pd.ExcelFile(EXCEL_FILE)
sheet_names = excel_file.sheet_names

print("Sheets found:")
for sheet in sheet_names:
    print(" -", sheet)

for sheet_name in sheet_names:
    print(f"\nProcessing sheet: {sheet_name}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    df = df.dropna(how="all")
    df = df.fillna("")

    csv_filename = f"{sheet_name}.csv"

    df.to_csv(
        csv_filename,
        index=False,
        encoding="utf-8"
    )

    print(f"CSV created: {csv_filename}")

print("\n================================================")
print("ALL CSV FILES GENERATED SUCCESSFULLY")
print("================================================")
print("\nCSV files saved in the root folder.")
print("\nDone.\n")