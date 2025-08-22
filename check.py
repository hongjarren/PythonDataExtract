import pandas as pd

main_file = "main.xlsx"

# Read only the BizGroup2024 sheet without assuming a header
df_biz = pd.read_excel(main_file, sheet_name="BizGroup2024")

print("Available columns in BizGroup2024 sheet:")
for col in df_biz.columns:
    print(f"- {col}")
