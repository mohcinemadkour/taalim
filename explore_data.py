import pandas as pd

# Read with no header first to find where the actual data starts
df = pd.read_excel('الثانوية التأهيلية صلاح الدين الايوبي_الثالثة إعدادي مسار دولي.xlsx', sheet_name='3APIC-1', header=None)

# Find the row with actual headers
for i, row in df.iterrows():
    if 'اسم' in str(row.values):
        print(f"Header row at index {i}:")
        print(row)
        print("\nData after this row:")
        data_df = pd.read_excel('الثانوية التأهيلية صلاح الدين الايوبي_الثالثة إعدادي مسار دولي.xlsx', sheet_name='3APIC-1', header=i)
        print(data_df.head(10))
        print("\nColumns:", data_df.columns.tolist())
        print("\nData types:")
        print(data_df.dtypes)
        break
