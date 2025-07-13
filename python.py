# (Excel Cleaner and Summary Generator)
# (It will Clean and Summarize our Excel data automatically using Python)

import pandas as pd

# Load the Excel file from Sheet3
df = pd.read_excel(r"data.xlsx", sheet_name='Sheet3')

# Remove completely empty rows
df.dropna(how='all', inplace=True)

# Strip column names to remove hidden spaces
df.columns = df.columns.str.strip()

# Convert 'Date' column to datetime (handling invalid entries)
if 'Date' in df.columns:
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df=df.dropna(subset=['Date'])
    df['Date'] = df['Date'].dt.strftime('%d-%m-%Y')

# Generate summary if possible
if 'Category' in df.columns and 'Amount' in df.columns:
    summary = df.groupby('Category')['Amount'].sum().reset_index()
else:
    summary = pd.DataFrame({"Note": ["No Category or Amount column found"]})

# Save both cleaned data and summary to a new Excel file
with pd.ExcelWriter("cleaned_output.xlsx", engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Cleaned Data')
    summary.to_excel(writer, index=False, sheet_name='Summary')

print("Cleaning and summary completed. Saved as 'cleaned_output.xlsx'")
