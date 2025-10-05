import pandas as pd
from datetime import timedelta

# Define the input Excel file name
input_file = 'data.xlsx'
output_file = 'data_Results.xlsx'

# --- 1. Load data from the single Excel file's sheets ---
print(f"Reading sheets from '{input_file}'...")

try:
    # Read the 'Source' sheet
    df_source = pd.read_excel(input_file, sheet_name='Source')
    # Read the 'month' sheet
    df_month = pd.read_excel(input_file, sheet_name='month')
except FileNotFoundError:
    print(f"ERROR: Could not find the file named '{input_file}'.")
    print("Please ensure MOCK_DATA.xlsx is in the same directory as this script.")
    exit()
except ValueError as e:
    print(f"ERROR: Could not find one of the required sheets ('Source' or 'month') in the Excel file.")
    print(f"Original error: {e}")
    exit()

# --- 2. Calculate the 'd-warn' date (2 days before d-exp) ---

# Convert the 'd-exp' column to datetime objects (handling case-insensitivity by finding the column)
d_exp_col_month = next((col for col in df_month.columns if col.lower() == 'd-exp'), None)
d_exp_col_source = next((col for col in df_source.columns if col.lower() == 'd-exp'), None)

if not d_exp_col_month:
    print("ERROR: Could not find the 'd-exp' column in the 'month' sheet.")
    exit()

if not d_exp_col_source:
    print("ERROR: Could not find the 'd-exp' column in the 'source' sheet.")
    exit()

# Convert d-exp columns to datetime in both dataframes
df_month[d_exp_col_month] = pd.to_datetime(df_month[d_exp_col_month])
df_source[d_exp_col_source] = pd.to_datetime(df_source[d_exp_col_source])

# Calculate the new 'd-warn' date
df_month['d-warn'] = df_month[d_exp_col_month] - timedelta(days=2)

# Format dates back to YYYY-MM-DD string format for consistency
df_month['d-warn'] = df_month['d-warn'].dt.strftime('%Y-%m-%d')
df_month[d_exp_col_month] = df_month[d_exp_col_month].dt.strftime('%Y-%m-%d')
df_source[d_exp_col_source] = df_source[d_exp_col_source].dt.strftime('%Y-%m-%d')


# --- 3. Identify Mismatches (Issues) ---

# Prepare clean versions of the dataframes for comparison
# Drop the empty 'd-warn' column from the month sheet for comparison purposes
df_month_clean = df_month.drop(columns=['d-warn'], errors='ignore')
df_source_clean = df_source.copy()

# Identify common columns for a perfect match check (case-insensitive)
common_cols_source = {col.lower(): col for col in df_source_clean.columns}
common_cols_month = {col.lower(): col for col in df_month_clean.columns}

# Find columns that exist in both sheets
match_cols_lower = list(set(common_cols_source.keys()) & set(common_cols_month.keys()))
# Get the original column names from the Source sheet for the merge
match_cols = [common_cols_source[col] for col in match_cols_lower]


# Perform a full outer merge to identify discrepancies
df_merged = pd.merge(
    df_source_clean,
    df_month_clean,
    on=match_cols,
    how='outer',
    indicator=True,
)

# Mismatched rows are those that are NOT in 'both' files
df_issues = df_merged[df_merged['_merge'] != 'both'].copy()

# Clean up the issues sheet
if not df_issues.empty:
    df_issues = df_issues.drop(columns=['_merge'])
    df_issues = df_issues.fillna('')
    print(f"⚠️ {len(df_issues)} mismatches found. Populating the 'issues' sheet.")
else:
    print("✅ No mismatches found between the 'Source' and 'month' sheets.")


# --- 4. Save the final results to a new Excel file ---

# Select the final column order for the updated month sheet
final_month_cols = [
    'id', 'd-warn', 'first_name', 'last_name', 'email', 'd-exp', 'ok-id'
]
# Adjust column names to match the case found in the original 'month' sheet
month_col_mapping = {col.lower(): col for col in df_month.columns}
final_month_cols_case_sensitive = [month_col_mapping.get(col.lower(), col) for col in final_month_cols]


with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Sheet 1: Keep the original 'Source' sheet
    df_source.to_excel(writer, sheet_name='Source', index=False)
    
    # Sheet 2: The updated 'month' sheet with the calculated d-warn date
    df_month[final_month_cols_case_sensitive].to_excel(
        writer, sheet_name='month (Updated)', index=False
    )
    
    # Sheet 3: The 'issues' sheet with any mismatched rows (if any)
    if not df_issues.empty:
        df_issues.to_excel(writer, sheet_name='issues', index=False)

print(f"\nSuccessfully generated and saved results to: {output_file}")
print("The file contains the following sheets:")
print("1. 'Source': The original source data (preserved)")
print("2. 'month (Updated)': The main data with the 'd-warn' calculated (d-exp - 2 days).")
if not df_issues.empty:
    print("3. 'issues': Rows that did not perfectly match on all fields between the 'Source' and 'month' sheets.")