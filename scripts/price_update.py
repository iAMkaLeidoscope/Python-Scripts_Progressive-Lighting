import pandas as pd
import re
import os
from datetime import datetime

# -------------------------------
# CONFIG
# -------------------------------
input_file = r"C:\Users\klaxamana\Files To-Do\VC-Modern_PriceUpdate.xlsx"
sheet_name = "Sheet1"
vendor_notes_file = r"C:\Users\klaxamana\Files To-Do\VendorDataNotes\VendorDataNotes_3.30.26.xlsx"
output_folder = r"C:\Users\klaxamana\Price Updates to Load"

# Get vendor name FIRST
vendor_name_input = os.path.basename(input_file).split("_")[0].strip()

today = datetime.today().strftime('%m.%d.%Y')
output_file = os.path.join(
    output_folder,
    f"{vendor_name_input}_PriceUpdate_{today}.csv"
)

# -------------------------------
# LOAD DATA
# -------------------------------
df = pd.read_excel(input_file, sheet_name=sheet_name, dtype=str)
df.rename(columns={'SKU': 'name'}, inplace=True)

vendor_notes = pd.read_excel(vendor_notes_file, dtype=str)

# Clean vendor notes (VERY IMPORTANT for matching)
vendor_notes = vendor_notes.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

# -------------------------------
# CLEAN NAMES
# -------------------------------
def clean_name(x):
    if pd.isna(x):
        return x
    x = re.sub(r"[\/]", "-", x)
    x = re.sub(r"'", "", x)
    x = re.sub(r" ", "-", x)
    x = re.sub(r"\+", "-", x)
    x = re.sub(r":", "-", x)
    x = re.sub(r"\(", "-", x)
    x = re.sub(r"\)", "-", x)
    x = re.sub(r"--", "-", x)
    x = re.sub(r"\.", "-", x)
    x = re.sub(r"\*", "", x)
    x = re.sub(r"_", "", x)
    return x

df['name'] = df['name'].apply(clean_name)
df = df[df['name'].notna() & (df['name'] != "")]

# -------------------------------
# DETERMINE VENDOR INFO
# -------------------------------
vendor_name_input = os.path.basename(input_file).split("_")[0].strip()

vendor_info = vendor_notes[vendor_notes['Vendor Name'] == vendor_name_input]

if vendor_info.empty:
    raise ValueError(f"Vendor '{vendor_name_input}' not found in master vendor list.")

vendor_info = vendor_info.iloc[0]

print("------ VENDOR MATCH DEBUG ------")
print(vendor_info[['Vendor Name', 'Discount Notes']])
print("--------------------------------")

suffix = vendor_info['Suffix'] if pd.notna(vendor_info['Suffix']) else ""
vendor_number = str(vendor_info['Vendor #']) if pd.notna(vendor_info['Vendor #']) else ""
buyer = vendor_info['Buyer'] if pd.notna(vendor_info['Buyer']) else ""
discount_note = vendor_info['Discount Notes'] if pd.notna(vendor_info['Discount Notes']) else ""
other_notes = vendor_info['Other Notes'] if pd.notna(vendor_info['Other Notes']) else ""

# -------------------------------
# PARSE DISCOUNT
# -------------------------------
discount_pct = 0

if pd.notna(discount_note):
    val = float(discount_note)

    # If value is already decimal (like 0.1), use it directly
    if val < 1:
        discount_pct = val
    else:
        discount_pct = val / 100
# DEBUG (temporary - remove later)
print(f"DEBUG → Raw discount_note: '{discount_note}'")
print(f"DEBUG → Final discount_pct: {discount_pct}")

# -------------------------------
# PARSE CASH / TRADE
# -------------------------------
trade = "Yes"
cash = "Yes"

if "No" in other_notes:
    trade = "No"
    cash = "No"

# -------------------------------
# APPLY TRANSFORMATIONS
# -------------------------------
df['name'] = df['name'] + suffix

df['DealerNet'] = df['DealerNet'].str.replace("$", "", regex=False).astype(float)
df['IMAP'] = df['IMAP'].replace("", "0").astype(float)

# Costs
df['PL_cost'] = round(df['DealerNet'] * (1 - discount_pct), 2)
df['LOL_cost'] = df['PL_cost']

# DN & Clearance
df['DN'] = df['DealerNet']
df['ClearanceCenterPrice'] = round(df['DN'] * 1.5)

# Price
df['price'] = df['IMAP']
df.loc[df['price'] == 0, 'price'] = df['DN'] * 2
df['price'] = df['price'].round(2)

df['wholesale_price'] = df['price']

# Vendor fields
df['vendor_name'] = vendor_number
df['trade'] = trade
df['cash'] = cash
df['buyer'] = buyer

# -------------------------------
# FINALIZE COLUMNS
# -------------------------------
keep_columns = [
    'name', 'PL_cost', 'LOL_cost', 'DN', 'ClearanceCenterPrice',
    'price', 'wholesale_price', 'vendor_name', 'trade', 'cash', 'buyer'
]

df_final = df[keep_columns]

# -------------------------------
# LOGGING
# -------------------------------
print("---------- PROCESS SUMMARY ----------")
print(f"Vendor: {vendor_name_input}")
print(f"Items Processed: {len(df_final)}")
print(f"Discount Applied: {discount_pct*100}%")
print(f"Buyer: {buyer}")
print(f"Trade: {trade}, Cash: {cash}")
print(f"Output: {output_file}")
print("-------------------------------------")

# -------------------------------
# EXPORT CSV
# -------------------------------

df_final.to_csv(output_file, index=False)

print("✅ Price update CSV generated successfully.")

# Open the file automatically
os.startfile(output_file)