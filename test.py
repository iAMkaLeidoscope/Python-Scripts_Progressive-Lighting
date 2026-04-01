import pandas as pd
import re
import os
from datetime import datetime

# -------------------------------
# CONFIG
# -------------------------------
input_file = r"C:\Users\klaxamana\Files To-Do\Craftmade_PriceUpdate.xlsx"
sheet_name = "Sheet1"
vendor_notes_file = r"C:\Users\klaxamana\Files To-Do\VendorDataNotes_3.30.26.xlsx"
output_folder = r"C:\Users\klaxamana\Price Updates to Load"

today = datetime.today().strftime('%Y.%m.%d')
output_file = os.path.join(output_folder, f"PriceUpdate_{today}.csv")

# -------------------------------
# LOAD DATA
# -------------------------------
df = pd.read_excel(input_file, sheet_name=sheet_name, dtype=str)
df.rename(columns={'SKU': 'name'}, inplace=True)

vendor_notes = pd.read_excel(vendor_notes_file, dtype=str)

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
# Extract vendor name from input filename
vendor_name_input = os.path.basename(input_file).split("_")[0]

vendor_info = vendor_notes[vendor_notes['Vendor Name'] == vendor_name_input].iloc[0]

suffix = vendor_info['Suffix'] if pd.notna(vendor_info['Suffix']) else ""
vendor_number = vendor_info['Vendor #'] if pd.notna(vendor_info['Vendor #']) else ""
buyer = vendor_info['Buyer'] if pd.notna(vendor_info['Buyer']) else ""
discount_note = vendor_info['Discount Notes'] if pd.notna(vendor_info['Discount Notes']) else ""
other_notes = vendor_info['Other Notes'] if pd.notna(vendor_info['Other Notes']) else ""

# -------------------------------
# APPLY VENDOR LOGIC
# -------------------------------
df['name'] = df['name'] + suffix

df['DealerNet'] = df['DealerNet'].str.replace("$","", regex=False).astype(float)
df['IMAP'] = df['IMAP'].replace("", "0").astype(float)

# Calculate discount if applicable
discount_pct = 0
if discount_note:
    match = re.search(r'(\d+)%', discount_note)
    if match:
        discount_pct = float(match.group(1))/100

df['PL_cost'] = round(df['DealerNet'] * (1 - discount_pct), 2)
df['LOL_cost'] = df['PL_cost']

df['DN'] = df['DealerNet']
df['ClearanceCenterPrice'] = round(df['DN'] * 1.5, 1)

# Apply special pricing rules from "Other Notes"
df['price'] = df['IMAP']

if other_notes:
    if "IMAP=DN*2.2" in other_notes:
        df['price'] = df['DN'] * 2.2
    elif "IMAP=DN*1.8" in other_notes:
        df['price'] = df['DN'] * 1.8
    # You can add more rules here based on your notes

df.loc[df['price'] == 0, 'price'] = df['DN'] * 2
df['price'] = df['price'].round(2)
df['wholesale_price'] = df['price']

df['vendor_name'] = vendor_number
df['trade'] = "Yes" if "No" not in other_notes else "No"
df['cash'] = "Yes" if "No" not in other_notes else "No"
df['buyer'] = buyer

# -------------------------------
# FINALIZE COLUMNS
# -------------------------------
keep_columns = ['name', 'PL_cost', 'LOL_cost', 'DN', 'ClearanceCenterPrice',
                'price', 'wholesale_price', 'vendor_name', 'trade', 'cash', 'buyer']

df_final = df[keep_columns]

# -------------------------------
# LOGGING
# -------------------------------
print(f"Vendor processed: {vendor_name_input}")
print(f"Total items: {len(df_final)}")
print(f"Discount applied: {discount_pct*100}%")
print(f"Buyer: {buyer}")
print(f"Trade: {df['trade'].iloc[0]}, Cash: {df['cash'].iloc[0]}")
if other_notes:
    print(f"Special notes applied: {other_notes}")
print(f"Output file: {output_file}")

# -------------------------------
# EXPORT CSV
# -------------------------------
df_final.to_csv(output_file, index=False)
print("Price update CSV generated successfully.")