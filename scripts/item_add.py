# ==========================================================
# ITEM ADD AUTOMATION SCRIPT
# (Aligned with WORKING Price Update Logic)
# ==========================================================

import pandas as pd
import os
import re
from datetime import datetime

# ----------------------------------------------------------
# CONFIG
# ----------------------------------------------------------

input_file = r"C:\Users\klaxamana\Files To-Do\Varaluz_ToAdd.xlsx"
sheet_name = "Sheet1"

vendor_notes_file = r"C:\Users\klaxamana\Files To-Do\VendorDataNotes\VendorDataNotes_3.30.26.xlsx"

output_folder = r"C:\Users\klaxamana\New Items to Load"

# ----------------------------------------------------------
# HELPER FUNCTIONS
# ----------------------------------------------------------

def clean_columns(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )
    return df


def clean_name(x):
    if pd.isna(x):
        return x

    x = str(x)
    replacements = {
        "/": "-",
        ".": "-",
        "'": "",
        " ": "-",
        "+": "-",
        ":": "-",
        "(": "-",
        ")": "-",
        "_": "-"
    }

    for old, new in replacements.items():
        x = x.replace(old, new)

    x = re.sub(r"-+", "-", x)
    return x


# Flexible column finder (prevents KeyErrors)
def find_column(df, possible):
    for name in possible:
        if name in df.columns:
            return name
    raise KeyError(f"Missing column. Tried: {possible}")


# ----------------------------------------------------------
# LOAD DATA
# ----------------------------------------------------------

vendor_name_input = os.path.basename(input_file).split("_")[0].strip()
print(f"Processing Vendor: {vendor_name_input}")

df = pd.read_excel(input_file, sheet_name=sheet_name, dtype=str)
vendor_notes = pd.read_excel(vendor_notes_file, dtype=str)

df = clean_columns(df)
vendor_notes = clean_columns(vendor_notes)

# ----------------------------------------------------------
# DETERMINE VENDOR INFO (MATCH WORKING SCRIPT)
# ----------------------------------------------------------

vendor_info = vendor_notes[
    vendor_notes["vendor_name"] == vendor_name_input
]

if vendor_info.empty:
    raise ValueError(
        f"Vendor '{vendor_name_input}' not found in VendorDataNotes."
    )

vendor = vendor_info.iloc[0]

print("------ VENDOR MATCH DEBUG ------")
print(vendor[["vendor_name", "discount_notes"]])
print("--------------------------------")

suffix = vendor.get("suffix", "") or ""
vendor_number = str(vendor.get("vendor_#", "") or "")
buyer = vendor.get("buyer", "") or ""
discount_note = vendor.get("discount_notes", "") or ""
other_notes = vendor.get("other_notes", "") or ""

# ----------------------------------------------------------
# PARSE DISCOUNT (same logic as working script)
# ----------------------------------------------------------

discount_pct = 0

if discount_note != "":
    val = float(discount_note)
    discount_pct = val if val < 1 else val / 100

print(f"Discount Applied: {discount_pct*100}%")

# ----------------------------------------------------------
# COLUMN AUTO-DETECTION
# ----------------------------------------------------------

col_sku = find_column(df, ["sku", "item", "model"])
col_desc = find_column(df, ["productdescription", "product_description", "description"])
col_upc = find_column(df, ["upc", "upccode"])
col_dn = find_column(df, ["dealernet", "dealer_net", "dn"])
col_imap = find_column(df, ["imap", "retail", "price"])

# Optional columns (won't crash if missing)
def optional(col):
    return col if col in df.columns else None

col_active = optional("active")

# ----------------------------------------------------------
# DROP INACTIVE (STATA BEHAVIOR)
# ----------------------------------------------------------

if col_active:
    df = df[df[col_active].notna()]

# ----------------------------------------------------------
# NAME CREATION
# ----------------------------------------------------------

df["name"] = (df[col_sku] + suffix).apply(clean_name)

# ----------------------------------------------------------
# DESCRIPTION FIXES
# ----------------------------------------------------------

df["displayname"] = df[col_desc].str.strip()
df["upccode"] = df[col_upc]

# ----------------------------------------------------------
# PRICING LOGIC
# ----------------------------------------------------------

df["dn"] = pd.to_numeric(df[col_dn], errors="coerce")
df["retail"] = pd.to_numeric(df[col_imap], errors="coerce")

df.loc[df["retail"].isna(), "retail"] = df["dn"] * 2

df["online"] = df["retail"]
df["wholesale"] = df["retail"]

df["cost"] = round(df["dn"] * (1 - discount_pct), 2)

df["vendor1_purchaseprice"] = df["cost"]
df["lol_cost"] = df["cost"]
df["pl_cost"] = df["cost"]

df["clearancecenterprice"] = (df["dn"] * 1.5).round().astype(int)

# ----------------------------------------------------------
# STATIC FIELDS
# ----------------------------------------------------------

trade = "Yes"
cash = "Yes"

if isinstance(other_notes, str) and "no" in other_notes.lower():
    trade = "No"
    cash = "No"

df["vendor1_name"] = vendor_number
df["buyer"] = buyer
df["trade"] = trade
df["cash"] = cash
df["sets"] = "1"

# ----------------------------------------------------------
# REMOVE INVALID ROWS
# ----------------------------------------------------------

df = df.dropna(subset=["dn", "retail"])

# ----------------------------------------------------------
# EXPORT
# ----------------------------------------------------------

# ----------------------------------------------------------
# FINALIZE OUTPUT (MATCHES STATA EXPORT EXACTLY)
# ----------------------------------------------------------

df_final = pd.DataFrame()

df_final["Name"] = df["name"]
df_final["UPCCode"] = df["upccode"]
df_final["displayname"] = df["displayname"]
df_final["vendorname"] = df[col_sku]

# dimensions
df_final["item_height"] = df.get("height")
df_final["item_length"] = df.get("length")
df_final["item_width"] = df.get("width_/_diameter")
df_final["item_depth"] = df.get("extension")

# packaging
df_final["pack_height"] = df.get("carton_1_height")
df_final["pack_depth"] = df.get("carton_1_width")
df_final["pack_length"] = df.get("carton_1_length")

# weights
df_final["weight"] = df.get("net_weight")
df_final["pack_weight"] = df.get("gross_weight")

# attributes
df_final["finish"] = df.get("finish").str.upper()
df_final["glass1"] = df.get("glass")

# media/location
df_final["item_image_url"] = df.get("imageurl")
df_final["wet_damp_location"] = df.get("wet_dry")

# vendor info
df_final["vendor1_name"] = df["vendor1_name"]
df_final["sets"] = df["sets"]

# pricing
df_final["DN"] = df["dn"]
df_final["Retail"] = df["retail"]
df_final["Online"] = df["online"]
df_final["Wholesale"] = df["wholesale"]
df_final["Cost"] = df["cost"]
df_final["ClearanceCenterPrice"] = df["clearancecenterprice"]

df_final["vendor1_purchaseprice"] = df["vendor1_purchaseprice"]
df_final["LOL_cost"] = df["lol_cost"]
df_final["PL_cost"] = df["pl_cost"]

# static fields
df_final["cash"] = df["cash"]
df_final["trade"] = df["trade"]
df_final["buyer"] = df["buyer"]

# ----------------------------------------------------------
# EXPORT CSV
# ----------------------------------------------------------

today = datetime.today().strftime("%m.%d.%Y")

output_file = os.path.join(
    output_folder,
    f"New_{vendor_name_input}_{today}.csv"
)

df_final.to_csv(output_file, index=False)

print("---------- PROCESS SUMMARY ----------")
print(f"Vendor: {vendor_name_input}")
print(f"Items Exported: {len(df_final)}")
print(f"Output: {output_file}")
print("-------------------------------------")

print("✅ Item Add CSV generated successfully.")

os.startfile(output_file)