import pandas as pd
import re
from dateutil import parser
from rapidfuzz import process, fuzz

# =========================
# FILE PATHS
# =========================
ALOHA1_FILE = "Aloha1.xlsx"      # appointments + Client ID + Insurance ID + Appt. Date, Rename Appointment Billing Info file to Aloha1
ALOHA2_FILE = "Aloha2.xlsx"      # Client ID, Client, Status, Rename Clients file to Aloha2
ZOHO_FILE = "Zoho.xlsx"          # Client, Insurance ID, Status
HIRASMUS_FILE = "HiRasmus.xlsx"  # Client, Insurance ID, Status

OUTPUT_FILE = "Master_Client_Reconciliation.xlsx"

# =========================
# HELPERS
# =========================
def parse_date(val):
    try:
        return parser.parse(str(val)).date()
    except Exception:
        return None

def normalize_name(name):
    if pd.isna(name):
        return ""
    name = str(name).lower()
    name = re.sub(r"[^\w\s]", "", name)
    parts = [p.strip() for p in name.split(",")]
    if len(parts) == 2:
        name = f"{parts[1]} {parts[0]}"
    return re.sub(r"\s+", " ", name).strip()

def first_non_null(s):
    s = s.dropna()
    return s.iloc[0] if not s.empty else pd.NA

def fuzzy_match(row, lookup_df, threshold=90):
    # 1) Insurance ID exact match
    if pd.notna(row["Insurance ID"]):
        hit = lookup_df[lookup_df["Insurance ID"] == row["Insurance ID"]]
        if not hit.empty:
            return hit.iloc[0]

    # 2) Fuzzy name fallback
    match = process.extractOne(
        row["norm_name"],
        lookup_df["norm_name"],
        scorer=fuzz.token_sort_ratio
    )
    if match and match[1] >= threshold:
        return lookup_df.loc[lookup_df["norm_name"] == match[0]].iloc[0]

    return None

# =========================
# LOAD FILES
# =========================
aloha1 = pd.read_excel(ALOHA1_FILE)
aloha2 = pd.read_excel(ALOHA2_FILE)
zoho = pd.read_excel(ZOHO_FILE)
hirasmus = pd.read_excel(HIRASMUS_FILE)

# Normalize Client ID column name
aloha2 = aloha2.rename(columns={"Client Id": "Client ID"})

# =========================
# STEP 1: LAST DATE OF SERVICE (ALOHA1)
# =========================
aloha1["Appt. Date"] = aloha1["Appt. Date"].apply(parse_date)

appts = aloha1.dropna(subset=["Insurance ID", "Appt. Date"]).copy()

if "Status" in appts.columns:
    bad = appts["Status"].astype(str).str.lower().str.contains(
        r"cancel|no show|noshow", na=False
    )
    appts = appts[~bad]

last_service = (
    appts
    .groupby("Insurance ID", as_index=False)["Appt. Date"]
    .max()
    .rename(columns={"Appt. Date": "Last Date of Service"})
)

# =========================
# STEP 2: BUILD ALOHA MAIN
# =========================
insurance_map = (
    aloha1[["Client ID", "Insurance ID"]]
    .drop_duplicates(subset=["Client ID"], keep="first")
)

aloha_main = aloha2.merge(insurance_map, on="Client ID", how="left")

aloha_main = (
    aloha_main
    .groupby("Client ID", as_index=False)
    .agg({
        "Client": first_non_null,
        "Insurance ID": first_non_null,
        "Status": first_non_null
    })
    .rename(columns={"Status": "Aloha Status"})
)

aloha_main = aloha_main.merge(last_service, on="Insurance ID", how="left")
aloha_main["norm_name"] = aloha_main["Client"].apply(normalize_name)

# =========================
# STEP 3: PREP ZOHO & HIRASMUS
# =========================
zoho["norm_name"] = zoho["Client"].apply(normalize_name)
hirasmus["norm_name"] = hirasmus["Client"].apply(normalize_name)

# =========================
# STEP 4: ATTACH STATUSES (ID → FUZZY)
# =========================
zoho_status = []
for _, row in aloha_main.iterrows():
    z = fuzzy_match(row, zoho)
    zoho_status.append(z["Status"] if z is not None else pd.NA)
aloha_main["Zoho Status"] = zoho_status

hirasmus_status = []
for _, row in aloha_main.iterrows():
    h = fuzzy_match(row, hirasmus)
    hirasmus_status.append(h["Status"] if h is not None else pd.NA)
aloha_main["HiRasmus Status"] = hirasmus_status

# =========================
# FINAL EXPORT
# =========================
final = aloha_main[[
    "Client",
    "Last Date of Service",
    "Zoho Status",
    "Aloha Status",
    "HiRasmus Status"
]]

final.to_excel(OUTPUT_FILE, index=False)

print("✅ Master reconciliation file created")
print(f"Unique clients exported: {len(final)}")
print(f"Output file: {OUTPUT_FILE}")
