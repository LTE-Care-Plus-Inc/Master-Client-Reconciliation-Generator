import streamlit as st
import pandas as pd
import re
from dateutil import parser
from rapidfuzz import process, fuzz
import io

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
# STREAMLIT UI
# =========================
st.set_page_config(page_title="Client Reconciliation", layout="wide")

st.title("📊 Client Status Reconciliation Tool")
st.markdown(
    """
**Purpose of this app**

This app creates a **single, deduplicated master client list** by comparing client records across **Aloha**, **Zoho**, and **HiRasmus**.
"""
)

# =========================
# FILE UPLOADS
# =========================
st.markdown("### Upload Files")

aloha1_file = st.file_uploader(
    "Aloha1: Rename Appointment Billing Info file to Aloha1 for correct input.\n\nUpload Aloha1.xlsx (Appointments / Billing Info)",
    type=["xlsx"]
)

aloha2_file = st.file_uploader(
    "Aloha2: Rename Clients file to Aloha2\n\nUpload Aloha2.xlsx (Client Roster / Status)",
    type=["xlsx"]
)

zoho_file = st.file_uploader(
    "Upload Zoho.xlsx (Zoho Status — Case Dropped / Dropped / Paused)",
    type=["xlsx"]
)

hirasmus_file = st.file_uploader(
    "Upload HiRasmus.xlsx (HiRasmus Client Status)",
    type=["xlsx"]
)

run_btn = st.button("🚀 Run Reconciliation")

# =========================
# MAIN LOGIC
# =========================
if run_btn:
    if not all([aloha1_file, aloha2_file, zoho_file, hirasmus_file]):
        st.error("❌ Please upload all four files before running.")
        st.stop()

    with st.spinner("Processing files…"):
        aloha1 = pd.read_excel(aloha1_file)
        aloha2 = pd.read_excel(aloha2_file)
        zoho = pd.read_excel(zoho_file)
        hirasmus = pd.read_excel(hirasmus_file)

        # Normalize Client ID column
        aloha2 = aloha2.rename(columns={"Client Id": "Client ID"})

        # =========================
        # STEP 1: LAST DATE OF SERVICE
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
        # STEP 4: ATTACH STATUSES
        # =========================
        zoho_status = []
        hirasmus_status = []

        for _, row in aloha_main.iterrows():
            z = fuzzy_match(row, zoho)
            h = fuzzy_match(row, hirasmus)

            zoho_status.append(z["Status"] if z is not None else pd.NA)
            hirasmus_status.append(h["Status"] if h is not None else pd.NA)

        aloha_main["Zoho Status"] = zoho_status
        aloha_main["HiRasmus Status"] = hirasmus_status

        final = aloha_main[[
            "Client",
            "Last Date of Service",
            "Zoho Status",
            "Aloha Status",
            "HiRasmus Status"
        ]]

    st.success("✅ Reconciliation complete!")

    # =========================
    # PREVIEW
    # =========================
    st.subheader("Preview (first 50 rows)")
    st.dataframe(final.head(50), use_container_width=True)

    # =========================
    # DOWNLOAD
    # =========================
    if not final.empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            final.to_excel(writer, index=False, sheet_name="Reconciliation")

        buffer.seek(0)

        st.download_button(
            label="⬇️ Download Excel",
            data=buffer,
            file_name="Master_Client_Reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No data to export.")
