import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Wholesale Markup Analytics", layout="wide")
st.title("💰 Wholesale Markup Analytics Tool")

# ==============================
# LOAD FILES
# ==============================
@st.cache_data
def load_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

st.header("Upload Files")

inv_file = st.file_uploader("Invoices", type=["xlsx", "csv"])
front_file = st.file_uploader("Frontline", type=["xlsx", "csv"])
tax_file = st.file_uploader("Taxes", type=["xlsx", "csv"])
store_file = st.file_uploader("Storelist", type=["xlsx", "csv"])

if inv_file and front_file and tax_file and store_file:

    inv = load_file(inv_file)
    front = load_file(front_file)
    tax = load_file(tax_file)
    store = load_file(store_file)

    st.success("Files loaded")

    # ==============================
    # COLUMN SELECTORS
    # ==============================
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        inv_store = st.selectbox("Invoice Store", inv.columns)
        inv_product = st.selectbox("Invoice ProductID", inv.columns)
        inv_cost = st.selectbox("Invoice Cost", inv.columns)

    with col2:
        front_product = st.selectbox("Frontline ProductID", front.columns)
        front_cost = st.selectbox("Frontline Cost", front.columns)
        front_family = st.selectbox("Family", front.columns)
        front_start = st.selectbox("Start Date", front.columns)
        front_end = st.selectbox("End Date", front.columns)

    with col3:
        tax_state = st.selectbox("Tax State", tax.columns)
        tax_value = st.selectbox("Tax Value", tax.columns)

    with col4:
        store_store = st.selectbox("Storelist Store", store.columns)
        store_state = st.selectbox("Storelist State", store.columns)

    if st.button("🚀 Run Analysis"):

        progress = st.progress(0)
        status = st.empty()

        # ==============================
        # 🔥 CLEAN KEYS (CRITICAL FIX)
        # ==============================
        inv[inv_product] = inv[inv_product].astype(str).str.strip()
        front[front_product] = front[front_product].astype(str).str.strip()

        inv[inv_store] = inv[inv_store].astype(str).str.strip()
        store[store_store] = store[store_store].astype(str).str.strip()

        tax[tax_state] = tax[tax_state].astype(str).str.strip()
        store[store_state] = store[store_state].astype(str).str.strip()

        # ==============================
        # STEP 1: ACTIVE FRONTLINE
        # ==============================
        status.text("Filtering active frontline...")
        today = pd.Timestamp.today()

        front[front_start] = pd.to_datetime(front[front_start], errors="coerce")
        front[front_end] = pd.to_datetime(front[front_end], errors="coerce")
        front[front_end] = front[front_end].fillna(pd.Timestamp.max)

        active_front = front[
            (front[front_start] <= today) &
            (front[front_end] >= today)
        ]

        active_front = (
            active_front.sort_values(front_start, ascending=False)
            .drop_duplicates(subset=[front_product])
        )

        progress.progress(20)

        # ==============================
        # STEP 2: STORE → STATE
        # ==============================
        status.text("Mapping store to state...")
        merged = inv.merge(
            store[[store_store, store_state]],
            left_on=inv_store,
            right_on=store_store,
            how="left"
        )

        progress.progress(40)

        # ==============================
        # STEP 3: FRONTLINE + FAMILY
        # ==============================
        status.text("Adding frontline + family...")
        merged = merged.merge(
            active_front[[front_product, front_cost, front_family]],
            left_on=inv_product,
            right_on=front_product,
            how="left"
        )

        progress.progress(55)

        # ==============================
        # STEP 4: TAX
        # ==============================
        status.text("Adding tax...")
        merged = merged.merge(
            tax[[tax_state, tax_value]],
            left_on=store_state,
            right_on=tax_state,
            how="left"
        )

        progress.progress(70)

        # ==============================
        # CLEAN DUPLICATE COLUMNS
        # ==============================
        merged = merged.loc[:, ~merged.columns.duplicated()].copy()

        # ==============================
        # CREATE CLEAN FIELDS
        # ==============================
        merged["State"] = merged[store_state]
        merged["Family"] = merged[front_family]

        merged["Invoice Cost"] = pd.to_numeric(merged[inv_cost], errors="coerce")
        merged["Frontline"] = pd.to_numeric(merged[front_cost], errors="coerce")
        merged["Tax"] = pd.to_numeric(merged[tax_value], errors="coerce")

        # Fix tax %
        merged["Tax"] = merged["Tax"].apply(
            lambda x: x/100 if pd.notna(x) and x > 1 else x
        )

        merged["Frontline"] = merged["Frontline"].fillna(0)
        merged["Tax"] = merged["Tax"].fillna(0)

        # ==============================
        # CALCULATIONS
        # ==============================
        status.text("Calculating metrics...")

        merged["Total Cost"] = merged["Frontline"] * (1 + merged["Tax"])
        merged["Markup"] = merged["Invoice Cost"] - merged["Total Cost"]

        merged["Markup %"] = merged["Markup"] / merged["Total Cost"]
        merged["Markup %"] = merged["Markup %"].replace([float("inf"), -float("inf")], 0)

        progress.progress(85)

        # ==============================
        # FREQUENCY
        # ==============================
        status.text("Calculating frequency...")

        freq_df = merged[["State", "Family", "Invoice Cost"]].dropna()

        freq = (
            freq_df
            .groupby(["State", "Family", "Invoice Cost"])
            .size()
            .reset_index(name="Frequency")
        )

        freq["Top"] = (
            freq.groupby(["State", "Family"])["Frequency"]
            .transform("max") == freq["Frequency"]
        )

        merged = merged.merge(
            freq,
            on=["State", "Family", "Invoice Cost"],
            how="left"
        )

        progress.progress(95)

        # ==============================
        # FINAL OUTPUT
        # ==============================
        final = merged[[
            "State",
            "Family",
            "Invoice Cost",
            "Frontline",
            "Tax",
            "Total Cost",
            "Markup",
            "Markup %",
            "Frequency",
            "Top"
        ]]

        full_output = merged.copy()

        # ==============================
        # EXPORT WITH HIGHLIGHT
        # ==============================
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final.to_excel(writer, sheet_name="Analysis", index=False)
            full_output.to_excel(writer, sheet_name="Full Output", index=False)

        output.seek(0)

        wb = load_workbook(output)
        ws = wb["Analysis"]

        green = PatternFill(start_color="C6EFCE", fill_type="solid")
        top_col = list(final.columns).index("Top") + 1

        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=top_col).value:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = green

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        progress.progress(100)
        status.text("✅ Done!")

        st.download_button(
            "📥 Download Analysis",
            data=final_output,
            file_name=f"markup_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
