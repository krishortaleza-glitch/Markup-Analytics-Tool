import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

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
    # COLUMN SELECTION
    # ==============================
    st.header("Select Columns")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.subheader("Invoices")
        inv_store = st.selectbox("Store", inv.columns)
        inv_product = st.selectbox("Product/UPC", inv.columns)
        inv_cost = st.selectbox("Invoice Cost", inv.columns)
        inv_family = st.selectbox("Family", inv.columns)

    with col2:
        st.subheader("Frontline")
        front_product = st.selectbox("Product/UPC (Frontline)", front.columns)
        front_cost = st.selectbox("Frontline Cost", front.columns)
        front_start = st.selectbox("Start Date", front.columns)
        front_end = st.selectbox("End Date", front.columns)

    with col3:
        st.subheader("Taxes")
        tax_product = st.selectbox("Product/UPC (Tax)", tax.columns)
        tax_cost = st.selectbox("Tax", tax.columns)

    with col4:
        st.subheader("Storelist")
        store_store = st.selectbox("Store (Storelist)", store.columns)
        store_state = st.selectbox("State", store.columns)

    # ==============================
    # PROCESS
    # ==============================
    if st.button("🚀 Run Analysis"):

        progress = st.progress(0)
        status = st.empty()

        # STEP 1: ACTIVE FRONTLINE FILTER
        status.text("Filtering active frontline costs...")

        today = pd.Timestamp.today()

        front[front_start] = pd.to_datetime(front[front_start], errors="coerce")
        front[front_end] = pd.to_datetime(front[front_end], errors="coerce")

        front[front_end] = front[front_end].fillna(pd.Timestamp.max)

        active_front = front[
            (front[front_start] <= today) &
            (front[front_end] >= today)
        ]

        # Keep latest start date per product
        active_front = (
            active_front
            .sort_values(front_start, ascending=False)
            .drop_duplicates(subset=[front_product])
        )

        progress.progress(20)

        # STEP 2: MERGE STORE
        status.text("Merging store-state...")
        merged = inv.merge(
            store[[store_store, store_state]],
            left_on=inv_store,
            right_on=store_store,
            how="left"
        )

        progress.progress(40)

        # STEP 3: MERGE FRONTLINE
        status.text("Adding frontline costs...")
        merged = merged.merge(
            active_front[[front_product, front_cost]],
            left_on=inv_product,
            right_on=front_product,
            how="left"
        )

        progress.progress(55)

        # STEP 4: MERGE TAX
        status.text("Adding tax...")
        merged = merged.merge(
            tax[[tax_product, tax_cost]],
            left_on=inv_product,
            right_on=tax_product,
            how="left"
        )

        progress.progress(70)

        # ==============================
        # CALCULATIONS
        # ==============================
        status.text("Calculating metrics...")

        merged.rename(columns={
            inv_cost: "Invoice Cost",
            front_cost: "Frontline",
            tax_cost: "Tax",
            inv_family: "Family",
            store_state: "State"
        }, inplace=True)

        merged["Frontline"] = merged["Frontline"].fillna(0)
        merged["Tax"] = merged["Tax"].fillna(0)

        merged["Total Cost"] = merged["Frontline"] + merged["Tax"]
        merged["Markup"] = merged["Invoice Cost"] - merged["Total Cost"]

        merged["Markup %"] = (
            merged["Markup"] / merged["Total Cost"]
        ).replace([float("inf"), -float("inf")], 0)

        progress.progress(85)

        # ==============================
        # FREQUENCY
        # ==============================
        status.text("Calculating frequency...")

        freq = (
            merged
            .groupby(["State", "Family", "Invoice Cost"])
            .size()
            .reset_index(name="Frequency")
        )

        merged = merged.merge(
            freq,
            on=["State", "Family", "Invoice Cost"],
            how="left"
        )

        # Most frequent price
        top_prices = (
            freq.sort_values("Frequency", ascending=False)
            .drop_duplicates(["State", "Family"])
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
            "Frequency"
        ]]

        # ==============================
        # EXPORT
        # ==============================
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final.to_excel(writer, sheet_name="Analysis", index=False)
            top_prices.to_excel(writer, sheet_name="Most Frequent Price", index=False)

        output.seek(0)

        progress.progress(100)
        status.text("✅ Done!")

        st.download_button(
            "📥 Download Analysis",
            data=output,
            file_name=f"markup_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )