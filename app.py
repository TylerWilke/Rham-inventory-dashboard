import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import re
from pathlib import Path

st.set_page_config(page_title="Inventory Sales Analysis", layout="wide", page_icon="📦")

# ── Rham Equipment brand colors ───────────────────────────────────────────────
NAVY       = "#1c3d6b"   # primary brand navy
NAVY_LIGHT = "#2a5298"   # lighter navy for secondary bars
SLATE      = "#4a6fa5"   # muted blue for supporting elements
STEEL      = "#6b93c4"   # light steel blue for low-GP highlights

st.markdown("""
<style>
    /* Top header bar */
    [data-testid="stAppViewContainer"] > .main > div:first-child {
        padding-top: 0;
    }
    .rham-header {
        background-color: #1c3d6b;
        padding: 12px 24px;
        margin: -1rem -1rem 1rem -1rem;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .rham-header h1 {
        color: white !important;
        font-size: 1.4rem !important;
        margin: 0 !important;
        padding: 0 !important;
        font-weight: 700;
        letter-spacing: 1px;
    }
    .rham-header span {
        color: #a8c4e0;
        font-size: 0.9rem;
        font-weight: 400;
    }
    /* Metric cards */
    [data-testid="metric-container"] {
        background-color: #e8eef5;
        border-left: 4px solid #1c3d6b;
        padding: 12px 16px;
        border-radius: 4px;
    }
    /* Sidebar title */
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1 {
        color: #1c3d6b;
    }
    /* Tab active color handled by primaryColor in config.toml */
</style>
""", unsafe_allow_html=True)

CATEGORY_LABELS = {
    1: "1 – Buy Outs",
    2: "2 – Value Add",
    3: "3 – Manufacturing",
    "N/A": "N/A",
}

def fmt_zar(value):
    return f"R {value:,.2f}".replace(",", " ")

def fmt_pct(value):
    return f"{value:.2f}%"

def fmt_num(value):
    return f"{value:,.2f}".replace(",", " ")

GROUP_LIST_PATH = Path(__file__).parent / "Group List.xls"


@st.cache_data(show_spinner="Loading Group List…")
def load_group_list():
    if not GROUP_LIST_PATH.exists():
        return pd.DataFrame(columns=["Group", "Notes", "Description", "Type"])
    raw = pd.read_excel(GROUP_LIST_PATH, header=None)
    df = raw.iloc[1:].copy()
    df = df.iloc[:, :4]
    df.columns = ["Group", "Notes", "Description", "Type"]
    df["Group"] = df["Group"].astype(str).str.strip()
    df["Type"] = pd.to_numeric(df["Type"], errors="coerce")
    df["Group_Category"] = df["Type"].apply(
        lambda x: int(x) if pd.notna(x) and x in (1, 2, 3) else "N/A"
    )
    return df.reset_index(drop=True)


@st.cache_data(show_spinner="Parsing sales file…")
def parse_sales_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None)

    # Scan up to first 60 rows for a row that contains "Item Code" in any column
    header_row = None
    item_code_col = None
    for i in range(min(60, len(raw))):
        for j, val in enumerate(raw.iloc[i]):
            if str(val).strip() == "Item Code":
                header_row = i
                item_code_col = j
                break
        if header_row is not None:
            break
    if header_row is None:
        raise ValueError("Could not find 'Item Code' header row in the uploaded file.")

    # Build dynamic column map by scanning the header row for known keywords
    hrow = raw.iloc[header_row]
    col_map = {}
    for j, val in enumerate(hrow):
        s = str(val).strip() if pd.notna(val) else ""
        if s == "Item Code":
            col_map["Item Code"] = j
        elif "Description" in s:
            col_map.setdefault("Item Description", j)
        elif "Markup" in s:
            col_map["Markup %"] = j
        elif "Profit %" in s or "GP %" in s:
            col_map["Gross Profit %"] = j
        elif "Profit" in s:
            col_map.setdefault("Gross Profit", j)
        elif s in ("Group", "Group Code"):
            col_map["Group"] = j
        elif "Quantity" in s or s in ("Qty", "Units"):
            col_map["Quantity"] = j
        elif s in ("Amount", "Sales Amount", "Revenue"):
            col_map["Amount"] = j
        elif "Cost" in s and "%" not in s:
            col_map.setdefault("Cost", j)
        elif "Date" in s:
            col_map.setdefault("Date", j)

    # Fallback: detect Date column by scanning data rows for Timestamp values
    if "Date" not in col_map:
        mapped_cols = set(col_map.values())
        for probe_idx in range(header_row + 1, min(header_row + 30, len(raw))):
            probe = raw.iloc[probe_idx]
            for j, v in enumerate(probe):
                if isinstance(v, pd.Timestamp) and j not in mapped_cols:
                    col_map["Date"] = j
                    break
            if "Date" in col_map:
                break

    # Fallback: detect Group column — first column with short strings between
    # Item Description and Date that isn't already mapped
    if "Group" not in col_map:
        mapped_cols = set(col_map.values())
        desc_col = col_map.get("Item Description", item_code_col)
        date_col = col_map.get("Date", len(hrow))
        from collections import Counter
        cand_counter = Counter()
        for probe_idx in range(header_row + 1, min(header_row + 30, len(raw))):
            probe = raw.iloc[probe_idx]
            cell0 = str(probe.iloc[0]) if pd.notna(probe.iloc[0]) else ""
            if cell0.startswith("Customer:"):
                continue
            for j in range(desc_col + 1, date_col):
                v = probe.iloc[j]
                if isinstance(v, str) and 2 <= len(v.strip()) <= 25 and j not in mapped_cols:
                    cand_counter[j] += 1
                    break
        if cand_counter:
            col_map["Group"] = cand_counter.most_common(1)[0][0]

    required = ["Item Code", "Item Description", "Amount", "Gross Profit", "Gross Profit %"]
    missing = [c for c in required if c not in col_map]
    if missing:
        raise ValueError(f"Could not detect columns: {missing}. Header row: {list(hrow)}")

    # Extract year from filename as fallback (e.g. "2022.xls", "2026 - June.xls")
    yr_match = re.search(r"\b(20\d{2})\b", filename)
    fallback_year = int(yr_match.group(1)) if yr_match else None

    def _get(row, key, default=None):
        c = col_map.get(key)
        if c is None:
            return default
        v = row.iloc[c]
        return v if pd.notna(v) else default

    records = []
    current_customer = None
    customer_id = None

    for idx in range(header_row + 1, len(raw)):
        row = raw.iloc[idx]

        # Customer header rows always appear in col 0
        cell0 = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
        if cell0.startswith("Customer:"):
            m = re.match(r"Customer:\s+(\S+)\s+\((.+)\)", cell0)
            if m:
                customer_id, current_customer = m.group(1).strip(), m.group(2).strip()
            else:
                current_customer = cell0.replace("Customer:", "").strip()
                customer_id = current_customer
            continue

        # Item Code is wherever the header said
        item_val = str(row.iloc[item_code_col]) if pd.notna(row.iloc[item_code_col]) else ""
        if not item_val or item_val in ("nan", "NaT", "None"):
            continue
        if pd.isna(row.iloc[col_map["Item Description"]]) and pd.isna(row.iloc[col_map["Amount"]]):
            continue

        try:
            records.append({
                "Customer ID": customer_id,
                "Customer": current_customer,
                "Item Code": item_val,
                "Item Description": str(_get(row, "Item Description", "")).strip(),
                "Date": pd.to_datetime(_get(row, "Date"), errors="coerce"),
                "Group": str(_get(row, "Group", "")).strip(),
                "Quantity": pd.to_numeric(_get(row, "Quantity"), errors="coerce"),
                "Amount": round(pd.to_numeric(_get(row, "Amount"), errors="coerce"), 2),
                "Cost": round(pd.to_numeric(_get(row, "Cost"), errors="coerce"), 2),
                "Gross Profit": round(pd.to_numeric(_get(row, "Gross Profit"), errors="coerce"), 2),
                "Gross Profit %": round(pd.to_numeric(_get(row, "Gross Profit %"), errors="coerce"), 2),
                "Markup %": round(pd.to_numeric(_get(row, "Markup %"), errors="coerce"), 2),
            })
        except Exception:
            continue

    df = pd.DataFrame(records)
    df = df.dropna(subset=["Amount"])
    df["Year"] = df["Date"].dt.year
    if fallback_year:
        df["Year"] = df["Year"].fillna(fallback_year)
    return df


def merge_group_list(df: pd.DataFrame, group_list: pd.DataFrame) -> pd.DataFrame:
    gl = group_list[["Group", "Description", "Group_Category"]].copy()
    merged = df.merge(gl, on="Group", how="left")
    merged["Group_Category"] = merged["Group_Category"].fillna("N/A")
    return merged


NUM_COLS = ["Amount", "Cost", "Gross Profit", "Profit Lost", "Target GP",
            "Sales", "GP", "Sales (R)", "Actual GP (R)", "Target GP (R)", "Profit Lost (R)"]
PCT_COLS = ["Gross Profit %", "Markup %", "Avg GP %", "Min GP %", "Max GP %",
            "GP Variance (pp)", "Avg_GP_Pct", "GP_Pct"]

def _space_fmt(v):
    try:
        return f"{float(v):,.2f}".replace(",", " ")
    except Exception:
        return v

def apply_gp_style(df_display: pd.DataFrame, threshold: float):
    def row_style(row):
        color = "background-color: #c8d8ee" if row["Gross Profit %"] < threshold else ""
        return [color] * len(row)
    fmt = {c: _space_fmt for c in df_display.columns if c in NUM_COLS + PCT_COLS}
    return df_display.style.apply(row_style, axis=1).format(fmt)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📦 Inventory Analysis")
    st.markdown("---")

    uploaded_files = st.file_uploader(
        "Upload Inventory Sales Files (.xls / .xlsx)",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        help="Upload one file per year — 2022, 2023, 2024, 2025, 2026",
    )

    if st.button("🔄 Reset", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.markdown("**Filters**")

    if uploaded_files:
        try:
            all_dfs = []
            for f in uploaded_files:
                parsed = parse_sales_file(f.getvalue(), f.name)
                all_dfs.append((f.name, parsed))

            raw_df = pd.concat([d for _, d in all_dfs], ignore_index=True)
            group_list = load_group_list()
            full_df = merge_group_list(raw_df, group_list)

            # Show a summary line per uploaded file
            for fname, d in sorted(all_dfs, key=lambda x: x[0]):
                yr_vals = d["Year"].dropna()
                yr_label = int(yr_vals.iloc[0]) if len(yr_vals) > 0 else "?"
                st.caption(f"✅ {yr_label} — {len(d):,} rows")

            customers = sorted(full_df["Customer"].dropna().unique())
            years = sorted(full_df["Year"].dropna().unique().astype(int))
            all_categories = [1, 2, 3, "N/A"]
            default_categories = [1, 2, 3]

            sel_customers = st.multiselect("Customers", customers, default=customers)
            sel_categories = st.multiselect(
                "Group Category", all_categories, default=default_categories,
                format_func=lambda x: CATEGORY_LABELS.get(x, str(x)),
            )
            sel_years = st.multiselect("Years", years, default=years)
            gp_threshold = st.slider("Gross Profit % Threshold", 0, 100, 50, step=1)
            exclude_negative = st.checkbox("Exclude negative Gross Profit", value=False)

            st.markdown("---")
            st.markdown("**Group List**")
            st.download_button(
                "⬇️ Download Group List",
                data=to_excel_bytes(group_list[["Group", "Description", "Type", "Group_Category"]]),
                file_name="Group List Review.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Download the Group List to review and correct Type classifications.",
            )

            df = full_df.copy()
            if sel_customers:
                df = df[df["Customer"].isin(sel_customers)]
            if sel_categories:
                df = df[df["Group_Category"].isin(sel_categories)]
            if sel_years:
                df = df[df["Year"].isin(sel_years)]
            # Save filtered df before negative exclusion for the Negative Amounts tab
            df_with_negatives = df.copy()
            if exclude_negative:
                df = df[df["Gross Profit"] >= 0]

        except Exception as e:
            st.error(f"Error parsing file: {e}")
            df = None
            full_df = None
            df_with_negatives = None
    else:
        df = None
        full_df = None
        df_with_negatives = None
        gp_threshold = 50
        exclude_negative = False

    st.markdown("---")
    st.caption("Group Categories:\n- **1** Buy Outs\n- **2** Value Add\n- **3** Manufacturing\n- **N/A** Unclassified (opt-in)")

# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="rham-header">
    <h1>RHAM EQUIPMENT &nbsp;<span>|&nbsp; Inventory Sales — Gross Profit Dashboard</span></h1>
</div>
""", unsafe_allow_html=True)

if not uploaded_files:
    st.info("👈 Upload one or more **Inventory Sales files** in the sidebar to get started.")
    st.markdown("""
    **Upload one file per year** (e.g. `2022.xls`, `2023.xls`, …, `2026 - June.xls`).
    The app will combine all years into a single dashboard automatically.

    **Expected format:** Inventory Sales Analysis report with customer sections,
    containing columns: Item Code, Item Description, Date, Group, Quantity, Amount, Cost, Gross Profit, Gross Profit %, Markup %.

    **Group List** is loaded automatically from `Group List.xls` placed beside this app.
    """)
    st.stop()

if df is None or len(df) == 0:
    st.warning("No data matches the current filters.")
    st.stop()

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📊 Overview",
    "⚠️ Low GP Analysis",
    "💸 Profit Lost",
    "👤 GP by Customer",
    "🔀 Item GP by Customer",
    "📅 By Year",
    "📋 Raw Data",
    "➖ Negative Amounts",
])

# ─── Tab 1: Overview ──────────────────────────────────────────────────────────
with tab1:
    total_sales = df["Amount"].sum()
    total_gp = df["Gross Profit"].sum()
    low_gp_count = (df["Gross Profit %"] < gp_threshold).sum()
    low_gp_pct = low_gp_count / len(df) * 100 if len(df) else 0
    overall_gp_pct = (total_gp / total_sales * 100) if total_sales else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Sales", fmt_zar(total_sales))
    c2.metric("Total Gross Profit", fmt_zar(total_gp))
    c3.metric("Overall Gross Profit %", fmt_pct(overall_gp_pct))
    c4.metric(f"Items Below {gp_threshold}% GP", f"{low_gp_count:,}".replace(",", " ") + f" ({low_gp_pct:.2f}%)")

    st.markdown("---")

    col_left, col_right = st.columns(2)
    with col_left:
        grp_summary = df.groupby("Group_Category").agg(
            Sales=("Amount", "sum"),
            GP=("Gross Profit", "sum"),
        ).reset_index()
        grp_summary["Gross Profit %"] = (grp_summary["GP"] / grp_summary["Sales"] * 100).round(2)
        grp_summary["Category"] = grp_summary["Group_Category"].map(CATEGORY_LABELS).fillna("N/A")
        fig = px.bar(grp_summary, x="Category", y=["Sales", "GP"],
                     barmode="group", title="Sales & Gross Profit by Group Category",
                     color_discrete_map={"Sales": NAVY, "GP": NAVY_LIGHT},  # noqa
                     labels={"value": "Amount (R)", "variable": ""})
        fig.update_layout(yaxis_title="Amount (R)", legend_title="")
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        cust_summary = df.groupby("Customer").agg(Sales=("Amount", "sum")).nlargest(10, "Sales").reset_index()
        fig2 = px.bar(cust_summary, x="Sales", y="Customer", orientation="h",
                      title="Top 10 Customers by Sales", color_discrete_sequence=[NAVY])
        fig2.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Sales (R)")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown(f"#### All Data  —  blue rows = Gross Profit % below {gp_threshold}%")
    display_cols = ["Customer", "Item Code", "Item Description", "Date", "Group",
                    "Group_Category", "Quantity", "Amount", "Cost", "Gross Profit", "Gross Profit %", "Markup %"]
    disp = df[display_cols].copy()
    styled = apply_gp_style(disp, gp_threshold)
    st.dataframe(styled, use_container_width=True, height=400)

# ─── Tab 2: Low Gross Profit Analysis ────────────────────────────────────────
with tab2:
    low = df[df["Gross Profit %"] < gp_threshold].copy()

    st.subheader(f"Items Below {gp_threshold}% Gross Profit  ({len(low):,} items)")

    if low.empty:
        st.success("No items below the Gross Profit threshold.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Sales (Low GP Items)", fmt_zar(low["Amount"].sum()))
        c2.metric("Gross Profit (Low GP Items)", fmt_zar(low["Gross Profit"].sum()))
        c3.metric("Avg Gross Profit %", fmt_pct(low["Gross Profit %"].mean()))

        st.markdown("#### By Group Category")
        cat_summary = low.groupby("Group_Category").agg(
            Items=("Item Code", "count"),
            Sales=("Amount", "sum"),
            GP=("Gross Profit", "sum"),
        ).reset_index()
        cat_summary["Avg Gross Profit %"] = (cat_summary["GP"] / cat_summary["Sales"] * 100).round(2)
        cat_summary["Category"] = cat_summary["Group_Category"].map(CATEGORY_LABELS).fillna("N/A")
        st.dataframe(cat_summary[["Category", "Items", "Sales", "GP", "Avg Gross Profit %"]],
                     use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(cat_summary, x="Category", y="Items",
                         title="Low Gross Profit Item Count by Category",
                         color_discrete_sequence=[STEEL])
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            top_low = low.nsmallest(15, "Gross Profit %")[["Item Code", "Item Description", "Amount", "Gross Profit %"]]
            fig2 = px.bar(top_low, x="Gross Profit %", y="Item Description", orientation="h",
                          title="Top 15 Lowest Gross Profit Items",
                          color="Gross Profit %", color_continuous_scale="Blues")
            fig2.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig2, use_container_width=True)

        with st.expander("View all low gross profit items"):
            st.write(f"**{len(low):,} items**")
            st.dataframe(
                low[["Customer", "Item Code", "Item Description", "Group",
                     "Amount", "Cost", "Gross Profit", "Gross Profit %", "Markup %"]]
                .sort_values("Customer"),
                use_container_width=True,
                column_config={"Item Description": st.column_config.TextColumn(width="small")},
            )

# ─── Tab 3: Profit Lost ───────────────────────────────────────────────────────
with tab3:
    st.subheader(f"Profit Lost — Items Below {gp_threshold}% Gross Profit Target")
    st.caption(
        "Profit Lost = the additional gross profit that *would have been earned* "
        "if each below-threshold item had been sold at exactly the GP target."
    )

    low_pl = df[df["Gross Profit %"] < gp_threshold].copy()
    low_pl["Target GP"] = (gp_threshold / 100 * low_pl["Amount"]).round(2)
    low_pl["Profit Lost"] = (low_pl["Target GP"] - low_pl["Gross Profit"]).round(2)

    total_lost   = low_pl["Profit Lost"].sum()
    total_actual = low_pl["Gross Profit"].sum()
    total_sales  = low_pl["Amount"].sum()
    total_target = low_pl["Target GP"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Sales (below-threshold items)", fmt_zar(total_sales))
    c2.metric("Actual Gross Profit Earned",          fmt_zar(total_actual))
    c3.metric(f"Target Gross Profit @ {gp_threshold}%", fmt_zar(total_target))
    c4.metric("Total Profit Lost",                   fmt_zar(total_lost))

    # ── Reconciliation (only shown when negatives are excluded) ──────────────
    if exclude_negative:
        st.markdown("---")
        st.markdown("#### Reconciliation — Impact of Removing Negative Gross Profit Items")
        st.caption("Shows what changed when negative GP items were excluded from the analysis.")

        # Below-threshold items BEFORE negative exclusion
        low_pl_before = df_with_negatives[df_with_negatives["Gross Profit %"] < gp_threshold].copy()
        low_pl_before["Target GP"]   = (gp_threshold / 100 * low_pl_before["Amount"]).round(2)
        low_pl_before["Profit Lost"] = (low_pl_before["Target GP"] - low_pl_before["Gross Profit"]).round(2)

        # The negative GP items that were removed
        neg_removed = low_pl_before[low_pl_before["Gross Profit"] < 0]

        pl_before   = low_pl_before["Profit Lost"].sum()
        pl_removed  = neg_removed["Profit Lost"].sum()
        pl_after    = total_lost  # already calculated from df (negatives excluded)

        recon = pd.DataFrame({
            "":       ["Sales (R)", "Cost (R)", "Gross Profit (R)", "Profit Lost (R)"],
            "Before (incl. negatives)": [
                low_pl_before["Amount"].sum(),
                low_pl_before["Cost"].sum(),
                low_pl_before["Gross Profit"].sum(),
                pl_before,
            ],
            "Negatives Removed": [
                neg_removed["Amount"].sum(),
                neg_removed["Cost"].sum(),
                neg_removed["Gross Profit"].sum(),
                pl_removed,
            ],
            "After (excl. negatives)": [
                low_pl["Amount"].sum(),
                low_pl["Cost"].sum(),
                low_pl["Gross Profit"].sum(),
                pl_after,
            ],
        })

        st.dataframe(
            recon.style.format({
                "Before (incl. negatives)": _space_fmt,
                "Negatives Removed":        _space_fmt,
                "After (excl. negatives)":  _space_fmt,
            }),
            use_container_width=True,
            hide_index=True,
        )

        import streamlit.components.v1 as components
        components.html("""
        <button onclick="
            const tabs = window.parent.document.querySelectorAll('button[role=tab]');
            for (let t of tabs) {
                if (t.innerText.includes('Negative')) { t.click(); window.parent.scrollTo(0,0); break; }
            }
        " style="
            background-color: #1c3d6b;
            color: white;
            border: none;
            padding: 10px 24px;
            font-size: 15px;
            font-weight: 600;
            border-radius: 6px;
            cursor: pointer;
            margin-top: 4px;
            letter-spacing: 0.5px;
        ">➖ View Negative Amounts</button>
        """, height=60)

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        lost_by_cat = low_pl.groupby("Group_Category").agg(
            Profit_Lost=("Profit Lost", "sum"),
            Items=("Item Code", "count"),
        ).reset_index()
        lost_by_cat["Category"] = lost_by_cat["Group_Category"].map(CATEGORY_LABELS).fillna("N/A")
        fig = px.bar(lost_by_cat, x="Category", y="Profit_Lost",
                     title="Profit Lost by Group Category",
                     labels={"Profit_Lost": "Profit Lost (R)"},
                     color_discrete_sequence=[NAVY])
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        lost_by_cust = low_pl.groupby("Customer").agg(
            Profit_Lost=("Profit Lost", "sum"),
        ).nlargest(10, "Profit_Lost").reset_index()
        fig2 = px.bar(lost_by_cust, x="Profit_Lost", y="Customer", orientation="h",
                      title="Top 10 Customers — Most Profit Lost",
                      labels={"Profit_Lost": "Profit Lost (R)"},
                      color_discrete_sequence=[STEEL])
        fig2.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Item-Level Profit Lost Detail")
    item_lost = low_pl.groupby(["Item Code", "Item Description", "Group"]).agg(
        Sales=("Amount", "sum"),
        Actual_GP=("Gross Profit", "sum"),
        Target_GP=("Target GP", "sum"),
        Profit_Lost=("Profit Lost", "sum"),
        Avg_GP_Pct=("Gross Profit %", "mean"),
        Transactions=("Item Code", "count"),
    ).reset_index().sort_values("Profit_Lost", ascending=False)
    item_lost["Avg_GP_Pct"] = item_lost["Avg_GP_Pct"].round(2)
    item_lost.columns = ["Item Code", "Item Description", "Group",
                         "Sales (R)", "Actual GP (R)", "Target GP (R)",
                         "Profit Lost (R)", "Avg GP %", "Transactions"]
    num_fmt = {c: _space_fmt for c in item_lost.columns if c in NUM_COLS + PCT_COLS}
    st.dataframe(item_lost.style.format(num_fmt), use_container_width=True, height=450,
                 column_config={"Item Description": st.column_config.TextColumn(width="small")})

    dl1, dl2, _ = st.columns([1, 1, 4])
    with dl1:
        st.download_button("⬇️ Download Excel", data=to_excel_bytes(item_lost),
                           file_name="profit_lost.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with dl2:
        st.download_button("⬇️ Download CSV", data=item_lost.to_csv(index=False).encode(),
                           file_name="profit_lost.csv", mime="text/csv")

    st.markdown("---")
    st.markdown("### 🎯 Action Plan — Items Making Up 80% of Profit Loss")
    st.caption("Fix these items first. Address their pricing and you recover 80% of the total profit loss.")

    # Build pareto directly from raw low_pl data
    pareto = low_pl.groupby(["Item Code", "Item Description", "Group"]).agg(
        Profit_Lost=("Profit Lost", "sum"),
        Sales=("Amount", "sum"),
        Avg_GP_Pct=("Gross Profit %", "mean"),
        Transactions=("Item Code", "count"),
    ).reset_index().sort_values("Profit_Lost", ascending=False)

    total_lost = pareto["Profit_Lost"].sum()
    pareto["Cumulative %"] = (pareto["Profit_Lost"].cumsum() / total_lost * 100).round(1)
    pareto["% of Total Loss"] = (pareto["Profit_Lost"] / total_lost * 100).round(1)

    # Keep only items up to and including the one that crosses 80%
    cutoff = pareto[pareto["Cumulative %"] <= 80]
    if len(cutoff) < len(pareto):
        cutoff = pareto.iloc[:len(cutoff) + 1]

    cutoff = cutoff.reset_index(drop=True)
    cutoff.index += 1
    cutoff.index.name = "Priority"

    st.metric("Items to action", f"{len(cutoff)} of {len(pareto)}",
              f"recovering {cutoff['Profit_Lost'].sum() / total_lost * 100:.1f}% of total profit loss")

    action_display = cutoff[["Item Code", "Item Description", "Group",
                              "Transactions", "Sales", "Avg_GP_Pct",
                              "Profit_Lost", "% of Total Loss", "Cumulative %"]].copy()
    action_display.columns = ["Item Code", "Item Description", "Group",
                               "Transactions", "Sales (R)", "Avg GP %",
                               "Profit Lost (R)", "% of Total Loss", "Cumulative %"]

    st.dataframe(
        action_display.style.format({
            "Sales (R)": _space_fmt,
            "Profit Lost (R)": _space_fmt,
            "Avg GP %": lambda v: f"{v:.2f}%",
            "% of Total Loss": lambda v: f"{v:.1f}%",
            "Cumulative %": lambda v: f"{v:.1f}%",
        }),
        use_container_width=True,
        column_config={"Item Description": st.column_config.TextColumn(width="small")},
    )

    st.download_button("⬇️ Download Action Plan", data=to_excel_bytes(action_display),
                       file_name="action_plan.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─── Tab 4: GP by Customer ────────────────────────────────────────────────────
with tab4:
    st.subheader("Gross Profit by Customer")
    all_customers = sorted(df["Customer"].dropna().unique())
    sel_cust = st.selectbox("Select Customer", all_customers)

    cust_df = df[df["Customer"] == sel_cust]
    cust_low = cust_df[cust_df["Gross Profit %"] < gp_threshold]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Sales", fmt_zar(cust_df["Amount"].sum()))
    c2.metric("Gross Profit", fmt_zar(cust_df["Gross Profit"].sum()))
    c3.metric("Avg Gross Profit %", fmt_pct(cust_df["Gross Profit %"].mean()))
    c4.metric("Low GP Items", f"{len(cust_low):,}")

    col1, col2 = st.columns(2)
    with col1:
        grp = cust_df.groupby("Group").agg(Sales=("Amount", "sum"), GP=("Gross Profit", "sum")).reset_index()
        grp["Gross Profit %"] = (grp["GP"] / grp["Sales"] * 100).round(2)
        fig = px.bar(grp.nlargest(10, "Sales"), x="Sales", y="Group", orientation="h",
                     title="Top 10 Groups by Sales", color_discrete_sequence=[NAVY])
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Sales (R)")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        if not cust_low.empty:
            fig2 = px.scatter(cust_low, x="Amount", y="Gross Profit %",
                              hover_data=["Item Code", "Item Description"],
                              title=f"Low Gross Profit Items — {sel_cust}",
                              color_discrete_sequence=[STEEL])
            fig2.add_hline(y=gp_threshold, line_dash="dash", line_color="gray",
                           annotation_text=f"{gp_threshold}% threshold")
            fig2.update_layout(xaxis_title="Sales (R)", yaxis_title="Gross Profit %")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.success("No low gross profit items for this customer.")

    st.markdown(f"#### Items Below {gp_threshold}% Gross Profit")
    st.dataframe(cust_low[["Item Code", "Item Description", "Group", "Group_Category",
                            "Quantity", "Amount", "Cost", "Gross Profit", "Gross Profit %"]],
                 use_container_width=True)

# ─── Tab 5: Item GP by Customer ──────────────────────────────────────────────
with tab5:
    st.subheader("Item Gross Profit % Across Customers")
    st.caption(
        "Shows the same inventory item sold to different customers at different GP%. "
        "Large variance between customers on the same item may indicate contract pricing opportunities."
    )

    # Build pivot: Item → Customer → avg GP%
    item_cust = df.groupby(["Item Code", "Item Description", "Customer"]).agg(
        Sales=("Amount", "sum"),
        GP=("Gross Profit", "sum"),
        GP_Pct=("Gross Profit %", "mean"),
        Qty=("Quantity", "sum"),
    ).reset_index()
    item_cust["GP_Pct"] = item_cust["GP_Pct"].round(2)

    # Filter to items sold to more than one customer (most interesting)
    col_a, col_b = st.columns([2, 2])
    with col_a:
        min_customers = st.slider("Min. number of customers per item", 1, 10, 2, key="min_cust_item")
    with col_b:
        search_item = st.text_input("Search item code or description", key="item_search")

    multi_cust_items = (
        item_cust.groupby("Item Code")["Customer"].nunique()
        .reset_index()
        .rename(columns={"Customer": "Num Customers"})
    )
    multi_cust_items = multi_cust_items[multi_cust_items["Num Customers"] >= min_customers]
    item_cust_filtered = item_cust[item_cust["Item Code"].isin(multi_cust_items["Item Code"])]

    if search_item:
        mask = (
            item_cust_filtered["Item Code"].str.contains(search_item, case=False, na=False) |
            item_cust_filtered["Item Description"].str.contains(search_item, case=False, na=False)
        )
        item_cust_filtered = item_cust_filtered[mask]

    st.markdown(f"**{item_cust_filtered['Item Code'].nunique():,} items** sold to {min_customers}+ customers")

    # Summary variance table
    variance = item_cust_filtered.groupby(["Item Code", "Item Description"]).agg(
        Customers=("Customer", "nunique"),
        Min_GP_Pct=("GP_Pct", "min"),
        Max_GP_Pct=("GP_Pct", "max"),
        Avg_GP_Pct=("GP_Pct", "mean"),
        Total_Sales=("Sales", "sum"),
    ).reset_index()
    variance["GP Variance"] = (variance["Max_GP_Pct"] - variance["Min_GP_Pct"]).round(2)
    variance = variance.sort_values("GP Variance", ascending=False)
    variance[["Min_GP_Pct", "Max_GP_Pct", "Avg_GP_Pct"]] = variance[
        ["Min_GP_Pct", "Max_GP_Pct", "Avg_GP_Pct"]].round(2)
    variance.columns = ["Item Code", "Item Description", "# Customers",
                        "Min GP %", "Max GP %", "Avg GP %", "Total Sales (R)", "GP Variance (pp)"]

    st.markdown("#### GP% Variance by Item (sorted by highest variance)")
    st.dataframe(variance, use_container_width=True, height=300,
                 column_config={"Item Description": st.column_config.TextColumn(width="small")})

    st.markdown("#### Customer-Level Detail for Selected Item")
    item_choices = sorted(item_cust_filtered["Item Code"].unique())
    if item_choices:
        sel_item = st.selectbox("Select Item Code", item_choices, key="sel_item_code")
        item_detail = item_cust_filtered[item_cust_filtered["Item Code"] == sel_item].sort_values("GP_Pct")
        item_name = item_detail["Item Description"].iloc[0] if len(item_detail) else sel_item

        st.markdown(f"**{sel_item} — {item_name}**")
        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(item_detail, x="Customer", y="GP_Pct",
                         title=f"GP% per Customer — {sel_item}",
                         labels={"GP_Pct": "Gross Profit %", "Customer": ""},
                         color="GP_Pct", color_continuous_scale=["#6b93c4", "#1c3d6b"])
            fig.add_hline(y=gp_threshold, line_dash="dash", line_color=SLATE,
                          annotation_text=f"Target {gp_threshold}%")
            fig.update_layout(xaxis_tickangle=-35)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            fig2 = px.bar(item_detail, x="Customer", y="Sales",
                          title=f"Sales (R) per Customer — {sel_item}",
                          labels={"Sales": "Sales (R)", "Customer": ""},
                          color_discrete_sequence=[NAVY_LIGHT])
            fig2.update_layout(xaxis_tickangle=-35)
            st.plotly_chart(fig2, use_container_width=True)

        st.dataframe(
            item_detail[["Customer", "Qty", "Sales", "GP", "GP_Pct"]].rename(
                columns={"Qty": "Quantity", "GP": "Gross Profit (R)", "GP_Pct": "Gross Profit %"}),
            use_container_width=True,
        )
    else:
        st.info("No items match the current filters.")

# ─── Tab 6: By Year ───────────────────────────────────────────────────────────
with tab6:
    st.subheader("Year-over-Year Gross Profit Analysis")

    yearly = df.groupby("Year").agg(
        Sales=("Amount", "sum"),
        GP=("Gross Profit", "sum"),
        Items=("Item Code", "count"),
        Low_GP=("Gross Profit %", lambda x: (x < gp_threshold).sum()),
    ).reset_index()
    yearly["Gross Profit %"] = (yearly["GP"] / yearly["Sales"] * 100).round(2)
    yearly["Low GP %"] = (yearly["Low_GP"] / yearly["Items"] * 100).round(2)

    c1, c2, c3 = st.columns(3)
    if len(yearly) >= 2:
        latest = yearly.iloc[-1]
        prev = yearly.iloc[-2]
        sales_chg = (latest["Sales"] - prev["Sales"]) / prev["Sales"] * 100
        gp_chg = (latest["GP"] - prev["GP"]) / prev["GP"] * 100
        c1.metric("Latest Year Sales", fmt_zar(latest["Sales"]), f"{sales_chg:+.2f}% vs prior")
        c2.metric("Latest Year Gross Profit", fmt_zar(latest["GP"]), f"{gp_chg:+.2f}% vs prior")
        c3.metric("Latest Gross Profit %", fmt_pct(latest["Gross Profit %"]))

    col1, col2 = st.columns(2)
    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=yearly["Year"], y=yearly["Sales"], name="Sales", marker_color=NAVY))
        fig.add_trace(go.Bar(x=yearly["Year"], y=yearly["GP"], name="Gross Profit", marker_color=NAVY_LIGHT))
        fig.update_layout(barmode="group", title="Sales & Gross Profit by Year",
                          xaxis_title="Year", yaxis_title="Amount (R)")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=yearly["Year"], y=yearly["Gross Profit %"],
                                  mode="lines+markers", name="Gross Profit %", line=dict(color=NAVY)))
        fig2.add_trace(go.Scatter(x=yearly["Year"], y=yearly["Low GP %"],
                                  mode="lines+markers", name="Low GP Items %", line=dict(color=STEEL, dash="dash")))
        fig2.add_hline(y=gp_threshold, line_dash="dot", line_color="gray",
                       annotation_text=f"{gp_threshold}% threshold")
        fig2.update_layout(title="Gross Profit % Trends", xaxis_title="Year", yaxis_title="%")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Year Summary Table")
    st.dataframe(yearly.rename(columns={"GP": "Gross Profit", "Low_GP": "Low GP Items"}),
                 use_container_width=True)

    st.markdown(f"#### Items Below {gp_threshold}% Gross Profit per Year")
    low_by_year = df[df["Gross Profit %"] < gp_threshold].groupby("Year").agg(
        Count=("Item Code", "count"),
        Sales=("Amount", "sum"),
        Avg_GP=("Gross Profit %", "mean"),
    ).reset_index()
    fig3 = px.bar(low_by_year, x="Year", y="Count",
                  title="Low Gross Profit Item Count by Year",
                  color_discrete_sequence=[STEEL])
    st.plotly_chart(fig3, use_container_width=True)

# ─── Tab 7: Raw Data ──────────────────────────────────────────────────────────
with tab7:
    st.subheader("Raw Filtered Data")
    st.write(f"**{len(df):,} rows** matching current filters")

    export_cols = ["Customer", "Item Code", "Item Description", "Date", "Year",
                   "Group", "Description", "Group_Category",
                   "Quantity", "Amount", "Cost", "Gross Profit", "Gross Profit %", "Markup %"]
    export_df = df[[c for c in export_cols if c in df.columns]]

    st.dataframe(export_df, use_container_width=True, height=500)

    dl1, dl2, _ = st.columns([1, 1, 4])
    with dl1:
        st.download_button(
            "⬇️ Download Excel",
            data=to_excel_bytes(export_df),
            file_name="inventory_gp_analysis_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dl2:
        st.download_button(
            "⬇️ Download CSV",
            data=export_df.to_csv(index=False).encode("utf-8"),
            file_name="inventory_gp_analysis_filtered.csv",
            mime="text/csv",
        )

    st.markdown("---")
    with st.expander("🔍 Group Audit — see how every Group Code is classified", expanded=False):
        st.markdown(
            "Cross-reference of every Group Code found in your uploaded data against the Group List. "
            "Use this to spot misclassifications — then download the Group List from the sidebar, fix the **Type** column, "
            "and send the updated file to have it reloaded."
        )
        # Build audit table from full_df (unfiltered) so all groups appear
        audit = (
            full_df.groupby(["Group", "Description", "Group_Category"], dropna=False)
            .agg(Row_Count=("Amount", "count"), Total_Amount=("Amount", "sum"))
            .reset_index()
        )
        audit = audit.sort_values("Group")
        audit["Group_Category_Label"] = audit["Group_Category"].apply(
            lambda x: CATEGORY_LABELS.get(x, str(x))
        )
        audit["Total_Amount"] = audit["Total_Amount"].apply(fmt_zar)
        audit_display = audit.rename(columns={
            "Group": "Group Code",
            "Description": "Group Description",
            "Group_Category_Label": "Category Applied",
            "Row_Count": "Rows",
            "Total_Amount": "Total Sales",
        })[["Group Code", "Group Description", "Category Applied", "Rows", "Total Sales"]]
        st.dataframe(audit_display, use_container_width=True, height=400)
        st.download_button(
            "⬇️ Download Group Audit",
            data=to_excel_bytes(audit_display),
            file_name="group_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ─── Tab 8: Negative Amounts ──────────────────────────────────────────────────
with tab8:
    st.subheader("Negative Amounts")

    neg_df = df_with_negatives[df_with_negatives["Gross Profit"] < 0].copy()

    if neg_df.empty:
        st.info("No negative amounts found in the uploaded data.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Negative Transactions", f"{len(neg_df):,}".replace(",", " "))
        c2.metric("Total Negative Amount", fmt_zar(neg_df["Amount"].sum()))
        c3.metric("Customers Affected", f"{neg_df['Customer'].nunique():,}")

        st.markdown("---")

        neg_display = neg_df[["Customer", "Item Code", "Item Description", "Date",
                               "Group", "Quantity", "Amount", "Cost",
                               "Gross Profit", "Gross Profit %"]].sort_values("Amount")

        st.dataframe(neg_display, use_container_width=True, height=500,
                     column_config={"Item Description": st.column_config.TextColumn(width="small")})

        nd1, nd2, _ = st.columns([1, 1, 4])
        with nd1:
            st.download_button("⬇️ Download Excel", data=to_excel_bytes(neg_display),
                               file_name="negative_amounts.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with nd2:
            st.download_button("⬇️ Download CSV", data=neg_display.to_csv(index=False).encode(),
                               file_name="negative_amounts.csv", mime="text/csv")
