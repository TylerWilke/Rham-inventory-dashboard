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
    1: "Buy Outs",
    2: "Value Add",
    3: "Manufacturing",
    4: "Capital Goods",
    5: "Regular Repairs",
    6: "Technicians Labour",
    7: "Lubricants",
    8: "Deposits",
    "N/A": "N/A",
}

# Preset GP thresholds per category (used by the Category Analysis tabs)
# Display order for category lists and tables
CATEGORY_ORDER = [4, 8, 2, 5, 3, 1, 6, 7, "N/A"]

CATEGORY_PRESETS = {
    1: {"label": "Buy Outs",              "threshold": 20, "emoji": "🛒"},
    2: {"label": "Value Add",              "threshold": 35, "emoji": "⚙️"},
    3: {"label": "Manufacturing",          "threshold": 55, "emoji": "🏭"},
    4: {"label": "Capital Goods",          "threshold": 55, "emoji": "🏗️"},
    5: {"label": "Regular Repairs",        "threshold": 35, "emoji": "🔧"},
    6: {"label": "Technicians Labour",     "threshold": 20, "emoji": "👷"},
    7: {"label": "Lubricants",             "threshold": 12, "emoji": "🛢️"},
}

def fmt_zar(value):
    return f"R {value:,.2f}".replace(",", " ")

def fmt_pct(value):
    return f"{value:.2f}%"

def fmt_num(value):
    return f"{value:,.2f}".replace(",", " ")

GROUP_LIST_PATH = Path(__file__).parent / "Group List.xlsx"


def load_group_list():
    if not GROUP_LIST_PATH.exists():
        return pd.DataFrame(columns=["Group", "Description", "Type", "Group_Category"])
    raw = pd.read_excel(GROUP_LIST_PATH, header=None)
    # Skip header row if present
    first = str(raw.iloc[0, 0]).strip().lower()
    df = raw.iloc[1:].copy() if first == "group" else raw.copy()
    # Category comes from Type (col 2), falling back to Group_Category (col 3)
    has_cat_col = df.shape[1] >= 4
    df = df.iloc[:, :4] if has_cat_col else df.iloc[:, :3]
    if has_cat_col:
        df.columns = ["Group", "Description", "Type", "Group_Category"]
        df["Group"] = df["Group"].astype(str).str.strip()
        # Type is the column users edit; Group_Category only fills gaps
        raw_cat = df["Type"].where(pd.to_numeric(df["Type"], errors="coerce").isin(range(1, 9)),
                                   df["Group_Category"])
    else:
        df.columns = ["Group", "Description", "Type"]
        df["Group"] = df["Group"].astype(str).str.strip()
        raw_cat = df["Type"]
    df["Group_Category"] = raw_cat.apply(
        lambda x: int(x) if pd.notna(x) and str(x).strip() not in ("", "None", "N/A", "nan")
                  and pd.to_numeric(x, errors="coerce") in (1, 2, 3, 4, 5, 6, 7, 8) else "N/A"
    )
    return df.reset_index(drop=True)


@st.cache_data(show_spinner="Parsing item list…")
def parse_item_list(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(io.BytesIO(file_bytes))
    raw.columns = [str(c).strip() for c in raw.columns]

    def clean_currency(val):
        if pd.isna(val):
            return None
        s = str(val).replace("R", "").replace("\xa0", "").replace(" ", "").replace(" ", "").replace(",", "")
        try:
            return float(s)
        except Exception:
            return None

    cost_col  = next((c for c in raw.columns if "unit cost" in c.lower()), None)
    price_col = next((c for c in raw.columns if "excl price" in c.lower() or "price 1" in c.lower()), None)
    group_col = next((c for c in raw.columns if c.lower() == "group"), None)
    code_col  = next((c for c in raw.columns if "item code" in c.lower()), None)
    desc_col  = next((c for c in raw.columns if "description" in c.lower()), None)

    df = pd.DataFrame()
    df["Item Code"]   = raw[code_col]  if code_col  else ""
    df["Description"] = raw[desc_col]  if desc_col  else ""
    df["Group"]       = raw[group_col].astype(str).str.strip() if group_col else ""
    df["Unit Cost"]   = raw[cost_col].apply(clean_currency)  if cost_col  else None
    df["Excl Price"]  = pd.to_numeric(raw[price_col], errors="coerce") if price_col else None
    df["GP %"]        = ((df["Excl Price"] - df["Unit Cost"]) / df["Excl Price"] * 100).where(
                          df["Excl Price"].notna() & (df["Excl Price"] > 0))
    return df.dropna(subset=["Item Code"]).reset_index(drop=True)


@st.cache_data(show_spinner="Parsing sales file…")
def parse_sales_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    ext = Path(filename).suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine=engine)

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

    def _to_cat(x):
        if pd.isna(x) or str(x).strip() in ("", "nan", "None", "N/A"):
            return "N/A"
        try:
            v = int(float(x))
            return v if v in (1, 2, 3, 4, 5, 6, 7, 8) else "N/A"
        except (ValueError, TypeError):
            return "N/A"

    merged["Group_Category"] = merged["Group_Category"].apply(_to_cat)
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


def to_excel_with_header(df: pd.DataFrame, report_name: str, filters: dict) -> bytes:
    """Export DataFrame to Excel with a filter-summary header block above the data."""
    from openpyxl.styles import Font, PatternFill, Alignment
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Build header rows: Setting | Value
        header_rows = [
            ("Company",   "Rham Equipment"),
            ("Report",    report_name),
            ("Generated", pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")),
        ] + list(filters.items())
        header_df = pd.DataFrame(header_rows, columns=["Setting", "Value"])
        # Write header block then data with a blank separator row
        header_df.to_excel(writer, index=False, startrow=0, sheet_name="Report")
        start = len(header_df) + 2          # +1 for the header row + 1 blank row
        df.to_excel(writer, index=False, startrow=start, sheet_name="Report")
        ws = writer.sheets["Report"]
        # Style the header block
        navy_fill = PatternFill("solid", fgColor="1C3D6B")
        white_font = Font(color="FFFFFF", bold=True)
        for row in ws.iter_rows(min_row=1, max_row=1):
            for cell in row:
                cell.fill = navy_fill
                cell.font = white_font
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 50
    return buf.getvalue()


def build_filter_summary(report_name: str, years, customers, categories, threshold: int, excl_neg: bool) -> dict:
    """Return an ordered dict of active filter settings for report headers."""
    yr_str   = ", ".join(str(y) for y in sorted(years)) if years else "All"
    cust_str = f"{len(customers)} customers" if len(customers) > 3 else ", ".join(customers)
    cat_str  = ", ".join(CATEGORY_LABELS.get(c, str(c)) for c in sorted(str(c) for c in categories))
    return {
        "Report":              report_name,
        "Years":               yr_str,
        "Customers":           cust_str,
        "Categories":          cat_str,
        "GP Threshold":        f"{threshold}%",
        "Negatives Excluded":  "Yes" if excl_neg else "No",
    }


def render_category_tab(cat_id: int, base_df: pd.DataFrame, sidebar_filters: dict):
    """Render a self-contained GP analysis for one category with its preset threshold."""
    preset   = CATEGORY_PRESETS[cat_id]
    label    = preset["label"]
    default  = preset["threshold"]

    _col_slider, _col_btn = st.columns([9, 1])
    with _col_btn:
        st.write("")
        st.button(
            "↺ Reset", key=f"reset_thresh_{cat_id}",
            help=f"Reset to {default}% default",
            on_click=lambda d=default, k=cat_id: st.session_state.update({f"cat_thresh_{k}": d}),
        )
    with _col_slider:
        thresh = st.slider(
            f"GP % Threshold — {label}",
            min_value=0, max_value=100, value=default, step=1,
            key=f"cat_thresh_{cat_id}",
            help=f"Default for {label} is {default}%. Drag to adjust.",
        )
    show_lost = st.toggle(
        "Show profit lost", value=False, key=f"show_lost_{cat_id}",
        help="Off: just the items under target. On: adds Target GP, Profit Lost and the Action Plan.",
    )

    cat_df = base_df[base_df["Group_Category"] == cat_id].copy()

    st.markdown(
        f"<div style='background:#1c3d6b;color:white;padding:8px 16px;border-radius:6px;"
        f"margin-bottom:12px'>"
        f"<b>{label}</b> &nbsp;·&nbsp; GP Target: <b>{thresh}%</b> &nbsp;·&nbsp; "
        f"Years: <b>{sidebar_filters['years_label']}</b> &nbsp;·&nbsp; "
        f"Customers: <b>{sidebar_filters['cust_label']}</b></div>",
        unsafe_allow_html=True,
    )

    if cat_df.empty:
        st.info(f"No data for {label} with the current sidebar filters.")
        return

    total_sales = cat_df["Amount"].sum()
    total_gp    = cat_df["Gross Profit"].sum()
    overall_gp_pct = (total_gp / total_sales * 100) if total_sales else 0
    low         = cat_df[cat_df["Gross Profit %"] < thresh].copy()
    low["Target GP"]    = (thresh / 100 * low["Amount"]).round(2)
    low["Profit Lost"]  = (low["Target GP"] - low["Gross Profit"]).round(2)
    total_lost  = low["Profit Lost"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Sales",          fmt_zar(total_sales))
    c2.metric("Total Gross Profit",   fmt_zar(total_gp))
    c3.metric("Overall GP %",         fmt_pct(overall_gp_pct))
    if show_lost:
        c4.metric(f"Profit Lost (below {thresh}%)", fmt_zar(total_lost))
    else:
        c4.metric(f"Lines Below {thresh}%", f"{len(low):,}".replace(",", " "),
                  f"{(low['Amount'].sum() / total_sales * 100) if total_sales else 0:.1f}% of sales",
                  delta_color="off")

    if low.empty:
        st.success(f"All {label} items are above the {thresh}% GP target.")
        return

    st.markdown(f"#### Items Below {thresh}% — {len(low):,} transactions")
    low_display = low[["Customer", "Item Code", "Item Description", "Group",
                        "Amount", "Gross Profit", "Gross Profit %",
                        "Target GP", "Profit Lost"]].copy()
    if not show_lost:
        low_display = low_display.drop(columns=["Target GP", "Profit Lost"])
    low_display["Notes"] = ""
    st.dataframe(
        low_display.style.format({c: _space_fmt for c in low_display.columns if c in NUM_COLS + PCT_COLS}),
        use_container_width=True, height=320,
        column_config={"Item Description": st.column_config.TextColumn(width="small"),
                       "Notes": st.column_config.TextColumn("Notes / Comments", width="medium")},
    )

    if not show_lost:
        return

    st.markdown("#### 🎯 Action Plan — 80% of Profit Loss")
    pareto = low.groupby(["Item Code", "Item Description", "Group"]).agg(
        Profit_Lost=("Profit Lost", "sum"),
        Sales=("Amount", "sum"),
        Avg_GP_Pct=("Gross Profit %", "mean"),
        Transactions=("Item Code", "count"),
        Quantity=("Quantity", "sum"),
    ).reset_index().sort_values("Profit_Lost", ascending=False)
    if not pareto.empty:
        pareto["Cumulative %"]    = (pareto["Profit_Lost"].cumsum() / total_lost * 100).round(1)
        pareto["% of Total Loss"] = (pareto["Profit_Lost"] / total_lost * 100).round(1)
        cutoff = pareto[pareto["Cumulative %"] <= 80]
        if len(cutoff) < len(pareto):
            cutoff = pareto.iloc[:len(cutoff) + 1]
        cutoff = cutoff.reset_index(drop=True)
        cutoff.index += 1
        cutoff.index.name = "Priority"
        cutoff["Notes"] = ""
        action_display = cutoff.rename(columns={
            "Profit_Lost": "Profit Lost (R)", "Sales": "Sales (R)", "Avg_GP_Pct": "Avg GP %"
        })
        st.metric("Items to action", f"{len(cutoff)} of {len(pareto)}",
                  f"recovering {cutoff['Profit_Lost'].sum() / total_lost * 100:.1f}% of profit loss")
        st.dataframe(
            action_display[["Item Code", "Item Description", "Group",
                             "Transactions", "Quantity", "Sales (R)", "Avg GP %",
                             "Profit Lost (R)", "% of Total Loss", "Cumulative %", "Notes"]].style.format({
                "Sales (R)": _space_fmt, "Profit Lost (R)": _space_fmt,
                "Quantity": _space_fmt,
                "Avg GP %": lambda v: f"{v:.2f}%",
                "% of Total Loss": lambda v: f"{v:.1f}%",
                "Cumulative %": lambda v: f"{v:.1f}%",
            }),
            use_container_width=True,
            column_config={"Item Description": st.column_config.TextColumn(width="small"),
                           "Notes": st.column_config.TextColumn("Notes / Comments", width="medium")},
        )

        filters = build_filter_summary(
            f"Action Plan — {label}",
            sidebar_filters["years"], sidebar_filters["customers"],
            [cat_id], thresh, sidebar_filters["excl_neg"],
        )
        dl_data = action_display[["Item Code", "Item Description", "Group",
                                   "Transactions", "Quantity", "Sales (R)", "Avg GP %",
                                   "Profit Lost (R)", "% of Total Loss", "Cumulative %", "Notes"]].copy()
        st.download_button(
            f"⬇️ Download {label} Action Plan",
            data=to_excel_with_header(dl_data, f"Action Plan — {label}", filters),
            file_name=f"action_plan_cat{cat_id}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


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
    item_list_file = st.file_uploader(
        "Upload Item Price List (.xlsx)",
        type=["xlsx"],
        help="Upload your item master list to check which items are below their category GP target.",
    )

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

            exclude_deposits = st.toggle(
                "Exclude deposits", value=False,
                help="Removes capital equipment deposits from every tab, chart and total.",
            )
            if exclude_deposits:
                full_df = full_df[full_df["Group_Category"] != 8]

            customers = sorted(full_df["Customer"].dropna().unique())
            years = sorted(full_df["Year"].dropna().unique().astype(int))
            all_categories = CATEGORY_ORDER
            default_categories = CATEGORY_ORDER

            sel_customers = st.multiselect("Customers", customers, default=customers)
            sel_categories = st.multiselect(
                "Group Category", all_categories, default=default_categories,
                format_func=lambda x: CATEGORY_LABELS.get(x, str(x)),
            )
            sel_years = st.multiselect("Years", years, default=years)
            gp_threshold = st.slider("Gross Profit % Threshold", 0, 100, 50, step=1)
            exclude_negative = st.checkbox("Exclude negative Gross Profit", value=True)

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
        exclude_deposits = False

    st.markdown("---")
    st.caption(
        "Group Categories:\n"
        "- **Capital Goods** (55% target)\n"
        "- **Deposits** (no target)\n"
        "- **Value Add** (35% target)\n"
        "- **Regular Repairs** (35% target)\n"
        "- **Manufacturing** (55% target)\n"
        "- **Buy Outs** (20% target)\n"
        "- **Technicians Labour** (20% target)\n"
        "- **Lubricants** (12% target)\n"
        "- **N/A** Unclassified (opt-in)"
    )

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

    **Group List** is loaded automatically from `Group List.xlsx` placed beside this app.
    """)
    st.stop()

if df is None or len(df) == 0:
    st.warning("No data matches the current filters.")
    st.stop()

tab1, tab_gp, cat4, tab_dep, cat2, cat5, cat3, cat1, cat6, cat7, tab_na, \
    tab3, tab2, tab4, tab5, tab6, tab7, tab8, tab_items = st.tabs([
    "📊 Overview",
    "🔎 GP Breakdown",
    "🏗️ Capital Goods",
    "🏦 Deposits",
    "⚙️ Value Add",
    "🔧 Repairs",
    "🏭 Manufacturing",
    "🛒 Buy Outs",
    "👷 Labour",
    "🛢️ Lubricants",
    "❓ N/A Unclassified",
    "💸 Profit Lost",
    "⚠️ Low GP Analysis",
    "👤 GP by Customer",
    "🔀 Item GP by Customer",
    "📅 By Year",
    "📋 Raw Data",
    "➖ Negative Amounts",
    "🏷️ Item Price Check",
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

    # ── GP Reconciliation — pre-filter breakdown by category ─────────────────
    st.markdown("---")
    st.markdown("### GP % Reconciliation — All Uploaded Data")
    st.caption(
        "Shows how the overall average GP % is built up from each category. "
        "Each category's **Contribution** = its GP divided by total sales across all categories — "
        "these contributions add up to the overall GP %."
    )

    # Use full_df (all uploaded data, no sidebar filters) for the reconciliation
    recon_df = full_df.copy()
    recon_total_sales = recon_df["Amount"].sum()
    recon_total_gp    = recon_df["Gross Profit"].sum()
    recon_overall_pct = (recon_total_gp / recon_total_sales * 100) if recon_total_sales else 0

    cat_recon = recon_df.groupby("Group_Category").agg(
        Sales=("Amount", "sum"),
        GP=("Gross Profit", "sum"),
    ).reset_index()
    cat_recon["Category GP %"]          = (cat_recon["GP"] / cat_recon["Sales"] * 100).round(2)
    cat_recon["Sales % of Total"]       = (cat_recon["Sales"] / recon_total_sales * 100).round(2)
    cat_recon["Contribution to GP %"]   = (cat_recon["GP"] / recon_total_sales * 100).round(2)
    cat_recon["Category"]               = cat_recon["Group_Category"].apply(
        lambda x: CATEGORY_LABELS.get(x, "N/A")
    )
    # Sort: categories 1-7 first, then N/A
    cat_recon["_sort"] = cat_recon["Group_Category"].apply(
        lambda x: CATEGORY_ORDER.index(x) if x in CATEGORY_ORDER else 99
    )
    cat_recon = cat_recon.sort_values("_sort").drop(columns="_sort")

    # Totals row
    totals = pd.DataFrame([{
        "Category":              "TOTAL",
        "Sales":                 recon_total_sales,
        "GP":                    recon_total_gp,
        "Category GP %":         round(recon_overall_pct, 2),
        "Sales % of Total":      100.0,
        "Contribution to GP %":  round(recon_overall_pct, 2),
    }])
    recon_display = pd.concat(
        [cat_recon[["Category", "Sales", "GP", "Category GP %",
                    "Sales % of Total", "Contribution to GP %"]],
         totals],
        ignore_index=True,
    )

    # Format numbers
    recon_styled = recon_display.copy()
    recon_styled["Sales"]               = recon_styled["Sales"].apply(fmt_zar)
    recon_styled["GP"]                  = recon_styled["GP"].apply(fmt_zar)
    recon_styled["Category GP %"]       = recon_styled["Category GP %"].apply(lambda v: f"{v:.2f}%")
    recon_styled["Sales % of Total"]    = recon_styled["Sales % of Total"].apply(lambda v: f"{v:.2f}%")
    recon_styled["Contribution to GP %"] = recon_styled["Contribution to GP %"].apply(lambda v: f"{v:.2f}%")

    st.dataframe(
        recon_styled.style.apply(
            lambda row: ["font-weight:bold; background-color:#e8eef5"] * len(row)
            if row["Category"] == "TOTAL" else [""] * len(row),
            axis=1,
        ),
        use_container_width=True,
        hide_index=True,
    )

    # Bar chart — contribution to overall GP%
    chart_data = cat_recon[cat_recon["Contribution to GP %"] != 0].copy()
    fig_recon = px.bar(
        chart_data, x="Category", y="Contribution to GP %",
        title=f"Each Category's Contribution to Overall GP % ({recon_overall_pct:.2f}%)",
        color_discrete_sequence=[NAVY],
        labels={"Contribution to GP %": "Contribution (pp)"},
        text=chart_data["Contribution to GP %"].apply(lambda v: f"{v:.2f}%"),
    )
    fig_recon.add_hline(
        y=recon_overall_pct, line_dash="dash", line_color=SLATE,
        annotation_text=f"Overall {recon_overall_pct:.2f}%",
    )
    fig_recon.update_traces(textposition="outside")
    fig_recon.update_layout(yaxis_title="Contribution to GP % (pp)")
    st.plotly_chart(fig_recon, use_container_width=True)

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

# ─── GP Breakdown: what makes up total Gross Profit ──────────────────────────
with tab_gp:
    bd_sales = df["Amount"].sum()
    bd_gp    = df["Gross Profit"].sum()
    bd_pct   = (bd_gp / bd_sales * 100) if bd_sales else 0

    # Lines with no (or negative) cost book the whole sale as profit — deposits,
    # admin fees, or items where the cost was never captured.
    no_cost      = df[df["Cost"].fillna(0) <= 0]
    no_cost_gp   = no_cost["Gross Profit"].sum()
    costed       = df[df["Cost"].fillna(0) > 0]
    costed_sales = costed["Amount"].sum()
    costed_pct   = (costed["Gross Profit"].sum() / costed_sales * 100) if costed_sales else 0

    item_gp = (df.groupby(["Item Code", "Item Description"], as_index=False)["Gross Profit"].sum()
                 .sort_values("Gross Profit", ascending=False))
    top10_share = (item_gp["Gross Profit"].head(10).sum() / bd_gp * 100) if bd_gp else 0

    st.subheader("What Makes Up Total Gross Profit")
    st.caption("Uses the sidebar filters. Start with the zero-cost lines — they count the full sale as profit.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Gross Profit", f"R {bd_gp / 1e6:,.1f}M", f"{bd_pct:.2f}% GP", delta_color="off")
    c2.metric("GP from Zero/Negative Cost Lines", f"R {no_cost_gp / 1e6:,.1f}M",
              f"{(no_cost_gp / bd_gp * 100) if bd_gp else 0:.2f}% of total GP", delta_color="off")
    c3.metric("GP % Excluding Zero-Cost Lines", fmt_pct(costed_pct),
              f"{costed_pct - bd_pct:+.2f} pp vs reported", delta_color="off")
    c4.metric("Top 10 Items' Share of GP", fmt_pct(top10_share))

    # ── Waterfall: how each category builds up the total ─────────────────────
    st.markdown("---")
    wf = df.groupby("Group_Category", as_index=False).agg(
        Sales=("Amount", "sum"), **{"Gross Profit": ("Gross Profit", "sum")})
    wf["Category"] = wf["Group_Category"].map(CATEGORY_LABELS).fillna("N/A")
    wf = wf.sort_values("Gross Profit", ascending=False)
    # Category name, achieved GP %, and its target (green = target met, red = missed)
    wf_pct = (wf["Gross Profit"] / wf["Sales"] * 100).where(wf["Sales"] != 0)
    wf_target = wf["Group_Category"].map(
        lambda c: CATEGORY_PRESETS[c]["threshold"] if c in CATEGORY_PRESETS else None)

    def _wf_label(name, achieved, target):
        if pd.isna(achieved):
            return name
        if target is None or pd.isna(target):
            return f"{name}<br><b>Achieved {achieved:.1f}%</b><br>Target —"
        colour = "#1e8e3e" if achieved >= target else "#d93025"
        return (f"{name}<br><b>Achieved <span style='color:{colour}'>{achieved:.1f}%</span></b>"
                f"<br>Target {target:.0f}%")

    wf_x = [_wf_label(c, p, t) for c, p, t in zip(wf["Category"], wf_pct, wf_target)]
    # Blended target: each category's target weighted by its sales (categories with targets only)
    has_t = wf_target.notna()
    t_sales = wf.loc[has_t, "Sales"].sum()
    blended = (wf.loc[has_t, "Sales"] * wf_target[has_t]).sum() / t_sales if t_sales else None
    total_x = (f"TOTAL GP<br><b>Achieved {bd_pct:.1f}%</b><br>Blended target {blended:.1f}%"
               if blended is not None else f"TOTAL GP<br><b>Achieved {bd_pct:.1f}%</b>")
    fig_wf = go.Figure(go.Waterfall(
        x=wf_x + [total_x],
        y=list(wf["Gross Profit"]) + [bd_gp],
        measure=["relative"] * len(wf) + ["total"],
        text=[fmt_zar(v) for v in wf["Gross Profit"]] + [fmt_zar(bd_gp)],
        textposition="outside",
        increasing={"marker": {"color": NAVY_LIGHT}},
        decreasing={"marker": {"color": STEEL}},
        totals={"marker": {"color": NAVY}},
    ))
    fig_wf.update_layout(title="Gross Profit Build-up by Category", yaxis_title="Gross Profit (R)",
                         showlegend=False, margin={"b": 110}, xaxis={"tickangle": 0})
    st.plotly_chart(fig_wf, use_container_width=True)

    # ── Drill-down: rank any dimension by GP contribution ────────────────────
    st.markdown("---")
    st.markdown("### GP Contributors")
    dim_options = {
        "Group": ["Group", "Description"],
        "Item": ["Item Code", "Item Description", "Group"],
        "Customer": ["Customer"],
        "Group Category": ["Group_Category"],
        "Year": ["Year"],
    }
    dc1, dc2, dc3 = st.columns([2, 1, 1])
    dim = dc1.radio("Break GP down by", list(dim_options), horizontal=True, key="gp_bd_dim")
    cat_choices = ["All"] + [c for c in CATEGORY_ORDER if c in set(df["Group_Category"])]
    if dim == "Group Category":
        contrib_cat = "All"
    else:
        contrib_cat = dc2.selectbox(
            "Category", cat_choices, key="gp_bd_cat",
            format_func=lambda c: "All categories" if c == "All" else CATEGORY_LABELS.get(c, str(c)),
        )
    top_n = dc3.slider("Show top", 5, 50, 20, step=5, key="gp_bd_topn")

    # Pareto within one category: shares are of that category's GP
    contrib_src = df if contrib_cat == "All" else df[df["Group_Category"] == contrib_cat]
    contrib_gp = contrib_src["Gross Profit"].sum()
    scope_label = "total GP" if contrib_cat == "All" else f"{CATEGORY_LABELS.get(contrib_cat, contrib_cat)} GP"

    keys = dim_options[dim]
    contrib = contrib_src.groupby(keys, dropna=False).agg(
        Lines=("Amount", "size"),
        Quantity=("Quantity", "sum"),
        Sales=("Amount", "sum"),
        Cost=("Cost", "sum"),
        GP=("Gross Profit", "sum"),
    ).reset_index().sort_values("GP", ascending=False)
    contrib["GP %"]          = (contrib["GP"] / contrib["Sales"] * 100).where(contrib["Sales"] != 0)
    contrib["Share of GP %"] = contrib["GP"] / contrib_gp * 100 if contrib_gp else 0
    contrib["Cumulative %"]  = contrib["Share of GP %"].cumsum()
    if dim == "Group Category":
        contrib["Group_Category"] = contrib["Group_Category"].map(CATEGORY_LABELS).fillna("N/A")
    contrib["Label"] = contrib[keys[0]].astype(str)

    top = contrib.head(top_n)
    fig_p = go.Figure()
    fig_p.add_bar(x=top["Label"], y=top["GP"], name="Gross Profit", marker_color=NAVY)
    fig_p.add_scatter(x=top["Label"], y=top["Cumulative %"], name="Cumulative % of GP",
                      yaxis="y2", mode="lines+markers", line={"color": STEEL})
    fig_p.update_layout(
        title=f"Top {len(top)} by Gross Profit — {top['Share of GP %'].sum():.1f}% of {scope_label}",
        yaxis={"title": "Gross Profit (R)"},
        yaxis2={"title": "Cumulative %", "overlaying": "y", "side": "right", "range": [0, 105]},
        legend={"orientation": "h", "y": -0.25},
    )
    st.plotly_chart(fig_p, use_container_width=True)

    contrib_disp = contrib.drop(columns="Label").head(top_n).copy()
    for col in ("Sales", "Cost", "GP"):
        contrib_disp[col] = contrib_disp[col].apply(fmt_zar)
    for col in ("GP %", "Share of GP %", "Cumulative %"):
        contrib_disp[col] = contrib_disp[col].apply(lambda v: fmt_pct(v) if pd.notna(v) else "")
    st.dataframe(contrib_disp, use_container_width=True, hide_index=True)
    st.download_button(
        f"⬇️ Download full GP by {dim}",
        data=to_excel_bytes(contrib.drop(columns="Label")),
        file_name=f"GP by {dim}"
                  + ("" if contrib_cat == "All" else f" - {CATEGORY_LABELS.get(contrib_cat, contrib_cat)}")
                  + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── Treemap: category → group, sized by GP ───────────────────────────────
    st.markdown("---")
    tm = df.groupby(["Group_Category", "Group"], as_index=False).agg(
        Sales=("Amount", "sum"), GP=("Gross Profit", "sum"))
    tm = tm[tm["GP"] > 0]
    tm["Category"] = tm["Group_Category"].map(CATEGORY_LABELS).fillna("N/A")
    tm["GP %"] = (tm["GP"] / tm["Sales"] * 100).round(1)
    fig_tm = px.treemap(tm, path=["Category", "Group"], values="GP", color="GP %",
                        color_continuous_scale="Blues", hover_data={"Sales": ":,.0f", "GP": ":,.0f"},
                        title="Where the Gross Profit Sits — box size = GP, colour = GP %")
    st.plotly_chart(fig_tm, use_container_width=True)
    st.caption("Groups with zero or negative total GP are left out of the treemap.")

    # ── Zero / negative cost lines ───────────────────────────────────────────
    st.markdown("---")
    st.markdown(f"### Zero or Negative Cost Lines  ({len(no_cost):,} lines)".replace(",", " "))
    if no_cost.empty:
        st.success("Every line has a cost recorded.")
    else:
        st.caption(
            "These lines have no cost against them, so the whole sale counts as gross profit. "
            "Check whether the cost was missed (e.g. machines, repairs) or whether the line "
            "isn't really a sale (e.g. deposits, admin)."
        )
        nc = no_cost.groupby(["Group", "Item Code", "Item Description"], as_index=False).agg(
            Lines=("Amount", "size"),
            Sales=("Amount", "sum"),
            Cost=("Cost", "sum"),
            GP=("Gross Profit", "sum"),
        ).sort_values("GP", ascending=False)
        nc["Share of Total GP %"] = nc["GP"] / bd_gp * 100 if bd_gp else 0
        nc_disp = nc.copy()
        for col in ("Sales", "Cost", "GP"):
            nc_disp[col] = nc_disp[col].apply(fmt_zar)
        nc_disp["Share of Total GP %"] = nc_disp["Share of Total GP %"].apply(fmt_pct)
        st.dataframe(nc_disp, use_container_width=True, hide_index=True, height=350)

        with st.expander("View every zero/negative cost line"):
            st.dataframe(
                no_cost[["Customer", "Item Code", "Item Description", "Date", "Group",
                         "Quantity", "Amount", "Cost", "Gross Profit", "Gross Profit %"]]
                .sort_values("Gross Profit", ascending=False),
                use_container_width=True, hide_index=True,
            )
        st.download_button(
            "⬇️ Download zero/negative cost lines",
            data=to_excel_bytes(no_cost),
            file_name="Zero Cost Lines.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── Biggest single lines ─────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 20 Biggest Single Lines by Gross Profit")
    big = df.nlargest(20, "Gross Profit")[["Customer", "Item Code", "Item Description", "Date", "Group",
                                           "Quantity", "Amount", "Cost", "Gross Profit", "Gross Profit %"]]
    big = big.assign(**{"Share of Total GP %": big["Gross Profit"] / bd_gp * 100 if bd_gp else 0})
    st.dataframe(big, use_container_width=True, hide_index=True)

# ─── N/A Unclassified: lines with no Group Category ──────────────────────────
with tab_na:
    # Ignores the Group Category filter so N/A lines show even when deselected
    na_base = full_df.copy()
    if sel_customers:
        na_base = na_base[na_base["Customer"].isin(sel_customers)]
    if sel_years:
        na_base = na_base[na_base["Year"].isin(sel_years)]
    if exclude_negative:
        na_base = na_base[na_base["Gross Profit"] >= 0]
    na = na_base[na_base["Group_Category"] == "N/A"].copy()

    st.subheader("N/A — Lines Without a Group Category")
    st.caption("Uses the Customer, Year and negative-GP filters, but always shows N/A lines — "
               "even if N/A is deselected under Group Category.")

    if na.empty:
        st.success("Every line has a Group Category.")
    else:
        listed_groups = set(group_list["Group"])

        def _na_reason(g):
            if not g or g in ("nan", "None"):
                return "No group code on the line"
            if g in listed_groups:
                return "In Group List but Type is blank"
            return "Group code missing from Group List"

        na["Reason"] = na["Group"].apply(_na_reason)
        na_sales = na["Amount"].sum()
        na_gp    = na["Gross Profit"].sum()
        base_gp  = na_base["Gross Profit"].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("N/A Sales", f"R {na_sales / 1e6:,.1f}M")
        c2.metric("N/A Gross Profit", f"R {na_gp / 1e6:,.1f}M",
                  f"{(na_gp / base_gp * 100) if base_gp else 0:.2f}% of total GP", delta_color="off")
        c3.metric("N/A GP %", fmt_pct(na_gp / na_sales * 100) if na_sales else "—")
        c4.metric("Group Codes to Classify", f"{na['Group'].nunique():,}")

        # ── Why lines are N/A ────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### Why These Lines Are N/A")
        reasons = na.groupby("Reason", as_index=False).agg(
            Groups=("Group", "nunique"),
            Lines=("Amount", "size"),
            Sales=("Amount", "sum"),
            GP=("Gross Profit", "sum"),
        ).sort_values("GP", ascending=False)
        reasons_disp = reasons.copy()
        for col in ("Sales", "GP"):
            reasons_disp[col] = reasons_disp[col].apply(fmt_zar)
        st.dataframe(reasons_disp, use_container_width=True, hide_index=True)
        st.caption(
            "**Missing from Group List** — add the code to the Group List with a Type (1–7).  \n"
            "**Type is blank** — the code is in the Group List but has no Type; fill it in.  \n"
            "**No group code** — the item has no group in Sage; fix the item master."
        )

        # ── By group code ────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### N/A Gross Profit by Group Code")
        by_grp = na.groupby(["Group", "Reason"], as_index=False, dropna=False).agg(
            Lines=("Amount", "size"),
            Sales=("Amount", "sum"),
            Cost=("Cost", "sum"),
            GP=("Gross Profit", "sum"),
        ).sort_values("GP", ascending=False)
        desc = group_list.drop_duplicates("Group").set_index("Group")["Description"]
        by_grp.insert(1, "Description", by_grp["Group"].map(desc))
        by_grp["GP %"]            = (by_grp["GP"] / by_grp["Sales"] * 100).where(by_grp["Sales"] != 0)
        by_grp["Share of N/A GP %"] = by_grp["GP"] / na_gp * 100 if na_gp else 0

        top_grp = by_grp.head(15).copy()
        top_grp["Label"] = top_grp["Group"].replace("", "(blank)")
        fig_na = px.bar(top_grp, x="GP", y="Label", orientation="h", color="Reason",
                        title="Top 15 N/A Group Codes by Gross Profit",
                        color_discrete_sequence=[NAVY, STEEL, SLATE],
                        labels={"GP": "Gross Profit (R)", "Label": ""})
        fig_na.update_layout(yaxis={"categoryorder": "total ascending"}, legend_title="")
        st.plotly_chart(fig_na, use_container_width=True)

        grp_disp = by_grp.copy()
        grp_disp["Group"] = grp_disp["Group"].replace("", "(blank)")
        for col in ("Sales", "Cost", "GP"):
            grp_disp[col] = grp_disp[col].apply(fmt_zar)
        for col in ("GP %", "Share of N/A GP %"):
            grp_disp[col] = grp_disp[col].apply(lambda v: fmt_pct(v) if pd.notna(v) else "")
        st.dataframe(grp_disp, use_container_width=True, hide_index=True, height=350)

        # Same layout as the Group List so Type can be filled in and sent back
        to_classify = by_grp[["Group", "Description"]].assign(Type="", Group_Category="")
        to_classify = to_classify.join(by_grp[["Lines", "Sales", "GP"]])
        to_classify = to_classify[to_classify["Group"] != ""]
        st.download_button(
            "⬇️ Download N/A groups to classify",
            data=to_excel_bytes(to_classify),
            file_name="N-A Groups to Classify.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Fill in the Type column (1–7), then add these rows to the Group List.",
        )

        # ── Top items ────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### Top 20 N/A Items by Gross Profit")
        na_items = na.groupby(["Item Code", "Item Description", "Group"], as_index=False).agg(
            Lines=("Amount", "size"),
            Sales=("Amount", "sum"),
            Cost=("Cost", "sum"),
            GP=("Gross Profit", "sum"),
        ).sort_values("GP", ascending=False).head(20)
        na_items["Share of N/A GP %"] = na_items["GP"] / na_gp * 100 if na_gp else 0
        items_disp = na_items.copy()
        for col in ("Sales", "Cost", "GP"):
            items_disp[col] = items_disp[col].apply(fmt_zar)
        items_disp["Share of N/A GP %"] = items_disp["Share of N/A GP %"].apply(fmt_pct)
        st.dataframe(items_disp, use_container_width=True, hide_index=True)

        with st.expander(f"View all {len(na):,} N/A lines".replace(",", " ")):
            st.dataframe(
                na[["Customer", "Item Code", "Item Description", "Date", "Group", "Reason",
                    "Quantity", "Amount", "Cost", "Gross Profit", "Gross Profit %"]]
                .sort_values("Gross Profit", ascending=False),
                use_container_width=True, hide_index=True,
            )
        st.download_button(
            "⬇️ Download all N/A lines",
            data=to_excel_bytes(na),
            file_name="N-A Lines.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ─── Deposits: category 8, kept apart from product sales ─────────────────────
with tab_dep:
    st.subheader("Deposits — Capital Equipment")
    if exclude_deposits:
        st.info("Deposits are excluded from the whole dashboard. Switch off **Exclude deposits** "
                "in the sidebar to see them here.")

    # Ignores the Group Category and negative-GP filters: deposits are invoiced and later
    # reversed with credit notes, so both sides are needed to see the net amount.
    dep_base = full_df.copy()
    if sel_customers:
        dep_base = dep_base[dep_base["Customer"].isin(sel_customers)]
    if sel_years:
        dep_base = dep_base[dep_base["Year"].isin(sel_years)]
    dep = dep_base[dep_base["Group_Category"] == 8]

    # Headline comparison counts lines the same way the rest of the dashboard does
    head_base = dep_base[dep_base["Gross Profit"] >= 0] if exclude_negative else dep_base
    non_dep   = head_base[head_base["Group_Category"] != 8]

    if dep.empty:
        if not exclude_deposits:
            st.info("No deposit lines in the current selection.")
    else:
        st.caption("Deposits carry no cost, so every deposit is 100% gross profit. Each deposit is "
                   "invoiced, then reversed with a credit note when the machine is invoiced — Sage "
                   "shows those credit notes at 0% GP, which is why averaged GP % looks lower.")
        invoiced = dep.loc[dep["Amount"] > 0, "Amount"].sum()
        reversed_ = dep.loc[dep["Amount"] < 0, "Amount"].sum()
        net_dep  = invoiced + reversed_

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Deposits Invoiced", f"R {invoiced / 1e6:,.1f}M",
                  f"{(dep['Amount'] > 0).sum():,} lines", delta_color="off")
        c2.metric("Reversed (Credit Notes)", f"R {reversed_ / 1e6:,.1f}M",
                  f"{(dep['Amount'] < 0).sum():,} lines", delta_color="off")
        c3.metric("Net Deposits", f"R {net_dep / 1e6:,.1f}M")
        c4.metric("Deposit Cost", fmt_zar(dep["Cost"].sum()))

        all_sales, all_gp = head_base["Amount"].sum(), head_base["Gross Profit"].sum()
        nd_sales, nd_gp   = non_dep["Amount"].sum(), non_dep["Gross Profit"].sum()
        all_pct = (all_gp / all_sales * 100) if all_sales else 0
        nd_pct  = (nd_gp / nd_sales * 100) if nd_sales else 0
        counted = head_base.loc[head_base["Group_Category"] == 8, "Gross Profit"].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Deposit GP Counted in Totals", f"R {counted / 1e6:,.1f}M")
        c2.metric("Share of Total GP", fmt_pct(counted / all_gp * 100) if all_gp else "—")
        c3.metric("GP % Including Deposits", fmt_pct(all_pct))
        c4.metric("GP % Excluding Deposits", fmt_pct(nd_pct), f"{nd_pct - all_pct:+.2f} pp", delta_color="off")
        if exclude_negative and reversed_ < 0:
            st.warning(f"**Exclude negative Gross Profit** is on, so the {fmt_zar(-reversed_)} of credit notes "
                       f"are dropped and the dashboard counts deposits at {fmt_zar(invoiced)} instead of the "
                       f"net {fmt_zar(net_dep)}.")

        st.markdown("---")
        col_left, col_right = st.columns(2)
        with col_left:
            dep_cust = dep.groupby("Customer", as_index=False).agg(
                Lines=("Amount", "size"), Deposits=("Amount", "sum")
            ).sort_values("Deposits", ascending=False)
            fig_dc = px.bar(dep_cust.head(15), x="Deposits", y="Customer", orientation="h",
                            title="Net Deposits by Customer", color_discrete_sequence=[NAVY])
            fig_dc.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_title="Deposits (R)")
            st.plotly_chart(fig_dc, use_container_width=True)
        with col_right:
            dep_month = (dep.dropna(subset=["Date"])
                            .assign(Month=lambda d: d["Date"].dt.to_period("M").dt.to_timestamp())
                            .groupby("Month", as_index=False)["Amount"].sum())
            fig_dm = px.bar(dep_month, x="Month", y="Amount", title="Net Deposits by Month",
                            color_discrete_sequence=[NAVY_LIGHT], labels={"Amount": "Deposits (R)"})
            st.plotly_chart(fig_dm, use_container_width=True)

        st.markdown("#### Deposit Lines")
        dep_cols = [c for c in ["Customer", "Item Code", "Item Description", "Date", "Group",
                                "Invoice", "Quantity", "Amount", "Cost", "Gross Profit"] if c in dep.columns]
        st.dataframe(dep[dep_cols].sort_values("Amount", ascending=False),
                     use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download deposit lines",
            data=to_excel_bytes(dep[dep_cols]),
            file_name="Deposits.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

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

    item_lost["Notes"] = ""
    pl_filters = build_filter_summary(
        "Profit Lost Detail", sel_years, sel_customers, sel_categories, gp_threshold, exclude_negative
    )
    dl1, dl2, _ = st.columns([1, 1, 4])
    with dl1:
        st.download_button("⬇️ Download Excel",
                           data=to_excel_with_header(item_lost, "Profit Lost Detail", pl_filters),
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

    action_display["Notes"] = ""
    ap_filters = build_filter_summary(
        "Action Plan", sel_years, sel_customers, sel_categories, gp_threshold, exclude_negative
    )
    st.download_button(
        "⬇️ Download Action Plan",
        data=to_excel_with_header(action_display, "Action Plan", ap_filters),
        file_name="action_plan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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

    export_df = export_df.copy()
    export_df["Notes"] = ""
    raw_filters = build_filter_summary(
        "Raw Data", sel_years, sel_customers, sel_categories, gp_threshold, exclude_negative
    )
    dl1, dl2, _ = st.columns([1, 1, 4])
    with dl1:
        st.download_button(
            "⬇️ Download Excel",
            data=to_excel_with_header(export_df, "Raw Data", raw_filters),
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

# ─── Category Analysis Tabs (1-7) ─────────────────────────────────────────────
# Each tab uses the sidebar year/customer filters but its own preset GP threshold.
_sidebar_filters = {
    "years":       sel_years,
    "customers":   sel_customers,
    "excl_neg":    exclude_negative,
    "years_label": ", ".join(str(y) for y in sorted(sel_years)) if sel_years else "All",
    "cust_label":  f"{len(sel_customers)} customers" if len(sel_customers) > 3
                   else (", ".join(sel_customers) if sel_customers else "All"),
}
# Base df for category tabs: apply customer/year filters but NOT the category filter
# (each tab handles its own category) and NOT the GP threshold filter
_base_cat_df = full_df.copy()
if sel_customers:
    _base_cat_df = _base_cat_df[_base_cat_df["Customer"].isin(sel_customers)]
if sel_years:
    _base_cat_df = _base_cat_df[_base_cat_df["Year"].isin(sel_years)]
if exclude_negative:
    _base_cat_df = _base_cat_df[_base_cat_df["Gross Profit"] >= 0]

_cat_tabs = [cat1, cat2, cat3, cat4, cat5, cat6, cat7]
for _cat_id, _tab in zip(range(1, 8), _cat_tabs):
    with _tab:
        render_category_tab(_cat_id, _base_cat_df, _sidebar_filters)

# ─── Tab: Item Price Check ────────────────────────────────────────────────────
with tab_items:
    st.subheader("🏷️ Item Price Check — Items Below Category GP Target")

    if item_list_file is None:
        st.info("👈 Upload an **Item Price List (.xlsx)** in the sidebar to use this tab.")
    else:
        item_df = parse_item_list(item_list_file.getvalue())
        gl = load_group_list()[["Group", "Group_Category"]].copy()
        item_df = item_df.merge(gl, on="Group", how="left")
        item_df["Group_Category"] = item_df["Group_Category"].fillna("N/A")

        # Attach the preset threshold for each item's category
        item_df["Target GP %"] = item_df["Group_Category"].apply(
            lambda c: CATEGORY_PRESETS[c]["threshold"] if c in CATEGORY_PRESETS else None
        )
        item_df["Category Label"] = item_df["Group_Category"].apply(
            lambda c: CATEGORY_LABELS.get(c, "N/A")
        )
        item_df["Below Target"] = (
            item_df["GP %"].notna()
            & item_df["Target GP %"].notna()
            & (item_df["GP %"] < item_df["Target GP %"])
        )

        classified = item_df[item_df["Group_Category"].isin(range(1, 8))]
        total_items   = len(classified)
        below_items   = classified["Below Target"].sum()
        below_pct     = below_items / total_items * 100 if total_items else 0

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Items (Cat 1-7)", f"{total_items:,}".replace(",", " "))
        c2.metric("Items Below GP Target", f"{int(below_items):,}".replace(",", " "))
        c3.metric("% Below Target", f"{below_pct:.1f}%")

        st.markdown("---")

        for cat_id in range(1, 8):
            preset = CATEGORY_PRESETS[cat_id]
            cat_items = item_df[item_df["Group_Category"] == cat_id].copy()
            if cat_items.empty:
                continue
            below = cat_items[cat_items["Below Target"]]
            n_total = len(cat_items)
            n_below = len(below)

            with st.expander(
                f"{preset['emoji']} {preset['label']} — {n_below} of {n_total} items below {preset['threshold']}% GP",
                expanded=(n_below > 0),
            ):
                if below.empty:
                    st.success(f"All {n_total} items meet the {preset['threshold']}% GP target.")
                else:
                    display = below[["Item Code", "Description", "Group", "Unit Cost",
                                     "Excl Price", "GP %", "Target GP %"]].copy()
                    display["Gap"] = display["Target GP %"] - display["GP %"]
                    display = display.sort_values("GP %")
                    st.dataframe(
                        display.style.format({
                            "Unit Cost":   "R {:,.2f}",
                            "Excl Price":  "R {:,.2f}",
                            "GP %":        "{:.2f}%",
                            "Target GP %": "{:.0f}%",
                            "Gap":         "{:.2f}%",
                        }),
                        use_container_width=True, height=400,
                    )
                    dl1, dl2, _ = st.columns([1, 1, 4])
                    with dl1:
                        st.download_button(
                            "⬇️ Download Excel",
                            data=to_excel_bytes(display),
                            file_name=f"below_target_cat{cat_id}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"item_dl_xl_{cat_id}",
                        )
                    with dl2:
                        st.download_button(
                            "⬇️ Download CSV",
                            data=display.to_csv(index=False).encode(),
                            file_name=f"below_target_cat{cat_id}.csv",
                            mime="text/csv",
                            key=f"item_dl_csv_{cat_id}",
                        )

        st.markdown("---")
        st.subheader("All Items — Full List")
        full_display = item_df[["Item Code", "Description", "Group", "Category Label",
                                 "Unit Cost", "Excl Price", "GP %", "Target GP %", "Below Target"]].copy()
        full_display = full_display.sort_values(["Category Label", "GP %"])
        st.dataframe(full_display, use_container_width=True, height=500)
        st.download_button(
            "⬇️ Download Full Item List with GP Check",
            data=to_excel_bytes(full_display),
            file_name="item_gp_check.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
