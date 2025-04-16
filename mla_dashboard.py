import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# --- Config ---
st.set_page_config("MLA Saleyard Dashboard", layout="wide")

# --- Load Data ---
@st.cache_data
def load_data():
    df = pd.read_excel("final_mla_output.xlsx")
    df["Report Date"] = pd.to_datetime(df["Report Date"], dayfirst=True, errors='coerce')
    for col in ["Saleyard", "Category", "Weight Range", "Sale Prefix"]:
        df[col] = df[col].astype(str).str.strip()
    fav_df = pd.read_csv("favourites.csv")
    gus_row = fav_df[fav_df["User"] == "Gus"]
    favourites = gus_row.iloc[0][1:].dropna().tolist() if not gus_row.empty else []
    return df, favourites

# --- Load data ---
df, favourites = load_data()
df_filtered = df.copy()

# --- Sidebar Filters ---
st.sidebar.header("🔍 Filters")

date_filter = st.sidebar.selectbox("Report Date Range", [
    "Last 7 Days", "Last 30 Days", "Last 60 Days", "Last 90 Days", "Last 180 Days"
])
date_cutoffs = {
    "Last 7 Days": datetime.datetime.today() - datetime.timedelta(days=7),
    "Last 30 Days": datetime.datetime.today() - datetime.timedelta(days=30),
    "Last 60 Days": datetime.datetime.today() - datetime.timedelta(days=60),
    "Last 90 Days": datetime.datetime.today() - datetime.timedelta(days=90),
    "Last 180 Days": datetime.datetime.today() - datetime.timedelta(days=180)
}
df_filtered = df[df["Report Date"] >= date_cutoffs[date_filter]]

# --- Saleyard Filter ---
saleyard_options = sorted(set(df_filtered["Saleyard"].dropna().unique()).union(favourites))
default_selection = [f for f in favourites if f in saleyard_options]
fav_selected = None

saleyards = st.sidebar.multiselect("Saleyards", saleyard_options, default=default_selection)
if saleyards:
    df_filtered = df_filtered[df_filtered["Saleyard"].isin(saleyards)]

# --- Dynamic Filters ---
filtered_categories = sorted(df_filtered["Category"].dropna().unique())
categories = st.sidebar.multiselect("Category", filtered_categories, key="category")
if categories:
    df_filtered = df_filtered[df_filtered["Category"].isin(categories)]

filtered_weights = sorted(df_filtered["Weight Range"].dropna().unique())
weights = st.sidebar.multiselect("Weight Range", filtered_weights, key="weight")
if weights:
    df_filtered = df_filtered[df_filtered["Weight Range"].isin(weights)]

filtered_prefixes = sorted(df_filtered["Sale Prefix"].dropna().unique())
prefixes = st.sidebar.multiselect("Sale Prefix", filtered_prefixes, key="prefix")
if prefixes:
    df_filtered = df_filtered[df_filtered["Sale Prefix"].isin(prefixes)]

# --- Spacer ---
st.sidebar.markdown("---")



# --- Favourite Buttons (stacked in sidebar) ---
st.sidebar.markdown("⭐ **Favourite Saleyards**")
for yard in favourites:
    if st.sidebar.button(yard):
        fav_selected = yard
        st.experimental_rerun()

# --- Re-apply saleyard if selected via button ---
if fav_selected and fav_selected in saleyard_options:
    saleyards = [fav_selected]
    df_filtered = df_filtered[df_filtered["Saleyard"].isin(saleyards)]

# --- Pivot Table ---
st.header("Boonz's Table")

if df_filtered.empty:
    st.warning("No data matches your filters.")
else:
    def custom_aggregations(group):
        head_sum = group["Head Count"].sum()
        if head_sum == 0:
            return pd.Series({
                "Sum of Av LW": 0,
                "Sum of Av c/kg LW": 0,
                "Sum of Av $/hd": 0
            })
        lw_numerator = ((group["Avg $/Head"] / group["Avg Lwt c/kg"]) * group["Head Count"]).sum()
        ckg_numerator = (group["Avg Lwt c/kg"] * group["Head Count"]).sum()
        dollar_numerator = (group["Avg $/Head"] * group["Head Count"]).sum()
        return pd.Series({
            "Sum of Av LW": (lw_numerator / head_sum) * 100,
            "Sum of Av c/kg LW": ckg_numerator / head_sum / 100,
            "Sum of Av $/hd": dollar_numerator / head_sum
        })

    pivot = df_filtered.groupby("Weight Range").apply(custom_aggregations).reset_index()
    for col in ["Sum of Av LW", "Sum of Av c/kg LW", "Sum of Av $/hd"]:
        pivot[col] = pivot[col].apply(lambda x: f"{x:,.2f}")

    pivot = pivot.sort_values(by="Weight Range")

    grand = custom_aggregations(df_filtered)
    grand["Weight Range"] = "Grand Total"
    for col in ["Sum of Av LW", "Sum of Av c/kg LW", "Sum of Av $/hd"]:
        grand[col] = f"{grand[col]:,.2f}"
    pivot = pd.concat([pivot, pd.DataFrame([grand])], ignore_index=True)

    st.dataframe(
        pivot,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Weight Range": st.column_config.TextColumn(label="Weight Range"),
            "Sum of Av LW": st.column_config.TextColumn(label="Sum of Av LW"),
            "Sum of Av c/kg LW": st.column_config.TextColumn(label="Sum of Av c/kg LW"),
            "Sum of Av $/hd": st.column_config.TextColumn(label="Sum of Av $/hd")
        }
    )

    # --- Export Immediately After Table ---
    with st.expander("📥 Download Excel File"):
        include_all_filters = st.checkbox("Include all filters (category, weight, prefix)", value=False)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            format_2dp = workbook.add_format({'num_format': '#,##0.00'})
            format_int = workbook.add_format({'num_format': '0'})
            format_default = workbook.add_format({})

            def autosize_columns(df, worksheet, formats={}):
                for idx, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    fmt = formats.get(col, format_default)
                    worksheet.set_column(idx, idx, max_len, fmt)

            pivot.to_excel(writer, sheet_name="Pivot Table", index=False)
            autosize_columns(pivot, writer.sheets["Pivot Table"])

            if include_all_filters:
                export_df = df_filtered.copy()
            else:
                export_df = df[
                    (df["Report Date"].isin(df_filtered["Report Date"])) &
                    (df["Saleyard"].isin(df_filtered["Saleyard"]))
                ].copy()

            export_df.to_excel(writer, sheet_name="All Data", index=False)
            autosize_columns(export_df, writer.sheets["All Data"], {
                col: format_2dp if any(k in col for k in ["c/kg", "$", "Avg"]) and col not in ["Head Count", "Head Change"]
                else format_int if col in ["Head Count", "Head Change"]
                else format_default for col in export_df.columns
            })

            for cat in export_df["Category"].dropna().unique():
                df_cat = export_df[export_df["Category"] == cat]
                sheet_name = cat[:31]
                df_cat.to_excel(writer, sheet_name=sheet_name, index=False)
                autosize_columns(df_cat, writer.sheets[sheet_name])

        st.download_button(
            label="⬇️ Download Now",
            data=output.getvalue(),
            file_name="boonz_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- Spacer Before Saleyard Reports Used ---
    st.markdown("### ")
    st.markdown("### ")

    
    # Apply all filters except Report Date range (for charts)
    df_chart_filtered = df.copy()

    # Apply saleyard filter
    if saleyards:
        df_chart_filtered = df_chart_filtered[df_chart_filtered["Saleyard"].isin(saleyards)]

    # Apply category filter
    if categories:
        df_chart_filtered = df_chart_filtered[df_chart_filtered["Category"].isin(categories)]

    # Apply weight range filter
    if weights:
        df_chart_filtered = df_chart_filtered[df_chart_filtered["Weight Range"].isin(weights)]

    # Apply prefix filter
    if prefixes:
        df_chart_filtered = df_chart_filtered[df_chart_filtered["Sale Prefix"].isin(prefixes)]

    # --- Weekly Chart Section ---
    st.subheader("7 Day Rolling Average")

    # Metric selection
    available_metrics = ["Sum of Av LW", "Sum of Av c/kg LW", "Sum of Av $/hd"]
    selected_metrics = [metric for metric in available_metrics if st.checkbox(metric, value=(metric == "Sum of Av $/hd"))]


    if selected_metrics:
        # Ensure we’re only using data from the filtered table
        chart_data = df_chart_filtered.copy()

        # Limit to data from 1 Jan 2024 onwards
        chart_data = chart_data[chart_data["Report Date"] >= datetime.datetime(2024, 1, 1)]


        # Group to weekly level
        chart_data["Week"] = chart_data["Report Date"].dt.to_period("W").apply(lambda r: r.start_time)

        for metric in selected_metrics:
            # Determine what raw column this metric derives from
            if metric == "Sum of Av LW":
                lw_numerator = ((chart_data["Avg $/Head"] / chart_data["Avg Lwt c/kg"]) * chart_data["Head Count"])
                chart_data["Rolling Numerator"] = lw_numerator * 100
                chart_data["Rolling Denominator"] = chart_data["Head Count"]

            elif metric == "Sum of Av c/kg LW":
                ckg_numerator = chart_data["Avg Lwt c/kg"] * chart_data["Head Count"]
                chart_data["Rolling Numerator"] = ckg_numerator / 100
                chart_data["Rolling Denominator"] = chart_data["Head Count"]

            elif metric == "Sum of Av $/hd":
                dollar_numerator = chart_data["Avg $/Head"] * chart_data["Head Count"]
                chart_data["Rolling Numerator"] = dollar_numerator
                chart_data["Rolling Denominator"] = chart_data["Head Count"]

            # Group and compute weekly weighted average
            grouped = chart_data.groupby(["Week", "Weight Range"])[["Rolling Numerator", "Rolling Denominator"]].sum().reset_index()
            grouped["Rolling Avg"] = grouped["Rolling Numerator"] / grouped["Rolling Denominator"]
            grouped = grouped.dropna(subset=["Rolling Avg"])

            grouped = grouped.dropna(subset=["Rolling Avg"])

            chart_title = metric

            st.markdown(f"**{chart_title} – Weekly Average by Weight Range**")
            chart_df = grouped.pivot(index="Week", columns="Weight Range", values="Rolling Avg").sort_index()
            st.line_chart(chart_df, use_container_width=True)

    st.markdown("### ")
    st.markdown("### ")
    
    
    # --- Used Saleyard Reports Summary ---
    st.subheader("📄 Saleyard Reports Used")
    used_reports = (
        df_filtered[["Saleyard", "Report Date"]]
        .dropna()
        .drop_duplicates()
        .sort_values(by=["Saleyard", "Report Date"])
    )
    used_reports["Report Date"] = used_reports["Report Date"].dt.strftime("%-d %B %Y")
    grouped = (
        used_reports.groupby("Saleyard")["Report Date"]
        .apply(lambda x: ", ".join(x))
        .reset_index()
        .rename(columns={"Report Date": "Report Dates"})
    )
    st.dataframe(grouped, hide_index=True, use_container_width=True)
