import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# --- Config ---
st.set_page_config("MLA Saleyard Dashboard", layout="wide")

@st.cache_data(show_spinner=False)
def generate_excel_file(pivot, export_df_dict):
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

        all_data = export_df_dict["All Data"]
        all_data.to_excel(writer, sheet_name="All Data", index=False)
        autosize_columns(all_data, writer.sheets["All Data"], {
            col: format_2dp if any(k in col for k in ["c/kg", "$", "Avg"]) and col not in ["Head Count", "Head Change"]
            else format_int if col in ["Head Count", "Head Change"]
            else format_default for col in all_data.columns
        })

        for sheet, df_cat in export_df_dict["By Category"].items():
            df_cat.to_excel(writer, sheet_name=sheet[:31], index=False)
            autosize_columns(df_cat, writer.sheets[sheet[:31]])

    return output.getvalue()


# --- Load Data ---
@st.cache_data
def load_data():
    df = pd.read_excel("final_mla_output.xlsx")
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True).str.strip()
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



# --- Date Filter Mode Selection ---
date_mode = st.sidebar.radio("Select Date Filter", ["Preset Range", "Custom Range"], horizontal=True)

min_report_date = df["Report Date"].min()
max_report_date = df["Report Date"].max()

if date_mode == "Preset Range":
    preset_option = st.sidebar.selectbox("Preset Date Range", [
        "Last 7 Days", "Last 30 Days", "Last 60 Days", "Last 90 Days", "Last 180 Days"
    ])
    date_cutoffs = {
        "Last 7 Days": datetime.datetime.today() - datetime.timedelta(days=7),
        "Last 30 Days": datetime.datetime.today() - datetime.timedelta(days=30),
        "Last 60 Days": datetime.datetime.today() - datetime.timedelta(days=60),
        "Last 90 Days": datetime.datetime.today() - datetime.timedelta(days=90),
        "Last 180 Days": datetime.datetime.today() - datetime.timedelta(days=180)
    }
    df_filtered = df[df["Report Date"] >= date_cutoffs[preset_option]]

else:
    # --- Custom Date Range ---
    col_start, col_end = st.sidebar.columns(2)
    with col_start:
        custom_start = st.date_input("Start", min_value=min_report_date, max_value=max_report_date, value=min_report_date, key="custom_start")
    with col_end:
        custom_end = st.date_input("End", min_value=min_report_date, max_value=max_report_date, value=max_report_date, key="custom_end")

    df_filtered = df[
        (df["Report Date"] >= pd.to_datetime(custom_start)) &
        (df["Report Date"] <= pd.to_datetime(custom_end))
    ]




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




tab1, tab2, tab3, tab4 = st.tabs(["📊 Main View", "📈 Charts", "🧩 Grids", "📄 Used Reports"])

with tab1:
    st.header("The Yarding Data")

    if df_filtered.empty:
        st.warning("No data matches your filters.")
    else:
        def custom_aggregations(group):
            head_sum = group["Head Count"].sum()
            
            if head_sum == 0 or "Avg Cwt c/kg" not in group.columns:
                return pd.Series({
                    "Head Count": head_sum,
                    "Average LW": 0,
                    "Average c/kg LW": 0,
                    "Average c/kg CW": 0,
                    "Average $/hd": 0
                })

            lw_numerator = ((group["Avg $/Head"] / group["Avg Lwt c/kg"]) * group["Head Count"]).sum()
            ckg_numerator = (group["Avg Lwt c/kg"] * group["Head Count"]).sum()
            cwt_numerator = (group["Avg Cwt c/kg"] * group["Head Count"]).sum()
            dollar_numerator = (group["Avg $/Head"] * group["Head Count"]).sum()

            return pd.Series({
                "Head Count": head_sum,
                "Average LW": (lw_numerator / head_sum) * 100,
                "Average c/kg LW": ckg_numerator / head_sum,
                "Average c/kg CW": cwt_numerator / head_sum,
                "Average $/hd": dollar_numerator / head_sum
            })


        pivot = df_filtered.groupby("Weight Range").apply(custom_aggregations).reset_index()
        for col in ["Average LW", "Average c/kg LW", "Average c/kg CW", "Average $/hd"]:
            pivot[col] = pivot[col].apply(lambda x: f"{x:,.2f}")
        pivot["Head Count"] = pivot["Head Count"].apply(lambda x: f"{int(x):,}")


        pivot = pivot.sort_values(by="Weight Range")

        grand = custom_aggregations(df_filtered)
        grand["Weight Range"] = "Grand Total"
        grand["Head Count"] = f"{int(grand['Head Count']):,}"
        grand["Average LW"] = f"{grand['Average LW']:,.2f}"
        grand["Average c/kg LW"] = f"{grand['Average c/kg LW']:,.2f}"
        grand["Average c/kg CW"] = f"{grand['Average c/kg CW']:,.2f}"
        grand["Average $/hd"] = f"{grand['Average $/hd']:,.2f}"


        pivot = pd.concat([pivot, pd.DataFrame([grand])], ignore_index=True)

        st.dataframe(
            pivot,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Weight Range": st.column_config.TextColumn(label="Weight Range"),
                "Head Count": st.column_config.TextColumn(label="Head Count"),
                "Average LW": st.column_config.TextColumn(label="Average LW"),
                "Average c/kg LW": st.column_config.TextColumn(label="Average c/kg LW"),
                "Average c/kg CW": st.column_config.TextColumn(label="Average c/kg CW"),
                "Average $/hd": st.column_config.TextColumn(label="Average $/hd")
            }
        )




        for yard in sorted(df_filtered['Saleyard'].dropna().unique()):
            with st.expander(f"📍 {yard} — Data"):
                yard_data = df_filtered[df_filtered["Saleyard"] == yard]
                yard_pivot = yard_data.groupby("Weight Range").apply(custom_aggregations).reset_index()

                # --- Add Grand Total Row ---
                yard_grand = custom_aggregations(yard_data)
                yard_grand["Weight Range"] = "Grand Total"
                yard_pivot = pd.concat([yard_pivot, pd.DataFrame([yard_grand])], ignore_index=True)

                # --- Format Columns Consistently ---
                yard_pivot["Head Count"] = yard_pivot["Head Count"].apply(lambda x: f"{int(x):,}")
                for col in ["Average LW", "Average c/kg LW", "Average c/kg CW", "Average $/hd"]:
                    yard_pivot[col] = yard_pivot[col].apply(lambda x: f"{x:,.2f}")

                st.dataframe(
                    yard_pivot,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Weight Range": st.column_config.TextColumn(label="Weight Range"),
                        "Head Count": st.column_config.TextColumn(label="Head Count"),
                        "Average LW": st.column_config.TextColumn(label="Average LW"),
                        "Average c/kg LW": st.column_config.TextColumn(label="Average c/kg LW"),
                        "Average c/kg CW": st.column_config.TextColumn(label="Average c/kg CW"),
                        "Average $/hd": st.column_config.TextColumn(label="Average $/hd")
                    }
                )




        
        

        
        


    # Check if pivot exists before rendering export tools
    if 'pivot' in locals() and not pivot.empty:
        with st.sidebar.expander("📥 Download Excel"):
            include_all_filters = st.checkbox("Filter with category, weight, prefix", value=False, key="sidebar_export_filters")

            if include_all_filters:
                export_df = df_filtered.copy()
            else:
                export_df = df[
                    (df["Report Date"].isin(df_filtered["Report Date"])) &
                    (df["Saleyard"].isin(df_filtered["Saleyard"]))
                ].copy()

            category_dfs = {
                cat: export_df[export_df["Category"] == cat] for cat in export_df["Category"].dropna().unique()
            }

            excel_data = generate_excel_file(pivot, {
                "All Data": export_df,
                "By Category": category_dfs
            })

            st.download_button(
                label="⬇️ Download Excel",
                data=excel_data,
                file_name="boonz_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.sidebar.caption("ℹ️ Export will be available once data is loaded.")



with tab2:
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
        st.subheader("Charts")
        min_c_date = df["Report Date"].min()
        max_c_date = df["Report Date"].max()

        col_start, col_end = st.columns(2)
        with col_start:
            chart_start = st.date_input("📊 Chart Start", value=min_c_date, min_value=min_c_date, max_value=max_c_date, key="chart_start")
        with col_end:
            chart_end = st.date_input("📊 Chart End", value=max_c_date, min_value=min_c_date, max_value=max_c_date, key="chart_end")







        # Convert to datetime for filtering
        chart_start = pd.to_datetime(chart_start)
        chart_end = pd.to_datetime(chart_end)

        # Apply to chart data
        df_chart_filtered = df_chart_filtered[
            (df_chart_filtered["Report Date"] >= chart_start) &
            (df_chart_filtered["Report Date"] <= chart_end)
        ]


        # Metric selection
        available_metrics = ["Average LW", "Average c/kg LW", "Average $/hd"]
        selected_metrics = [metric for metric in available_metrics if st.checkbox(metric, value=(metric == "Average $/hd"))]


        if selected_metrics:
            chart_data = df_chart_filtered.copy()





            # Group to weekly level
            chart_data["Week"] = chart_data["Report Date"].dt.to_period("W").apply(lambda r: r.start_time)

            for metric in selected_metrics:
                # Determine what raw column this metric derives from
                if metric == "Average LW":
                    lw_numerator = ((chart_data["Avg $/Head"] / chart_data["Avg Lwt c/kg"]) * chart_data["Head Count"])
                    chart_data["Rolling Numerator"] = lw_numerator * 100
                    chart_data["Rolling Denominator"] = chart_data["Head Count"]

                elif metric == "Average c/kg LW":
                    ckg_numerator = chart_data["Avg Lwt c/kg"] * chart_data["Head Count"]
                    chart_data["Rolling Numerator"] = ckg_numerator
                    chart_data["Rolling Denominator"] = chart_data["Head Count"]

                elif metric == "Average $/hd":
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
                import altair as alt

                chart_df_reset = chart_df.reset_index().melt(id_vars="Week", var_name="Weight Range", value_name="Value")

                line_chart = alt.Chart(chart_df_reset).mark_line().encode(
                    x=alt.X("Week:T", axis=alt.Axis(format="%b %Y", title="")),
                    y=alt.Y("Value:Q", title=chart_title),
                    color="Weight Range:N"
                ).properties(
                    width="container",
                    height=400
                ).interactive()

                st.altair_chart(line_chart, use_container_width=True)    


with tab3:
    st.subheader("🧩 Grid Placeholder")
    st.write("This will house structured grids or detailed table views.")
    st.dataframe(pd.DataFrame({
        "Column A": ["Row 1", "Row 2", "Row 3"],
        "Column B": [10, 20, 30]
    }), use_container_width=True)


with tab4:
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

