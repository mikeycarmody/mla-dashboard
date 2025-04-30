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

# --- Sidebar Filters ---
st.sidebar.header("🔍 Filters")

# --- Date Filter Mode Selection ---
date_mode = st.sidebar.radio("Select Date Filter", ["Preset Range", "Custom Range"], horizontal=True)

min_report_date = df["Report Date"].min()
max_report_date = df["Report Date"].max()

# Start fresh
df_filtered = df.copy()

# --- Apply Date Filters First ---
if date_mode == "Preset Range":
    preset_option = st.sidebar.selectbox("Preset Date Range", [
        "Last Sale", "Last 7 Days", "Last 30 Days", "Last 60 Days", "Last 90 Days", "Last 180 Days"
    ])

    date_cutoffs = {
        "Last 7 Days": datetime.datetime.today() - datetime.timedelta(days=7),
        "Last 30 Days": datetime.datetime.today() - datetime.timedelta(days=30),
        "Last 60 Days": datetime.datetime.today() - datetime.timedelta(days=60),
        "Last 90 Days": datetime.datetime.today() - datetime.timedelta(days=90),
        "Last 180 Days": datetime.datetime.today() - datetime.timedelta(days=180)
    }
    if preset_option == "Last Sale":
        # Keep only the latest Report Date per Saleyard
        last_sale_dates = df_filtered.groupby("Saleyard")["Report Date"].max().reset_index()
        df_filtered = pd.merge(df_filtered, last_sale_dates, on=["Saleyard", "Report Date"], how="inner")
    else:
        df_filtered = df_filtered[df_filtered["Report Date"] >= date_cutoffs[preset_option]]


else:
    # --- Custom Date Range ---
    col_start, col_end = st.sidebar.columns(2)
    with col_start:
        custom_start = st.date_input("Start", min_value=min_report_date, max_value=max_report_date, value=min_report_date, key="custom_start")
    with col_end:
        custom_end = st.date_input("End", min_value=min_report_date, max_value=max_report_date, value=max_report_date, key="custom_end")

    df_filtered = df_filtered[
        (df_filtered["Report Date"] >= pd.to_datetime(custom_start)) &
        (df_filtered["Report Date"] <= pd.to_datetime(custom_end))
    ]



comparison_range = st.sidebar.selectbox("Select Comparison Range", [
    "Last 30 Days", "Last 60 Days", "Last 90 Days", "Last 180 Days"
], key="comparison_range")

comparison_cutoffs = {
    "Last 30 Days": datetime.datetime.today() - datetime.timedelta(days=30),
    "Last 60 Days": datetime.datetime.today() - datetime.timedelta(days=60),
    "Last 90 Days": datetime.datetime.today() - datetime.timedelta(days=90),
    "Last 180 Days": datetime.datetime.today() - datetime.timedelta(days=180)
}

# Comparison start date based on current selection
comparison_start = comparison_cutoffs[comparison_range]
comparison_end = datetime.datetime.today()



 # --- Saleyard Filter ---
#saleyard_options = sorted(set(df_filtered["Saleyard"].dropna().unique()).union(favourites))
#default_selection = [f for f in favourites if f in saleyard_options]
#fav_selected = None

#saleyards = st.sidebar.multiselect("Saleyards", saleyard_options, default=default_selection)
#if saleyards:
    #df_filtered = df_filtered[df_filtered["Saleyard"].isin(saleyards)]
    
    
    
# --- Saleyard Filter with State Selection and Favourites ---

# --- Saleyard Filter with State Toggle + Sticky Selections ---

# Hardcoded saleyard-to-state map
saleyard_state_map = {
    # NSW
    "Armidale": "NSW", "Bathurst (Closed 2008)": "NSW", "Bombala": "NSW", "Casino": "NSW",
    "Coonamble": "NSW", "CTLX Carcoar": "NSW", "Dubbo": "NSW", "Finley": "NSW",
    "Forbes": "NSW", "Glen Innes": "NSW", "Goulburn": "NSW", "Griffith": "NSW",
    "Gunnedah": "NSW", "HRLX Singleton": "NSW", "IRLX Inverell": "NSW", "Moss Vale": "NSW",
    "Scone": "NSW", "SELX Yass": "NSW", "Silverdale": "NSW", "TRLX Tamworth": "NSW", "Wagga": "NSW",

    # VIC
    "Bairnsdale": "VIC", "Bendigo": "VIC", "Camperdown": "VIC", "Colac": "VIC",
    "CVLX Ballarat": "VIC", "Dandenong (Closed 1999)": "VIC", "Echuca": "VIC",
    "Korumburra": "VIC", "Leongatha": "VIC", "Mortlake": "VIC", "NVLX Barnawartha": "VIC",
    "NVLX Wodonga": "VIC", "Pakenham": "VIC", "Shepparton": "VIC", "Swan Hill": "VIC",
    "Warrnambool": "VIC", "Wycheproof": "VIC", "Yea": "VIC",

    # QLD
    "Blackall": "QLD", "Charters Towers": "QLD", "CQLX Gracemere": "QLD", "Dalby": "QLD",
    "Emerald": "QLD", "Longreach": "QLD", "Mareeba": "QLD", "Monto": "QLD",
    "Moreton": "QLD", "Murgon": "QLD", "Oakey": "QLD", "Roma Prime": "QLD",
    "Roma Store": "QLD", "Toowoomba": "QLD", "Wandoan": "QLD", "Warwick": "QLD",

    # SA
    "Gepps Cross (Closed 2002)": "SA", "Millicent": "SA", "Mount Gambier": "SA",
    "Mt Compass": "SA", "Naracoorte": "SA", "SA Livestock Exchange": "SA",

    # WA
    "Boyanup": "WA", "Midland": "WA", "Mount Barker": "WA", "Muchea": "WA",

    # TAS
    "Killafaddy": "TAS", "Powranna": "TAS",

    # NATIONAL
    "National AuctionsPlus online": "NATIONAL"
}

# --- State Selector ---
selected_state = st.sidebar.radio("📍 Filter Saleyards by State", ["ALL", "NSW", "VIC", "QLD"], horizontal=True)

# All available saleyards (including favourites)
saleyard_options_full = sorted(set(df["Saleyard"].dropna().unique()).union(favourites))

# Filter dropdown options by state
if selected_state == "ALL":
    visible_saleyards = saleyard_options_full
else:
    visible_saleyards = sorted([s for s in saleyard_options_full if saleyard_state_map.get(s) == selected_state])

# Previously selected saleyards (sticky)
current_selected = st.session_state.get("saleyard", [])
preserve_selected = [s for s in current_selected if s in saleyard_options_full]

# Merge preserved selections + visible ones for full dropdown (removes dups)
dropdown_options = sorted(set(visible_saleyards).union(preserve_selected))

# Auto-select favourites only on first load
default_selection = preserve_selected or [f for f in favourites if f in dropdown_options]

# Render dropdown
saleyards = st.sidebar.multiselect("Saleyards", dropdown_options, default=default_selection, key="saleyard")

# Filter the data
if saleyards:
    df_filtered = df_filtered[df_filtered["Saleyard"].isin(saleyards)]



# --- Dynamic Filters ---

# Category
filtered_categories = sorted(df_filtered["Category"].dropna().unique())
category_default = [c for c in st.session_state.get("category", []) if c in filtered_categories]
categories = st.sidebar.multiselect("Category", filtered_categories, default=category_default, key="category")
if categories:
    df_filtered = df_filtered[df_filtered["Category"].isin(categories)]

# Weight Range
filtered_weights = sorted(df_filtered["Weight Range"].dropna().unique())
weight_default = [w for w in st.session_state.get("weight", []) if w in filtered_weights]
weights = st.sidebar.multiselect("Weight Range", filtered_weights, default=weight_default, key="weight")
if weights:
    df_filtered = df_filtered[df_filtered["Weight Range"].isin(weights)]

# Sale Prefix
filtered_prefixes = sorted(df_filtered["Sale Prefix"].dropna().unique())
prefix_default = [p for p in st.session_state.get("prefix", []) if p in filtered_prefixes]
prefixes = st.sidebar.multiselect("Sale Prefix", filtered_prefixes, default=prefix_default, key="prefix")
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


        # --- Calculate main pivot ---
        pivot = df_filtered.groupby("Weight Range").apply(custom_aggregations).reset_index()

        # --- Get comparison pivot ---
        df_comparison = df[
            (df["Report Date"] >= comparison_start) &
            (df["Report Date"] <= comparison_end)
        ]

        # Match filters (same saleyards etc)
        if saleyards:
            df_comparison = df_comparison[df_comparison["Saleyard"].isin(saleyards)]
        if categories:
            df_comparison = df_comparison[df_comparison["Category"].isin(categories)]
        if weights:
            df_comparison = df_comparison[df_comparison["Weight Range"].isin(weights)]
        if prefixes:
            df_comparison = df_comparison[df_comparison["Sale Prefix"].isin(prefixes)]

        pivot_comparison = df_comparison.groupby("Weight Range").apply(custom_aggregations).reset_index()


        # Skip formatting if already formatted
        def format_int_safe(val):
            try:
                return f"{int(float(str(val).replace(',', ''))):,}"
            except:
                return val  # already formatted

        pivot["Head Count"] = pivot["Head Count"].apply(format_int_safe)



        pivot = pivot.sort_values(by="Weight Range")

        grand = custom_aggregations(df_filtered)
        grand["Weight Range"] = "Grand Total"
        grand["Head Count"] = f"{int(grand['Head Count']):,}"
        grand["Average LW"] = f"{grand['Average LW']:,.2f}"
        grand["Average c/kg LW"] = f"{grand['Average c/kg LW']:,.2f}"
        grand["Average c/kg CW"] = f"{grand['Average c/kg CW']:,.2f}"
        grand["Average $/hd"] = f"{grand['Average $/hd']:,.2f}"


        pivot = pd.concat([pivot, pd.DataFrame([grand])], ignore_index=True)
        
        def format_with_change(current, previous, metric):
            try:
                # Ensure current and previous are floats
                current_val = float(str(current).replace(",", ""))
                previous_val = float(str(previous).replace(",", ""))
            except:
                return current  # fallback if conversion fails

            if previous_val == 0 or pd.isna(previous_val):
                if metric in ["Head Count", "Average LW"]:
                    return f"{current_val:,.0f}"
                else:
                    return f"{current_val:,.2f}"
            
            pct_change = ((current_val - previous_val) / previous_val) * 100
            arrow = "🟢" if pct_change > 0 else "🔴"
            em_space = " " * 4
            if metric in ["Head Count", "Average LW"]:
                return f"{current_val:,.0f}{em_space}{arrow} {abs(pct_change):.1f}%"
            else:
                return f"{current_val:,.2f}{em_space}{arrow} {abs(pct_change):.1f}%"


                # Convert pivot_comparison to dict for quick lookup
        comp_dict = pivot_comparison.set_index("Weight Range").to_dict(orient="index")
        
        # Inject comparisons including Head Count
        for i, row in pivot.iterrows():
            wr = row["Weight Range"]
            comp_row = comp_dict.get(wr, {})
            
            for metric in ["Head Count", "Average LW", "Average c/kg LW", "Average c/kg CW", "Average $/hd"]:
                curr = row[metric]
                prev = comp_row.get(metric)
                pivot.at[i, metric] = format_with_change(curr, prev, metric)

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

