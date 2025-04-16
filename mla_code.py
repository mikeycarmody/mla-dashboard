import os
import shutil
import pandas as pd
from datetime import datetime

# --- Paths ---
folder_path = "/home/mikey/MLA/downloads"
history_path = "/home/mikey/MLA/history"
output_csv = "/home/mikey/MLA/combined_mla_output.csv"
output_excel = "/home/mikey/MLA/final_mla_output.xlsx"
output_rows = []

print("🚀 Starting MLA merge process...")
os.makedirs(history_path, exist_ok=True)

# --- Step 1: Process each MLA file ---
for filename in os.listdir(folder_path):
    if not filename.endswith(".csv") or filename.startswith("combined"):
        continue

    file_path = os.path.join(folder_path, filename)
    print(f"\n📄 Processing file: {filename}")

    # --- Parse filename to get saleyard and date ---
    if len(filename) > 15:
        base_name = filename[:-4]  # Strip .csv
        date_str_raw = base_name[-10:]  # Expecting DD-MM-YYYY
        try:
            report_date_str = datetime.strptime(date_str_raw, "%d-%m-%Y").strftime("%d/%m/%Y")
        except ValueError:
            print(f"⚠️ Invalid date format in filename, skipping: {filename}")
            continue

        saleyard_guess = base_name[:-11].replace("_", " ").strip()
        saleyard = saleyard_guess
        print(f"📌 Parsed from filename → Saleyard: {saleyard}, Report Date: {report_date_str}")
    else:
        print(f"⚠️ Filename too short or malformed: {filename}")
        continue

    # --- Read file ---
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    header = None
    data_started = False

    for i, line in enumerate(lines):
        if line.strip().startswith("Category,"):
            header = line.strip().split(",")
            data_start_index = i + 1
            data_started = True
            break

    if not data_started:
        print(f"⚠️ Skipping file {filename} — no data table found.")
        continue

    # --- Extract data rows ---
    data_rows = []
    for line in lines[data_start_index:]:
        if line.strip() == "":
            break
        fields = line.strip().split(",")
        if len(fields) >= len(header):
            data_rows.append(fields[:len(header)])

    if not data_rows:
        print(f"⚠️ No data rows found in {filename}. Skipping.")
        continue

    df = pd.DataFrame(data_rows, columns=header)

    # --- Keep only relevant columns ---
    keep_cols = [
        "Category", "Weight Range", "Sale Prefix", "Head Count", "Head Change",
        "Min Lwt c/kg", "Max Lwt c/kg", "Avg Lwt c/kg", "Avg Lwt Change",
        "Min $/Head", "Max $/Head", "Avg $/Head"
    ]
    df = df[[col for col in keep_cols if col in df.columns]].copy()
    df["Saleyard"] = saleyard
    df["Report Date"] = report_date_str
    output_rows.append(df)

    # --- Move file to history ---
    shutil.move(file_path, os.path.join(history_path, filename))
    print(f"📦 Moved to history: {filename}")

# --- Step 2: Save combined CSV ---
if output_rows:
    final_df = pd.concat(output_rows, ignore_index=True)

    # Merge with existing Excel data if it exists
    if os.path.exists(output_excel):
        existing = pd.read_excel(output_excel, sheet_name="All Data")
        final_df = pd.concat([existing, final_df], ignore_index=True).drop_duplicates()


    # Reorder columns if needed
    col_order = ["Saleyard", "Report Date"] + [col for col in final_df.columns if col not in ("Saleyard", "Report Date")]
    final_df = final_df[col_order]

    # --- Merge with previous combined output if it exists ---
    if os.path.exists(output_csv):
        prev_df = pd.read_csv(output_csv)
        final_df = pd.concat([prev_df, final_df], ignore_index=True).drop_duplicates()
    final_df.to_csv(output_csv, index=False)
    print(f"\n✅ CSV saved: {output_csv}")

    # Convert numeric fields
    numeric_cols = [
        "Head Count", "Head Change",
        "Min Lwt c/kg", "Max Lwt c/kg", "Avg Lwt c/kg", "Avg Lwt Change",
        "Min $/Head", "Max $/Head", "Avg $/Head"
    ]
    for col in numeric_cols:
        if col in final_df.columns:
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

    # --- Step 3: Save Excel with formatting ---
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, sheet_name="All Data", index=False)
        workbook = writer.book
        format_2dp = workbook.add_format({'num_format': '#,##0.00'})

        def format_sheet(df, sheet_name):
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                col_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, col_width, format_2dp if pd.api.types.is_numeric_dtype(df[col]) else None)

        # Format "All Data" sheet
        format_sheet(final_df, "All Data")

        # Format per-category sheets
        for category in final_df["Category"].dropna().unique():
            safe_sheet = category[:31]
            df_cat = final_df[final_df["Category"] == category]
            format_sheet(df_cat, safe_sheet)

    print(f"✅ Excel with formatting saved: {output_excel}")
else:
    print("❌ No usable data found to merge.")
