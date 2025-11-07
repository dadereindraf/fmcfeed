import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- Fungsi bantu untuk parsing tanggal ---
def try_parse_date(value):
    """Coba ubah string ke datetime.date, kalau gagal return None."""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except:
            pass
    return None

# --- Fungsi utama memproses data ---
def process_data(file):
    data = file.decode("utf-8").splitlines()
    lines = [line.strip() for line in data if line.strip() and not line.startswith("||||")]

    processed_data = []
    for line in lines:
        parts = line.split("|")
        if len(parts) >= 5:
            table_name = parts[0]
            date_transaction = parts[1]
            date_availability = parts[2]
            time_availability = parts[3]
            now_size_condition = parts[4]

            # Tambahkan kolom baru
            sla_date = "D+1"
            processed_data.append([
                table_name, sla_date, date_transaction, date_availability,
                time_availability, now_size_condition
            ])

    df = pd.DataFrame(processed_data, columns=[
        "TABLE NAME", "SLA DATE", "DATE TRANSACTION", "DATE AVAILABILITY",
        "TIME AVAILABILITY", "NOW SIZE CONDITION"
    ])

    # --- Hitung kolom COMPLETENESS ---
    df["COMPLETENESS"] = df["TIME AVAILABILITY"].apply(
        lambda x: "NOT MET" if pd.isna(x) or str(x).strip() in ["", "-"] else "MET"
    )

    # --- Hitung kolom TIMELINESS ---
    def check_timeliness(row):
        date_trans = try_parse_date(row["DATE TRANSACTION"])
        date_avail = try_parse_date(row["DATE AVAILABILITY"])
        sla_val = row["SLA DATE"]
        sla_days = 1  # default D+1
        if isinstance(sla_val, str) and sla_val.startswith("D+"):
            try:
                sla_days = int(sla_val.replace("D+", ""))
            except:
                pass

        if not date_avail or str(row["DATE AVAILABILITY"]).strip() in ["", "-"]:
            return "NOT MET"
        if not date_trans:
            return "NOT MET"

        delta = (date_avail - date_trans).days
        return "NOT MET" if delta > sla_days else "MET"

    df["TIMELINESS"] = df.apply(check_timeliness, axis=1)

    # --- Hitung kolom NOTE ---
    def check_note(row):
        if row["TIMELINESS"] == "NOT MET":
            val = str(row["NOW SIZE CONDITION"]).strip()
            if val in ["", "-"]:
                return "Source Issue"
            else:
                return "Reprocess"
        return ""

    df["NOTE"] = df.apply(check_note, axis=1)

    return df

# --- Fungsi buat workbook ---
def create_workbook(df):
    wb = Workbook()
    wb.remove(wb["Sheet"])
    sheets = {
        "Main": wb.create_sheet("Main"),
        "Daily": wb.create_sheet("Daily"),
        "Weekly": wb.create_sheet("Weekly"),
        "Monthly": wb.create_sheet("Monthly"),
        "Billing": wb.create_sheet("Billing"),
    }

    def add_table_to_sheet(ws, table_name, group):
        ws.append([f"TABLE NAME: {table_name}"])
        ws.append(list(group.columns))
        for row in group.values.tolist():
            ws.append(row)
        ws.append([])

    for table_name, group in df.groupby("TABLE NAME"):
        if "bil" in table_name.lower() or "billing" in table_name.lower():
            add_table_to_sheet(sheets["Billing"], table_name, group)
        elif len(group) >= 10:
            add_table_to_sheet(sheets["Daily"], table_name, group)
        elif 1 < len(group) < 6:
            add_table_to_sheet(sheets["Weekly"], table_name, group)
        elif len(group) == 1:
            add_table_to_sheet(sheets["Monthly"], table_name, group)
        else:
            add_table_to_sheet(sheets["Main"], table_name, group)

    # Bersihkan teks prefix
    for sheet in sheets.values():
        for row in sheet.iter_rows():
            if row[0].value and "TABLE NAME:" in str(row[0].value):
                row[0].value = row[0].value.replace("TABLE NAME: ", "")

    return wb

# --- Format warna Excel ---
def format_excel_with_feeds(wb):
    first_header_fill = PatternFill(start_color="3C7D22", end_color="3C7D22", fill_type="solid")
    second_header_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        max_row = sheet.max_row
        max_col = 9

        for col_num, col_cells in enumerate(sheet.columns, start=1):
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col_cells)
            sheet.column_dimensions[get_column_letter(col_num)].width = max_length + 2

        row = 1
        while row <= max_row:
            if sheet.cell(row=row, column=1).value:
                sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
                cell = sheet.cell(row=row, column=1)
                cell.fill = first_header_fill
                cell.font = header_font
                cell.alignment = header_alignment
                row += 1
                for col in range(1, max_col + 1):
                    cell = sheet.cell(row=row, column=col)
                    cell.fill = second_header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment
                row += 1
                while row <= max_row and sheet.cell(row=row, column=1).value:
                    for col in range(1, max_col + 1):
                        cell = sheet.cell(row=row, column=col)
                        cell.border = thin_border
                    row += 1
            else:
                row += 1
    return wb

# --- Save ke BytesIO ---
def save_workbook_to_bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit UI ---
st.title("📊 Data Processing and Excel Export with Logic Evaluation")
st.write("Upload a `data.txt` file to process and see MET/NOT MET logic applied automatically.")

uploaded_file = st.file_uploader("Choose a file", type="txt")
if uploaded_file:
    df = process_data(uploaded_file.read())
    st.write("### Processed Data")
    st.dataframe(df)

    if st.button("Generate Excel File"):
        workbook = create_workbook(df)
        formatted_workbook = format_excel_with_feeds(workbook)
        excel_file = save_workbook_to_bytes(formatted_workbook)
        st.download_button(
            label="📥 Download Evaluated Excel File",
            data=excel_file,
            file_name="output_evaluated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )