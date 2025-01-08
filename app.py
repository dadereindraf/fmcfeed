import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Fungsi untuk memproses data dari file txt
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
            now_size_condiion = parts[4]
            processed_data.append([table_name, date_transaction, date_availability, time_availability, now_size_condiion])

    df = pd.DataFrame(processed_data, columns=["TABLE NAME", "DATE TRANSACTION", "DATE AVAILABILITY", "TIME AVAILABILITY", "NOW SIZE CONDITION"])
    return df

# Fungsi untuk membuat workbook dengan sheet Main, Daily, Weekly, Monthly, dan Billing
def create_workbook(df):
    wb = Workbook()
    wb.remove(wb["Sheet"])  # Hapus sheet default

    # Tambahkan sheet utama
    sheets = {
        "Main": wb.create_sheet("Main"),
        "Daily": wb.create_sheet("Daily"),
        "Weekly": wb.create_sheet("Weekly"),
        "Monthly": wb.create_sheet("Monthly"),
        "Billing": wb.create_sheet("Billing"),
    }

    # Fungsi untuk menambahkan tabel ke sheet
    def add_table_to_sheet(ws, table_name, group):
        ws.append([f"TABLE NAME: {table_name}"])
        ws.append(["TABLE NAME", "DATE TRANSACTION", "DATE AVAILABILITY", "TIME AVAILABILITY","NOW SIZE CONDITION"])
        for row in group.values.tolist():
            ws.append(row)
        ws.append([])

    # Proses setiap tabel berdasarkan jumlah baris
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

    # Hapus teks "TABLE NAME: " dari kolom pertama di setiap sheet
    for sheet in sheets.values():
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
            if row[0].value and "TABLE NAME:" in str(row[0].value):
                row[0].value = row[0].value.replace("TABLE NAME: ", "")

    return wb

# Fungsi untuk memformat Excel dengan warna
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
        row = 1
        while row <= max_row:
            if sheet.cell(row=row, column=1).value:
                sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
                cell = sheet.cell(row=row, column=1)
                cell.fill = first_header_fill
                cell.font = header_font
                cell.alignment = header_alignment
                row += 1
                for col in range(1, 6):
                    cell = sheet.cell(row=row, column=col)
                    cell.fill = second_header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment
                row += 1
                while row <= max_row and sheet.cell(row=row, column=1).value:
                    for col in range(1, 6):
                        cell = sheet.cell(row=row, column=col)
                        cell.border = thin_border
                    row += 1
            else:
                row += 1
    return wb

# Fungsi untuk menyimpan workbook ke BytesIO
def save_workbook_to_bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit UI
st.title("📊 Data Processing and Excel Export with Colors")
st.write("Upload a `data.txt` file to process and download a formatted Excel file.")

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
            label="📥 Download Colored Excel File",
            data=excel_file,
            file_name="output_colored.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
