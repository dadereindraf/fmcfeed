import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook

# Fungsi untuk memproses data dari file txt
def process_data(file):
    data = file.decode("utf-8").splitlines()
    lines = [line.strip() for line in data if line.strip() and not line.startswith("||||")]

    processed_data = []
    for line in lines:
        parts = line.split("|")
        if len(parts) >= 5:
            event_date = parts[1]
            table_name = parts[0]
            start_date = parts[2]
            end_date = parts[3]
            value = parts[4]
            processed_data.append([table_name, event_date, start_date, end_date, value])

    df = pd.DataFrame(processed_data, columns=["TABLE NAME", "EVENT DATE", "DATE TRANSACTION", "DATE AVAILABILITY", "NOW SIZE CONDITION"])
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
        ws.append(["TABLE NAME", "DATE TRANSACTION", "DATE AVAILABILITY", "NOW SIZE CONDITION"])
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

# Fungsi untuk menyimpan workbook ke BytesIO
def save_workbook_to_bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit UI
st.title("Data Processing and Excel Export for Priority and Non Priority Table")
st.write("Upload a `data.txt` file to process and download it as an Excel file.")

uploaded_file = st.file_uploader("Choose a file", type="txt")
if uploaded_file is not None:
    df = process_data(uploaded_file.read())
    st.write("### Processed Data")
    st.dataframe(df)

    if st.button("Generate Excel File"):
        workbook = create_workbook(df)
        excel_file = save_workbook_to_bytes(workbook)
        st.download_button(
            label="📥 Download Excel File",
            data=excel_file,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
