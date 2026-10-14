from __future__ import annotations



import io

import re

import zipfile

from datetime import datetime



import pandas as pd

import streamlit as st

from openpyxl.styles import Alignment, Font, PatternFill





st.set_page_config(page_title="Noon → Zoho Import Converter", page_icon="📦", layout="wide")



REQUIRED_COLUMNS = {

    "Reference Nr",
    "Transaction Type",

    "Order Nr",

    "Item Nr",

    "Transaction Date",

    "Title",

    "SKUs",

    "Net Proceeds",

}



INVOICE_COLUMNS = [

    "Invoice Date", "Invoice Number", "PurchaseOrder", "item name", "SKU",

    "Quantity", "Item Price", "VAT Treatment", "Accounts Receivable",

    "Is Inclusive Tax", "Template Name", "Currency Code", "Exchange Rate",

    "Account", "Item Tax", "Item Tax %", "Item Tax Type", "Item Type",

    "Noon Order Nr", "Reference Nr", "Customer Name", "Status",

]
def read_upload(uploaded_file) -> pd.DataFrame:

    suffix = uploaded_file.name.lower().rsplit(".", 1)[-1]

    if suffix == "csv":

        raw = uploaded_file.getvalue()

        last_error = None

        for encoding in ("utf-8-sig", "utf-8", "cp1252"):

            try:

                return pd.read_csv(io.BytesIO(raw), encoding=encoding)

            except UnicodeDecodeError as error:

                last_error = error

        raise ValueError(f"Could not read the CSV encoding: {last_error}")

    if suffix in {"xlsx", "xls"}:

        return pd.read_excel(io.BytesIO(uploaded_file.getvalue()))

    raise ValueError("Please upload a CSV, XLSX, or XLS file.")





def clean_source(data: pd.DataFrame) -> pd.DataFrame:

    data = data.copy()

    data.columns = data.columns.astype(str).str.strip()

    missing = sorted(REQUIRED_COLUMNS - set(data.columns))

    if missing:

        raise ValueError("Missing required column(s): " + ", ".join(missing))



    transaction_type = data["Transaction Type"].astype("string").str.strip().str.casefold()
    data = data[transaction_type.isin(["order", "order update"])].copy()
    if data.empty:
        raise ValueError("No rows were found with Transaction Type 'Order' or 'Order Update'.")

    data["Reference Nr"] = data["Reference Nr"].astype("string").str.strip()

    data = data[data["Reference Nr"].str.startswith("PS", na=False)].copy()

    if data.empty:

        raise ValueError("No rows were found where Reference Nr starts with PS.")



    data["Net Proceeds"] = pd.to_numeric(data["Net Proceeds"], errors="coerce")

    invalid_amounts = data["Net Proceeds"].isna()

    if invalid_amounts.any():

        rows = ", ".join(map(str, (data.index[invalid_amounts] + 2).tolist()[:20]))

        raise ValueError(f"Net Proceeds is not numerical on source row(s): {rows}")



    data["Transaction Date"] = pd.to_datetime(data["Transaction Date"], errors="coerce")

    invalid_dates = data["Transaction Date"].isna()

    if invalid_dates.any():

        rows = ", ".join(map(str, (data.index[invalid_dates] + 2).tolist()[:20]))

        raise ValueError(f"Transaction Date is invalid on source row(s): {rows}")



    for column in ("Order Nr", "Item Nr", "Title", "SKUs"):

        blank = data[column].isna() | data[column].astype(str).str.strip().eq("")

        if blank.any():

            rows = ", ".join(map(str, (data.index[blank] + 2).tolist()[:20]))

            raise ValueError(f"{column} is blank on source row(s): {rows}")

    return data





def build_invoice_rows(data: pd.DataFrame) -> pd.DataFrame:

    result = pd.DataFrame({

        "Invoice Date": data["Transaction Date"].astype(str),

        "Invoice Number": data["Item Nr"].astype(str),

        "PurchaseOrder": data["Item Nr"].astype(str),

        "item name": data["Title"],

        "SKU": data["SKUs"].astype(str),

        "Quantity": 1,

        "Item Price": data["Net Proceeds"],

        "VAT Treatment": "vat_not_registered",

        "Accounts Receivable": "Accounts Receivable",

        "Is Inclusive Tax": True,

        "Template Name": "Standard Template",

        "Currency Code": "AED",

        "Exchange Rate": 1,

        "Account": "Net Proceeds",

        "Item Tax": "Standard Rate",

        "Item Tax %": 5,

        "Item Tax Type": "ItemAmount",

        "Item Type": "goods",

        "Noon Order Nr": data["Order Nr"].astype(str),

        "Reference Nr": data["Reference Nr"].astype(str),

        "Customer Name": "Noon Receivables",

        "Location Name": "Noon Inventory",

        "Status": "Sent"

    })

    return result[INVOICE_COLUMNS]

def excel_bytes(data: pd.DataFrame, date_columns: list[str]) -> bytes:

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl", datetime_format="DD/MM/YYYY") as writer:

        data.to_excel(writer, index=False, sheet_name="Import")

        sheet = writer.book["Import"]

        sheet.freeze_panes = "A2"

        sheet.auto_filter.ref = sheet.dimensions

        header_fill = PatternFill("solid", fgColor="1F4E78")

        for cell in sheet[1]:

            cell.font = Font(bold=True, color="FFFFFF")

            cell.fill = header_fill

            cell.alignment = Alignment(horizontal="center")

        for cells in sheet.columns:

            values = ["" if cell.value is None else str(cell.value) for cell in cells]

            sheet.column_dimensions[cells[0].column_letter].width = min(max(max(map(len, values)) + 2, 12), 45)

        headers = {cell.value: cell.column for cell in sheet[1]}

        for name in date_columns:

            for row in range(2, sheet.max_row + 1):

                sheet.cell(row=row, column=headers[name]).number_format = "DD/MM/YYYY"

    return output.getvalue()





def safe_name(value: str) -> str:

    return re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._")





def reference_sort_key(reference: str):

    match = re.search(r"(\d{8})$", reference)

    if match:

        try:

            return datetime.strptime(match.group(1), "%Y%m%d")

        except ValueError:

            pass

    return datetime.min





st.title("Noon → Zoho Books Import Converter")

st.caption("Prepare invoice and sales-return import workbooks from a Noon Transaction View export.")



uploaded = st.file_uploader("Upload Noon Transaction View", type=["csv", "xlsx", "xls"])



if uploaded is None:

    st.info("Upload a Noon Transaction View file to begin.")

    st.stop()



try:

    source = clean_source(read_upload(uploaded))

except Exception as error:

    st.error(str(error))

    st.stop()



references = sorted(source["Reference Nr"].dropna().astype(str).unique(), key=reference_sort_key, reverse=True)

selected_reference = st.selectbox("Reference number", ["All"] + references, index=0)

if selected_reference == "All":

    selected = source.copy()

    export_name = "ALL_PS"

else:

    selected = source[source["Reference Nr"] == selected_reference].copy()

    export_name = safe_name(selected_reference)

positive = selected[selected["Net Proceeds"] > 0].copy()

negative = selected[selected["Net Proceeds"] < 0].copy()

zero_count = int((selected["Net Proceeds"] == 0).sum())



metric_cols = st.columns(4)

metric_cols[0].metric("Sales rows", len(positive))

metric_cols[1].metric("Return rows", len(negative))

metric_cols[2].metric("Zero rows discarded", zero_count)

metric_cols[3].metric("Positive net proceeds", f"AED {positive['Net Proceeds'].sum():,.2f}")



reference_name = export_name

files: dict[str, bytes] = {}

if not positive.empty:

    sales_name = f"invoice_import_{reference_name}_sales.xlsx"


    files[sales_name] = excel_bytes(build_invoice_rows(positive), ["Invoice Date"])


if not negative.empty:

    returns_name = f"invoice_import_{reference_name}_sales_returns.xlsx"

    files[returns_name] = excel_bytes(build_invoice_rows(negative), ["Invoice Date"])



st.subheader("Download Zoho import files")

if not files:

    st.warning("This reference contains only zero Net Proceeds rows, so there is nothing to export.")

else:

    button_cols = st.columns(min(len(files), 3))

    for index, (name, content) in enumerate(files.items()):

        button_cols[index % len(button_cols)].download_button(

            f"Download {name}", content, file_name=name,

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            use_container_width=True,

        )



    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:

        for name, content in files.items():

            archive.writestr(name, content)

    st.download_button(

        "Download all files as ZIP", zip_buffer.getvalue(),

        file_name=f"zoho_imports_{reference_name}.zip", mime="application/zip",

        type="primary", use_container_width=True,

    )



with st.expander("Preview selected rows"):

    st.dataframe(

        selected[["Transaction Date", "Reference Nr", "Order Nr", "Item Nr", "Title", "SKUs", "Net Proceeds"]],

        use_container_width=True,

        hide_index=True,

    )
