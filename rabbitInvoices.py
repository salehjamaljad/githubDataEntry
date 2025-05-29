import streamlit as st
import pandas as pd
import zipfile
import io
from datetime import datetime
import xlsxwriter
from streamlit_gsheets import GSheetsConnection
def rabbitInvoices():
    st.title("Rabbit & Khateer Processor")
    conn = st.connection("gsheets", type=GSheetsConnection)    
    
    df_invoice_number = conn.read(worksheet="Saved", cell="A1", ttl=5, headers=False)
    default_base_invoice_num = int(df_invoice_number.iat[0, 0])
    if "base_invoice_num" not in st.session_state:
        st.session_state.base_invoice_num = default_base_invoice_num
    base_invoice_num = st.number_input("رقم الفاتورة الأساسي", value=st.session_state.base_invoice_num, step=1)
    uploaded_zip = st.file_uploader("Upload a ZIP file containing Excel files", type="zip")

    if uploaded_zip:
        with zipfile.ZipFile(uploaded_zip) as zip_ref:
            output_zip_io = io.BytesIO()
            with zipfile.ZipFile(output_zip_io, "w") as output_zip:
                khateer_data = []
                khodar_data = []

                for file_index, file_name in enumerate(zip_ref.namelist()):
                    if not file_name.endswith(".xlsx") or file_name.startswith("__MACOSX"):
                        continue

                    with zip_ref.open(file_name) as file:
                        try:
                            df = pd.read_excel(file, skiprows=8)
                            file.seek(0)
                            df2 = pd.read_excel(file)

                            df = df[:-9].reset_index(drop=True)

                            branch = str(df2.iloc[1, 1]).strip()
                            order_number = int(df2.iloc[2, 6])
                            invoice_total = df2.iloc[-9, -1]
                            delivery_date = pd.to_datetime(
                                df2.iloc[1, 6], errors="coerce"
                            ).strftime("%Y-%m-%d")

                            name_lc = file_name.lower()
                            prefix = ""
                            if "khateer" in name_lc:
                                prefix = "خطير"
                            elif "khodar" in name_lc:
                                prefix = "رابيت"

                            parts = filter(None, [prefix, branch, delivery_date])
                            output_filename = "_".join(parts) + ".xlsx"

                            invoice_number = base_invoice_num + file_index  # <<< Here is the new logic

                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                                df.to_excel(writer, index=False, sheet_name="Data")
                                workbook = writer.book
                                invoice_ws = workbook.add_worksheet("فاتورة")

                                meta_format = workbook.add_format({'bold': True, 'border': 2})
                                bold_border_right = workbook.add_format({'bold': True, 'border': 2})
                                bold_center = workbook.add_format({'bold': True, 'align': 'center'})
                                bold_merge = workbook.add_format({'bold': True, 'border': 2, 'align': 'center', 'valign': 'vcenter'})
                                headers_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})

                                # Optional image insertion - if you want to keep it, otherwise comment out
                                try:
                                    invoice_ws.insert_image("A1", "Picture1.png", {'x_scale': 1.5, 'y_scale': 1})
                                except:
                                    pass  # continue silently if no image

                                # Company info
                                invoice_ws.write("A5", "شركه خضار للتجارة والتسويق", meta_format)
                                centered_meta_format = workbook.add_format({'bold': True, 'border': 2, 'align': 'center', 'valign': 'vcenter'})

                                invoice_ws.merge_range("B1:C1", "شركه خضار للتجارة والتسويق", centered_meta_format)
                                invoice_ws.merge_range("B2:C2", "Khodar for Trading & Marketing", centered_meta_format)


                                # Invoice metadata labels
                                invoice_ws.write("F1", "فاتورة مبيعات", meta_format)
                                invoice_ws.write("F2", "رقم الفاتورة #", meta_format)
                                invoice_ws.write("F3", "تاريخ الاستلام", meta_format)
                                invoice_ws.write("F4", "امر شراء رقم", meta_format)
                                invoice_ws.write("F6", "اسم العميل", meta_format)
                                invoice_ws.write("F7", "الفرع", meta_format)

                                client_prefix = "رابيت" if "khodar" in name_lc else "خطير"
                                invoice_ws.write("E2", invoice_number, meta_format)
                                invoice_ws.write("E3", delivery_date, meta_format)
                                invoice_ws.write("E4", str(order_number), workbook.add_format({'border': 2, 'align': 'center', 'bold':True}))
                                invoice_ws.write("E6", f"{prefix} - فرع {branch}", meta_format)
                                invoice_ws.write("E7", branch, meta_format)

                                # Invoice table headers at row 11 (zero-indexed: row 10)
                                invoice_ws.write("A11", "Barcode", headers_format)
                                invoice_ws.write("B11", "Arabic Product Name", headers_format)
                                invoice_ws.write("C11", "Unit Cost", headers_format)
                                invoice_ws.write("D11", "quantity", headers_format)
                                invoice_ws.write("E11", "total", headers_format)
                                # Define formats
                                border_format = workbook.add_format({'border': 1})
                                barcode_format = workbook.add_format({'num_format': '0', 'border': 1})
                                qty_total_format = workbook.add_format({'border': 1})  # Full border for quantity and total
                                # Write rows, empty qty and total
                                for idx, row in df.iterrows():
                                    row_num = 11 + idx
                                    barcode_value = row.get("Barcode", "")
                                    if pd.isna(barcode_value) or barcode_value == '':
                                        invoice_ws.write_blank(row_num, 0, "", border_format)
                                    else:
                                        try:
                                            barcode_int = int(barcode_value)
                                            invoice_ws.write_number(row_num, 0, barcode_int, barcode_format)
                                        except Exception:
                                            invoice_ws.write_string(row_num, 0, str(barcode_value), border_format)

                                    invoice_ws.write(row_num, 1, row.get("Arabic Product Name", ""), border_format)
                                    invoice_ws.write(row_num, 2, row.get("Unit Cost", ""), border_format)
                                    invoice_ws.write(row_num, 3, "", qty_total_format)  # quantity with full border
                                    invoice_ws.write(row_num, 4, "", qty_total_format)  # total with full border

                                last_row = 11 + len(df)
                                # Column width adjustments
                                invoice_ws.set_column("A:A", 25)  # ~235px
                                invoice_ws.set_column("B:B", 30)  # ~275px
                                invoice_ws.set_column("C:C", 15)  # ~145px
                                invoice_ws.set_column("D:D", 15)  # ~145px
                                invoice_ws.set_column("E:E", 27)  # ~250px
                                invoice_ws.set_column("F:F", 11)  # ~105px
                                invoice_ws.merge_range(last_row, 0, last_row, 3, "Subtotal", bold_merge)
                                invoice_ws.write_blank(last_row, 4, "", bold_border_right)

                                invoice_ws.merge_range(last_row + 1, 0, last_row + 1, 3, "Total", bold_merge)
                                invoice_ws.write(last_row + 1, 4, invoice_total, bold_border_right)


                                footer_start = last_row + 3
                                footer_texts = [
                                    "شركة خضار للتجارة و التسويق",
                                    "ش.ذ.م.م",
                                    "سجل تجارى / 13138  بطاقه ضريبية/721/294/448"
                                ]
                                for i, text in enumerate(footer_texts):
                                    row = footer_start + i
                                    invoice_ws.merge_range(row, 0, row, 3, text, bold_center)

                                invoice_ws.set_column("A:A", 25)
                                invoice_ws.set_column("B:B", 30)
                                invoice_ws.set_column("C:E", 15)

                            # After exiting the 'with pd.ExcelWriter' block:
                            excel_buffer.seek(0)
                            output_zip.writestr(output_filename, excel_buffer.getvalue())

                            # Prepare for pivot
                            pivot_cols = ["SKU", "Barcode", "Arabic Product Name", "Unit Cost", "Total PC"]
                            if all(col in df.columns for col in pivot_cols):
                                pivot_df = df[pivot_cols].copy()
                                pivot_df.rename(columns={"Total PC": branch}, inplace=True)

                                if "khateer" in name_lc:
                                    khateer_data.append(pivot_df)
                                elif "khodar" in name_lc:
                                    khodar_data.append(pivot_df)

                        except Exception as e:
                            st.warning(f"Failed to process {file_name}: {e}")

                # Merge and process pivoted data
                def create_aggregated_df(list_of_dfs):
                    if not list_of_dfs:
                        return None

                    merged_df = list_of_dfs[0]
                    for df in list_of_dfs[1:]:
                        merged_df = pd.merge(merged_df, df, on=["SKU", "Barcode", "Arabic Product Name", "Unit Cost"], how="outer")

                    # Identify and sort branch columns alphabetically
                    branch_cols = sorted([col for col in merged_df.columns if col not in ["SKU", "Barcode", "Arabic Product Name", "Unit Cost"]])
                    merged_df[branch_cols] = merged_df[branch_cols].fillna(0)

                    # Add Total Quantity
                    merged_df["Total Quantity"] = merged_df[branch_cols].sum(axis=1)

                    # Reorder: move Unit Cost to last before Grand Total
                    reordered_cols = ["SKU", "Barcode", "Arabic Product Name"] + branch_cols + ["Total Quantity", "Unit Cost"]
                    merged_df = merged_df[reordered_cols]

                    # Add Grand Total
                    merged_df["Grand Total"] = merged_df["Total Quantity"] * merged_df["Unit Cost"]

                    return merged_df


                khateer_pivot = create_aggregated_df(khateer_data)
                khodar_pivot = create_aggregated_df(khodar_data)

                if khateer_pivot is not None:
                    khateer_buffer = io.BytesIO()
                    khateer_pivot.to_excel(khateer_buffer, index=False)
                    output_zip.writestr("مجمع خطير.xlsx", khateer_buffer.getvalue())

                if khodar_pivot is not None:
                    khodar_buffer = io.BytesIO()
                    khodar_pivot.to_excel(khodar_buffer, index=False)
                    output_zip.writestr("مجمع رابيت.xlsx", khodar_buffer.getvalue())
            last_invoice_number = base_invoice_num + file_index  # Add this line here
            st.success("Processing complete.")
            st.info(f"آخر رقم فاتورة تم استخدامه هو: {last_invoice_number}")
            if st.download_button(
                label="Download ZIP with Cleaned and Pivoted Files",
                data=output_zip_io.getvalue(),
                file_name=f"rabbit & Khateer Files_{delivery_date}.zip",
                mime="application/zip"
            ):
                if base_invoice_num == default_base_invoice_num:
                    df_invoice_number.iat[0, 0] = last_invoice_number
                else:
                    df_invoice_number.iat[0, 0] = default_base_invoice_num
                conn.update(worksheet="Saved", data=df_invoice_number)
if __name__ == "__main__":
    rabbitInvoices()
