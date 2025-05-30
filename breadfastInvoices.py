import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
import zipfile
import os
import tempfile
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta
from config import barcode_to_product, categories_dict
def breadfastInvoices():
    st.title("Extract Barcodes, Product Names & Quantity from PDF")
    action = st.selectbox(
        "اختر الفرع",
        [
            "الاسكندرية",
            "المنصورة"
        ],
    )
    if action == "الاسكندرية": 
        conn = st.connection("gsheets", type=GSheetsConnection)
        df_invoice_number = conn.read(worksheet="Saved", cell="A1", ttl=5, headers=False)
        
        default_invoice_num_loran = int(df_invoice_number.iat[0, 0])
        if "invoice_num_loran" not in st.session_state:
            st.session_state.invoice_num_loran = default_invoice_num_loran
        invoice_num_loran = st.number_input("رقم الفاتورة - لوران", value=st.session_state.invoice_num_loran, step=1)
        invoice_num_smouha = invoice_num_loran + 1
        # Calculate the day after tomorrow
        default_date = datetime.today() + timedelta(days=1)

        # Use it as the default value
        delivery_date = st.date_input('Enter the delivery date', value=default_date)

        uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")

        def extract_prices(text_block):
            matches = re.findall(r"\s(\d+\.\d{6})\s", text_block)
            return [round(float(p), 2) for p in matches]


        def insert_nulls(barcodes, ids):
            target_indexes = [i for i, id_val in enumerate(ids) if id_val == "6484003"]
            for count, original_index in enumerate(target_indexes):
                adjusted_index = original_index + count
                barcodes.insert(adjusted_index, "")
            return barcodes

        if uploaded_file:
            all_text = ""

            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_text += "\n" + text

            # --- Step 1: Find all Alexandria FP # occurrences and their positions ---
            alex_matches = list(re.finditer(r"Alexandria FP #\d+", all_text))
            
            if len(alex_matches) < 2:
                st.error("Less than two 'Alexandria FP #' entries found in the PDF.")
            else:
                second_fp_text = alex_matches[1].group()  # e.g., "Alexandria FP #2"
                split_pos = alex_matches[1].start()

                # Determine branches based on which FP is second
                if "FP #2" in second_fp_text:
                    branch_before = "سموحة"
                    branch_after = "لوران"
                else:
                    branch_before = "لوران"
                    branch_after = "سموحة"

                text_part1 = all_text[:split_pos]
                text_part2 = all_text[split_pos:]

                po_matches = re.findall(r"#P\d+", all_text)

                # Default fallback if less than 2 found
                po_loran = po_matches[0] if len(po_matches) > 0 else ""
                po_smouha = po_matches[1] if len(po_matches) > 1 else ""

                def extract_data(text_block, branch_name):
                    ids = re.findall(r"\[(\d+)\]", text_block)
                    barcodes = re.findall(r"\s(22\d{11})\s", text_block)
                    quantities = re.findall(r"\s(\d+(?:\.\d+)?)\.0000000\s", text_block)

                    n = len(ids)
                    barcodes = insert_nulls(barcodes, ids)
                    barcodes = barcodes[:n] + [""] * max(0, n - len(barcodes))
                    quantities = quantities[:n] + ["0"] * max(0, n - len(quantities))
                    quantities = [int(float(q)) for q in quantities]
                    prices = extract_prices(text_block)
                    prices = prices[:n] + [""] * max(0, n - len(prices))
                    


                    df = pd.DataFrame({
                        "ID": ids,
                        "Barcode": barcodes,
                        "Quantity": quantities,
                        "pp": prices
                    })

                    def to_int_or_empty(x):
                        try:
                            return int(x)
                        except:
                            return ""

                    df["Barcode"] = df["Barcode"].apply(to_int_or_empty)
                    df["Product Name"] = df["Barcode"].astype(str).map(barcode_to_product).fillna("غير معروف")
                    df["فرع"] = branch_name
                    

                    return df

                df1 = extract_data(text_part1, branch_before)
                df2 = extract_data(text_part2, branch_after)

                # Display
                st.subheader(f"Orders Before Second 'Alexandria FP #' ({branch_before})")
                st.dataframe(df1)

                st.subheader(f"Orders After Second 'Alexandria FP #' ({branch_after})")
                st.dataframe(df2)

            # --- Create Combined Pivot Table (مجمع اسكندرية) ---
                combined_df = pd.concat([df1, df2], ignore_index=True)

                # Pivot the quantities
                pivot_df = combined_df.pivot_table(
                    index=["Barcode", "Product Name", "pp"],
                    columns="فرع",
                    values="Quantity",
                    aggfunc="sum",
                    fill_value=0
                ).reset_index()

                # Ensure both branch columns exist even if one branch had no data
                for col in ["لوران", "سموحة"]:
                    if col not in pivot_df.columns:
                        pivot_df[col] = 0

                # Add total_quantity and total columns
                pivot_df["total_quantity"] = pivot_df["سموحة"] + pivot_df["لوران"]
                pivot_df["total"] = pivot_df["total_quantity"] * pivot_df["pp"]

                # Reorder columns: place total_quantity before pp, total after pp
                pivot_df = pivot_df[["Barcode", "Product Name", "لوران", "سموحة", "total_quantity", "pp", "total"]]
                
                # 1. Create a reverse lookup dictionary: product name -> category
                product_to_category = {}
                for category, products in categories_dict.items():
                    for product in products:
                        product_to_category[product] = category

                # 2. Add 'category' column to pivot_df
                pivot_df["category"] = pivot_df["Product Name"].map(product_to_category).fillna("غير معرف")

                # 3. Define custom sort order for categories
                category_order = ["فاكهه", "خضار", "جاهز", "اعشاب", "غير معرف"]
                pivot_df["category_order"] = pivot_df["category"].apply(lambda x: category_order.index(x))

                # 4. Sort by category then product name
                pivot_df.sort_values(by=["category_order", "Product Name"], inplace=True)

                # 5. Drop the helper sort column
                pivot_df.drop(columns=["category_order"], inplace=True)


                # Create Excel file for مجمع اسكندرية
                def create_pivot_excel(df, filename):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name="مجمع اسكندرية")

                        workbook = writer.book
                        worksheet = writer.sheets["مجمع اسكندرية"]
                        number_format = workbook.add_format({'num_format': '0'})
                        quantity_format = workbook.add_format({'num_format': '0'})
                        price_format = workbook.add_format({'num_format': '0.00'})
                        total_format = workbook.add_format({'num_format': '0.00'})

                        worksheet.set_column("A:A", 20, number_format)
                        worksheet.set_column("B:B", 40)
                        worksheet.set_column("C:D", 10, quantity_format)
                        worksheet.set_column("E:E", 12, quantity_format)      # total_quantity
                        worksheet.set_column("F:F", 10, price_format)         # pp
                        worksheet.set_column("G:G", 15, total_format)         # total

                    output.seek(0)
                    return output

                pivot_excel = create_pivot_excel(pivot_df, "مجمع اسكندرية.xlsx")


                # Generate Excel files for both branches
                def create_excel_file(df, filename, invoice_num, branch_name, po_value):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # --- Sheet 1: Orders as-is ---
                        df.to_excel(writer, index=False, sheet_name="Orders")
                        workbook = writer.book
                        worksheet = writer.sheets["Orders"]
                        qty_col = df.columns.get_loc("Quantity")  # zero-based index
                        pp_col = df.columns.get_loc("pp")

                        last_row = len(df) + 1  # Excel rows are 1-based, header is row 1

                        # Write "Grand Total" label
                        worksheet.write(last_row, 0, "Grand Total", workbook.add_format({'bold': True, 'border': 1}))

                        # Write sum formulas for Qty and PP columns
                        worksheet.write_formula(last_row, qty_col, f"=SUM({chr(65 + qty_col)}2:{chr(65 + qty_col)}{last_row})", workbook.add_format({'bold': True, 'border': 1}))
                        worksheet.write_formula(last_row, pp_col, f"=SUM({chr(65 + pp_col)}2:{chr(65 + pp_col)}{last_row})", workbook.add_format({'bold': True, 'border': 1}))

                        # --- Sheet 2: Invoice Template ---
                        invoice_ws = workbook.add_worksheet("فاتورة")

                        # Formats
                        meta_format = workbook.add_format({'bold': True, 'border': 2})
                        bold_border_right = workbook.add_format({'bold': True, 'border': 2})
                        bold_center = workbook.add_format({'bold': True, 'align': 'center'})
                        bold_merge = workbook.add_format({'bold': True, 'border': 2, 'align': 'center', 'valign': 'vcenter'})
                        headers_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})

                        # A1: Insert image
                        try:
                            invoice_ws.insert_image("A1", "Picture1.png", {'x_scale': 0.5, 'y_scale': 0.5})
                        except:
                            pass  # If image is missing, continue silently

                        # A5: خضار.كوم (bold, all borders)
                        invoice_ws.write("A5", "شركه خضار للتجارة والتسويق", meta_format)

                        # C1–C2: company name
                        invoice_ws.write("C1", "شركه خضار للتجارة والتسويق", meta_format)
                        invoice_ws.write("C2", "Khodar for Trading & Marketing", meta_format)

                        # F1–F7: invoice metadata labels
                        invoice_ws.write("F1", "فاتورة مبيعات", meta_format)
                        invoice_ws.write("F2", "رقم الفاتورة #", meta_format)
                        invoice_ws.write("F3", "تاريخ الاستلام", meta_format)
                        invoice_ws.write("F4", "امر شراء رقم", meta_format)
                        invoice_ws.write("F6", "اسم العميل", meta_format)
                        invoice_ws.write("F7", "الفرع", meta_format)

                        # Fill metadata values
                        invoice_ws.write("E2", invoice_num, meta_format)
                        invoice_ws.write("E3", delivery_date.strftime("%Y-%m-%d"), meta_format)
                        invoice_ws.write("E4", str(po_value), workbook.add_format({'border': 2, 'align': 'center', 'bold':True}))
                        invoice_ws.write("E6", f"بريدفاست - فرع {branch_name}", meta_format)
                        invoice_ws.write("E7", branch_name, meta_format)

                        # A11: Start of invoice table headers
                        invoice_ws.write("A11", "Barcode", headers_format)
                        invoice_ws.write("B11", "Product Name", headers_format)
                        invoice_ws.write("C11", "PP", headers_format)
                        invoice_ws.write("D11", "Qty", headers_format)
                        invoice_ws.write("E11", "Total", headers_format)

                        # Fill invoice table starting from A12 with empty Qty/Total
                        for idx, row in df.iterrows():
                            row_num = 11 + idx
                            barcode_value = row["Barcode"]
                            if pd.isna(barcode_value) or barcode_value == '':
                                # Write empty cell if barcode is missing
                                invoice_ws.write_blank(row_num, 0, "", workbook.add_format({'border': 1}))
                            else:
                                try:
                                    # Try writing as number with number format (no scientific notation)
                                    barcode_int = int(barcode_value)
                                    barcode_format = workbook.add_format({'num_format': '0', 'border': 1})
                                    invoice_ws.write_number(row_num, 0, barcode_int, barcode_format)
                                except Exception:
                                    # If barcode can't be converted to int, write as string
                                    invoice_ws.write_string(row_num, 0, str(barcode_value), workbook.add_format({'border': 1}))

                            invoice_ws.write(row_num, 1, row["Product Name"])
                            invoice_ws.write(row_num, 2, row["pp"])
                            invoice_ws.write(row_num, 3, "")  # Empty Qty
                            invoice_ws.write(row_num, 4, "")  # Empty Total

                        # After writing data rows
                        last_row = 11 + len(df)

                        # Subtotal row
                        invoice_ws.merge_range(last_row, 0, last_row, 3, "Subtotal", bold_merge)
                        invoice_ws.write_blank(last_row, 4, "", bold_border_right)

                        # Total row
                        invoice_ws.merge_range(last_row + 1, 0, last_row + 1, 3, "Total", bold_merge)
                        invoice_ws.write_blank(last_row + 1, 4, "", bold_border_right)

                        # Leave one blank row, then write footer
                        footer_start = last_row + 3
                        footer_texts = [
                            "شركة خضار للتجارة و التسويق",
                            "ش.ذ.م.م",
                            "سجل تجارى / 13138  بطاقه ضريبية/721/294/448"
                        ]

                        for i, text in enumerate(footer_texts):
                            row = footer_start + i
                            invoice_ws.merge_range(row, 0, row, 3, text, bold_center)

                        # Set column widths
                        invoice_ws.set_column("A:A", 25)
                        invoice_ws.set_column("B:B", 25)
                        invoice_ws.set_column("C:E", 25)

                    output.seek(0)
                    return output



                # Create the Excel files
                excel1 = create_excel_file(
                    df1,
                    f"orders_branch_{branch_before}.xlsx",
                    invoice_num_loran if branch_before == "لوران" else invoice_num_smouha,
                    branch_before,
                    po_loran if branch_before == "لوران" else po_smouha
                )

                # Create for branch_after
                excel2 = create_excel_file(
                    df2,
                    f"orders_branch_{branch_after}.xlsx",
                    invoice_num_loran if branch_after == "لوران" else invoice_num_smouha,
                    branch_after,
                    po_loran if branch_after == "لوران" else po_smouha
                )




                # Create a ZIP file in memory
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    zip_file.writestr(f"orders_branch_{branch_before}.xlsx", excel1.getvalue())
                    zip_file.writestr(f"orders_branch_{branch_after}.xlsx", excel2.getvalue())
                    zip_file.writestr("مجمع اسكندرية.xlsx", pivot_excel.getvalue())

                zip_buffer.seek(0)

                # Debug: Show extracted PO values
                st.subheader("Extracted PO Values")
                st.write(f"PO لوران: {po_loran}")
                st.write(f"PO سموحة: {po_smouha}")
                st.info(f"اخر رقم فاتورة هو:{invoice_num_smouha}")

                if st.download_button(
                label="Download ZIP - alexandria Invoice",
                data=zip_buffer.getvalue(),
                file_name=f"breadfast_alex_{delivery_date}.zip",
                mime="application/zip"
                ):
                    # Set the next invoice number based on current state
                    if invoice_num_loran == default_invoice_num_loran:
                        df_invoice_number.iat[0, 0] = invoice_num_smouha + 1
                    else:
                        df_invoice_number.iat[0, 0] = default_invoice_num_loran

                    conn.update(worksheet="Saved", data=df_invoice_number)
    elif action == 'المنصورة':
        # --- UI Input ---
        conn = st.connection("gsheets", type=GSheetsConnection)
        
        df_invoice_number = conn.read(worksheet="Saved", cell="A1", ttl=5, headers=False)
        default_mansoura_invoice_num = int(df_invoice_number.iat[0, 0])
        if "mansoura_invoice_num" not in st.session_state:
            st.session_state.mansoura_invoice_num = default_mansoura_invoice_num
        mansoura_invoice_num = st.number_input("رقم الفاتورة - المنصورة", value=st.session_state.mansoura_invoice_num, step=1)
        # Calculate the day after tomorrow
        default_date = datetime.today() + timedelta(days=1)

        # Use it as the default value
        delivery_date = st.date_input('Enter the delivery date', value=default_date)
        uploaded_file = st.file_uploader("Upload Mansoura PDF", type="pdf")

        # --- Functions ---
        def extract_prices(text_block):
            matches = re.findall(r"\s(\d+\.\d{6})\s", text_block)
            return [round(float(p), 2) for p in matches]

        def insert_nulls(barcodes, ids):
            target_indexes = [i for i, id_val in enumerate(ids) if id_val == "6484003"]
            for count, original_index in enumerate(target_indexes):
                adjusted_index = original_index + count
                barcodes.insert(adjusted_index, "")
            return barcodes
        if uploaded_file:
            all_text = ""
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_text += "\n" + text

            # Extract PO value
            po_match = re.search(r"#P\d+", all_text)
            po_value = po_match.group() if po_match else ""

            def extract_data(text_block):
                ids = re.findall(r"\[(\d+)\]", text_block)
                barcodes = re.findall(r"\s(22\d{11})\s", text_block)
                quantities = re.findall(r"\s(\d+(?:\.\d+)?)\.0000000\s", text_block)

                n = len(ids)
                barcodes = insert_nulls(barcodes, ids)
                barcodes = barcodes[:n] + [""] * max(0, n - len(barcodes))
                quantities = quantities[:n] + ["0"] * max(0, n - len(quantities))
                quantities = [int(float(q)) for q in quantities]
                prices = extract_prices(text_block)
                prices = prices[:n] + [""] * max(0, n - len(prices))

                df = pd.DataFrame({
                    "ID": ids,
                    "Barcode": barcodes,
                    "Quantity": quantities,
                    "pp": prices
                })

                def to_int_or_empty(x):
                    try:
                        return int(x)
                    except:
                        return ""

                df["Barcode"] = df["Barcode"].apply(to_int_or_empty)
                df["Product Name"] = df["Barcode"].astype(str).map(barcode_to_product).fillna("غير معروف")
                df["فرع"] = "المنصورة"

                return df

            df = extract_data(all_text)

            # --- Pivot Table ---
            pivot_df = df.pivot_table(
                index=["Barcode", "Product Name", "pp"],
                columns="فرع",
                values="Quantity",
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            pivot_df["total_quantity"] = pivot_df["المنصورة"]
            pivot_df["total"] = pivot_df["total_quantity"] * pivot_df["pp"]

            # Add category column
            product_to_category = {product: cat for cat, products in categories_dict.items() for product in products}
            pivot_df["category"] = pivot_df["Product Name"].map(product_to_category).fillna("غير معرف")

            # Sort
            category_order = ["فاكهه", "خضار", "جاهز", "اعشاب", "غير معرف"]
            pivot_df["category_order"] = pivot_df["category"].apply(lambda x: category_order.index(x))
            pivot_df.sort_values(by=["category_order", "Product Name"], inplace=True)
            pivot_df.drop(columns=["category_order"], inplace=True)

            # --- Excel Writers ---
            def create_pivot_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name="مجمع المنصورة")
                    ws = writer.sheets["مجمع المنصورة"]
                    fmt = writer.book.add_format({'num_format': '0.00'})
                    ws.set_column("A:A", 20)
                    ws.set_column("B:B", 40)
                    ws.set_column("C:D", 12)
                    ws.set_column("E:E", 14)
                    ws.set_column("F:F", 10, fmt)
                    ws.set_column("G:G", 15, fmt)
                output.seek(0)
                return output

            def create_invoice_excel(df, invoice_num, branch, po_value):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name="Orders")
                    workbook = writer.book
                    worksheet = writer.sheets["Orders"]
                    qty_col = df.columns.get_loc("Quantity")  # zero-based index
                    pp_col = df.columns.get_loc("pp")

                    last_row = len(df) + 1  # Excel rows are 1-based, header is row 1

                    # Write "Grand Total" label
                    worksheet.write(last_row, 0, "Grand Total", workbook.add_format({'bold': True, 'border': 1}))

                    # Write sum formulas for Qty and PP columns
                    worksheet.write_formula(last_row, qty_col, f"=SUM({chr(65 + qty_col)}2:{chr(65 + qty_col)}{last_row})", workbook.add_format({'bold': True, 'border': 1}))
                    worksheet.write_formula(last_row, pp_col, f"=SUM({chr(65 + pp_col)}2:{chr(65 + pp_col)}{last_row})", workbook.add_format({'bold': True, 'border': 1}))
                    

                    invoice_ws = workbook.add_worksheet("فاتورة")

                    try:
                        invoice_ws.insert_image("A1", "Picture1.png", {'x_scale': 0.5, 'y_scale': 0.5})
                    except:
                        pass

                    meta_fmt = workbook.add_format({'bold': True, 'border': 2})
                    center_fmt = workbook.add_format({'bold': True, 'align': 'center'})
                    merge_fmt = workbook.add_format({'bold': True, 'border': 2, 'align': 'center', 'valign': 'vcenter'})
                    header_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})

                    invoice_ws.write("A5", "شركه خضار للتجارة والتسويق", meta_fmt)
                    invoice_ws.write("C1", "شركه خضار للتجارة والتسويق", meta_fmt)
                    invoice_ws.write("C2", "Khodar for Trading & Marketing", meta_fmt)
                    invoice_ws.write("F1", "فاتورة مبيعات", meta_fmt)
                    invoice_ws.write("F2", "رقم الفاتورة #", meta_fmt)
                    invoice_ws.write("F3", "تاريخ الاستلام", meta_fmt)
                    invoice_ws.write("F4", "امر شراء رقم", meta_fmt)
                    invoice_ws.write("F6", "اسم العميل", meta_fmt)
                    invoice_ws.write("F7", "الفرع", meta_fmt)

                    invoice_ws.write("E2", invoice_num, meta_fmt)
                    invoice_ws.write("E3", delivery_date.strftime("%Y-%m-%d"), meta_fmt)
                    invoice_ws.write("E4", str(po_value), workbook.add_format({'border': 2, 'align': 'center', 'bold': True}))
                    invoice_ws.write("E6", f"بريدفاست - فرع {branch}", meta_fmt)
                    invoice_ws.write("E7", branch, meta_fmt)

                    invoice_ws.write("A11", "Barcode", header_fmt)
                    invoice_ws.write("B11", "Product Name", header_fmt)
                    invoice_ws.write("C11", "PP", header_fmt)
                    invoice_ws.write("D11", "Qty", header_fmt)
                    invoice_ws.write("E11", "Total", header_fmt)

                    for idx, row in df.iterrows():
                        r = 11 + idx
                        barcode = row["Barcode"]
                        if barcode == "" or pd.isna(barcode):
                            invoice_ws.write_blank(r, 0, "", workbook.add_format({'border': 1}))
                        else:
                            try:
                                invoice_ws.write_number(r, 0, int(barcode), workbook.add_format({'num_format': '0', 'border': 1}))
                            except:
                                invoice_ws.write(r, 0, str(barcode), workbook.add_format({'border': 1}))
                        invoice_ws.write(r, 1, row["Product Name"])
                        invoice_ws.write(r, 2, row["pp"])
                        invoice_ws.write(r, 3, "")
                        invoice_ws.write(r, 4, "")

                    last = 11 + len(df)
                    invoice_ws.merge_range(last, 0, last, 3, "Subtotal", merge_fmt)
                    invoice_ws.write_blank(last, 4, "", meta_fmt)
                    invoice_ws.merge_range(last+1, 0, last+1, 3, "Total", merge_fmt)
                    invoice_ws.write_blank(last+1, 4, "", meta_fmt)

                    for i, txt in enumerate(["شركة خضار للتجارة و التسويق", "ش.ذ.م.م", "سجل تجارى / 13138  بطاقه ضريبية/721/294/448"]):
                        invoice_ws.merge_range(last+3+i, 0, last+3+i, 3, txt, center_fmt)

                    invoice_ws.set_column("A:E", 25)

                output.seek(0)
                return output

            pivot_excel = create_pivot_excel(pivot_df)
            invoice_excel = create_invoice_excel(df, mansoura_invoice_num, "المنصورة", po_value)

            # Create ZIP
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr("مجمع المنصورة.xlsx", pivot_excel.getvalue())
                zip_file.writestr("فاتورة المنصورة.xlsx", invoice_excel.getvalue())

            zip_buffer.seek(0)
            st.info(f"اخر رقم فاتورة هو: {mansoura_invoice_num}")

            # This part will run only when the user presses the download button
            if st.download_button(
                label="Download ZIP - Mansoura Invoice",
                data=zip_buffer.getvalue(),
                file_name=f"mansoura_invoice_files_{delivery_date}.zip",
                mime="application/zip"
            ):
                # Set the next invoice number based on current state
                if mansoura_invoice_num == default_mansoura_invoice_num:
                    df_invoice_number.iat[0, 0] = mansoura_invoice_num + 1
                else:
                    df_invoice_number.iat[0, 0] = default_mansoura_invoice_num

                conn.update(worksheet="Saved", data=df_invoice_number)

if __name__ == "__main__":
    breadfastInvoices()
