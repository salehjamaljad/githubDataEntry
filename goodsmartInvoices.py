import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from streamlit_gsheets import GSheetsConnection
from config import barcode_to_product, categories_dict

def goodsmartInvoices():
    def assign_category_with_barcode(df, barcode_to_product, categories_dict):
        # Reverse mapping product → category from categories_dict
        product_to_category = {}
        for cat, products in categories_dict.items():
            for p in products:
                product_to_category[p.strip()] = cat

        def get_category(row):
            barcode = str(row.get("Barcode", "")).strip()
            prod_name = str(row.get("Product Name", "")).strip()
            
            # 1. Try barcode to product
            product_from_barcode = barcode_to_product.get(barcode, "").strip()
            if product_from_barcode and product_from_barcode in product_to_category:
                return product_to_category[product_from_barcode]
            
            # 2. Try product name from the row itself
            if prod_name in product_to_category:
                return product_to_category[prod_name]

            # 3. Not found
            return "غير مصنف"

        df["Category"] = df.apply(get_category, axis=1)

        # Sort by category order then product name
        category_order = ["فاكهه", "خضار", "جاهز", "اعشاب", "غير مصنف"]
        df["Category"] = pd.Categorical(df["Category"], categories=category_order, ordered=True)
        df.sort_values(["Category", "Product Name"], inplace=True)

        return df


    def create_excel_file(df, invoice_num, delivery_date, po_value):
        output = BytesIO()
        branch_name = "Zaied"
        client_name = "Goodsmart - Zaied Branch"

        # Assign categories and sort
        df = assign_category_with_barcode(df, barcode_to_product, categories_dict)


        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Sheet 1: Raw orders with category
            df.to_excel(writer, index=False, sheet_name="Orders", startrow=0)

            # Get worksheet
            worksheet1 = writer.sheets["Orders"]
            workbook = writer.book

            # Set column formats
            number_format = workbook.add_format({'num_format': '0'})  # No scientific notation
            worksheet1.set_column("A:A", 20, number_format)      # Barcode
            worksheet1.set_column("B:B", 30)                     # Product Name
            worksheet1.set_column("C:C", 15)                     # pp
            worksheet1.set_column("D:D", 10)                     # Qty
            worksheet1.set_column("E:E", 20)                     # Total Cost
            worksheet1.set_column("F:F", 15)                     # Category

            # Add Grand Total row
            last_row_index = len(df) + 1  # +1 because of header row
            bold_border = workbook.add_format({'bold': True, 'top': 2})

            # Write label in B
            worksheet1.write(last_row_index, 1, "Grand Total", bold_border)

            # Write SUM formulas in pp, Qty, Total Cost columns
            worksheet1.write_formula(last_row_index, 2, f"=SUM(C2:C{last_row_index})", bold_border)  # pp
            worksheet1.write_formula(last_row_index, 3, f"=SUM(D2:D{last_row_index})", bold_border)  # Qty
            worksheet1.write_formula(last_row_index, 4, f"=SUM(E2:E{last_row_index})", bold_border)  # Total Cost



            # Sheet 2: Invoice
            invoice_ws = workbook.add_worksheet("فاتورة")

            # Formatting
            meta_format = workbook.add_format({'bold': True, 'border': 2})
            bold_center = workbook.add_format({'bold': True, 'align': 'center'})
            bold_merge = workbook.add_format({'bold': True, 'border': 2, 'align': 'center', 'valign': 'vcenter'})
            headers_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})
            border_format = workbook.add_format({'border': 1})
            bold_border_right = workbook.add_format({'bold': True, 'border': 2})

            # A1: Insert image
            try:
                invoice_ws.insert_image("A1", "Picture1.png", {'x_scale': 0.5, 'y_scale': 0.5})
            except:
                pass  # If image is missing, continue silently

            # A5: خضار.كوم (bold, all borders)
            invoice_ws.write("A5", "شركه خضار للتجارة والتسويق", meta_format)   
            
            # Meta section
            invoice_ws.write("C1", "شركه خضار للتجارة والتسويق", meta_format)
            invoice_ws.write("C2", "Khodar for Trading & Marketing", meta_format)
            invoice_ws.write("F1", "فاتورة مبيعات", meta_format)
            invoice_ws.write("F2", "رقم الفاتورة #", meta_format)
            invoice_ws.write("F3", "تاريخ الاستلام", meta_format)
            invoice_ws.write("F4", "امر شراء رقم", meta_format)
            invoice_ws.write("F6", "اسم العميل", meta_format)
            invoice_ws.write("F7", "الفرع", meta_format)

            # Meta values
            invoice_ws.write("E2", invoice_num, meta_format)
            invoice_ws.write("E3", delivery_date.strftime("%Y-%m-%d"), meta_format)
            invoice_ws.write("E4", po_value, workbook.add_format({'border': 2, 'align': 'center', 'bold': True}))
            invoice_ws.write("E6", client_name, meta_format)
            invoice_ws.write("E7", branch_name, meta_format)

            # Table headers
            invoice_ws.write("A11", "Barcode", headers_format)
            invoice_ws.write("B11", "Product Name", headers_format)
            invoice_ws.write("C11", "PP", headers_format)
            invoice_ws.write("D11", "Qty", headers_format)
            invoice_ws.write("E11", "Total", headers_format)

            # Fill data
            for idx, row in df.iterrows():
                row_num = 11 + idx
                barcode = row["Barcode"]
                name = row["Product Name"]
                cost = row["pp"]

                # Barcode formatting (with fallback for strings)
                if pd.isna(barcode) or barcode == '':
                    invoice_ws.write_blank(row_num, 0, "", border_format)
                else:
                    try:
                        invoice_ws.write_number(row_num, 0, int(barcode), workbook.add_format({'num_format': '0', 'border': 1}))
                    except:
                        invoice_ws.write_string(row_num, 0, str(barcode), border_format)

                invoice_ws.write(row_num, 1, name, border_format)
                invoice_ws.write(row_num, 2, cost, border_format)
                invoice_ws.write(row_num, 3, "", border_format)  # Empty Qty
                invoice_ws.write(row_num, 4, "", border_format)  # Empty Total


            last_row = 11 + len(df)

            # Subtotal and Total
            invoice_ws.merge_range(last_row, 0, last_row, 3, "Subtotal", bold_merge)
            invoice_ws.write_blank(last_row, 4, "", bold_border_right)
            invoice_ws.merge_range(last_row + 1, 0, last_row + 1, 3, "Total", bold_merge)
            invoice_ws.write_blank(last_row + 1, 4, "", bold_border_right)

            # Footer
            footer_texts = [
                "شركة خضار للتجارة و التسويق",
                "ش.ذ.م.م",
                "سجل تجارى / 13138  بطاقه ضريبية/721/294/448"
            ]
            for i, text in enumerate(footer_texts):
                row = last_row + 3 + i
                invoice_ws.merge_range(row, 0, row, 3, text, bold_center)

            # Column widths
            invoice_ws.set_column("A:A", 25)
            invoice_ws.set_column("B:B", 25)
            invoice_ws.set_column("C:E", 25)

        output.seek(0)
        return output

    # Streamlit UI
    st.title("GoodsMart Invoice Generator")
    # Calculate the day after tomorrow
    default_date = datetime.today() + timedelta(days=2)

    # Use it as the default value
    delivery_date = st.date_input('Enter the delivery date', value=default_date)
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_invoice_number = conn.read(worksheet="Saved", cell="A1", ttl=5, headers=False)
    default_invoice_number = int(df_invoice_number.iat[0, 0])
    if "invoice_number" not in st.session_state:
        st.session_state.invoice_number = default_invoice_number
    invoice_number = st.number_input("Invoice Number", value=st.session_state.invoice_number, step=1)
    po_value = st.text_input("Purchase Order Number")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])


    if uploaded_file and invoice_number and po_value:
        # Read the uploaded Excel file
        df = pd.read_excel(uploaded_file)

        # Clean column names to avoid KeyError
        df.columns = df.columns.str.strip()

        # Define required columns and renaming map
        required_columns = {
            "Barcode": "Barcode",
            "Arabic Name": "Product Name",
            "Cost": "pp",
            "Qty": "Qty",
            "Total Cost": "Total Cost"
        }

        # Check for missing columns
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            st.error(f"Missing required columns in uploaded file: {missing}")
            st.stop()

        # Extract and rename required columns
        df = df[list(required_columns.keys())].copy()
        df.rename(columns=required_columns, inplace=True)

        # Generate Excel with invoice sheet
        excel_file = create_excel_file(df, int(invoice_number), delivery_date, po_value)
        st.info(f"last invoice number: {invoice_number}")
        
        # Provide download button
        if st.download_button(
            label="Download Invoice Excel",
            data=excel_file,
            file_name=f"GoodsMart_Delivery_{delivery_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            if invoice_number == default_invoice_number:
                df_invoice_number.iat[0, 0] = invoice_number+1
            else:
                df_invoice_number.iat[0, 0] = default_invoice_number
            conn.update(worksheet="Saved", data=df_invoice_number)
if __name__ == "__main__":
    goodsmartInvoices()
