import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

def goodsmartInvoices():
    barcode_to_product = {
            "2283957660118": "ملوخية جاهزة500جم",
            "2283957710019": "فاصوليا مقطعة فريش 350 جرام",
            "2283957660106": "كابوتشا مقطع 350 جم",
            "2283957660097": "ميكس كرنب سلطة مقطع فريش 350 جرام",
            "2283957660093": "بطاطس شيبسى فريش 350 جرام",
            "2283957660091": "بطاطس صوابع فريش 350 جرام",
            "2283957660101": "فلفل مقور محشي 350 جم",
            "2283957660105": "بطاطس شرائح 350 جم",
            "2283957660104": "جزر مقطع 350 جم",
            "2283957660103": "كوسة حلقات 350جم",
            "2283957980016": "فلفل حلو 250 جرام",
            "2283957660070": "رمان 1كجم",
            "2283958160045": "ابوفروة 250جم",
            "2283957660068": "يوسفي كلمنتينا 1كجم",
            "2283958160042": "يوسفي بلدي 1كجم",
            "2283958160092": "بلح برحي 500 جم",
            "2283958160132": "خوخ مستورد 500 جم",
            "2283958160095": "نكتارين مستورد 500 جم",
            "2283957880019": "فلفل اخضر كوبي 250جم",
            "2283958160018": "ثوم صيني ابيض 200جم",
            "2283958160038": "فول سوداني 500جم",
            "2283957680015": "رمان مفرط 350 جم",
            "2283957750015": "بامية جاهزة 350 جم",
            "2283958130014": "ملوخية جاهزة350جم",
            "2283958160072": "كريز 250 جم",
            "2283958160084": "مانجو فونس 1 ك",
            "2283958160083": "مانجو فص عويس 500 جم",
            "2283958160073": "مانجو عويس 1 ك",
            "2283958160066": "مانجو صديقة 1ك",
            "2283958160063": "مانجو زبدية 1ك",
            "2283958160054": "برقوق احمر محلي 1ك",
            "2283958160082": "عنب بناتى 1ك",
            "2283958160051": "عنب ايرلي سويت ابيض 1ك",
            "2283958160050": "عنب فليم احمر 1ك",
            "2283958160070": "قصب مقشر350جم",
            "2283958090011": "قلقاس مكعبات فريش 350 جرام",
            "2283958070013": "محشى مشكل فريش 350 جرام",
            "2283958050015": "كوسة مقورة فريش 350 جرام",
            "2283957740022": "كمثرى افريقي500 جرام",
            "2283957910013": "فلفل الوان معبأ 500 جرام",
            "2283957690014": "تفاح اصفر ايطالى 1ك معبأ",
            "2283957740020": "برتقال عصير 2ك معبأ",
            "2283958160057": "جوز هند قطعة",
            "2283958160043": "يوسفي موركت 1ك",
            "2283958160031": "جوافة 1ك معبأ",
            "2283957600013": "بطاطا 1ك",
            "2283957550011": "برتقال بسرة 1ك",
            "2283957580018": "بطاطس معبأ 1ك",
            "2283957540012": "بصل احمر معبأ 1ك",
            "2283958040016": "بصل ابيض معبأ 1ك",
            "2283957990015": "باذنجان كوبى معبأ 1ك",
            "2283958160055": "عنب كريمسون لبنانى 500 جرام معبأ",
            "2283958160060": "قرع مكعبات صافى 350 جرام",
            "2283957770013": "عبوة ثوم مفصص 100 جرام",
            "2283957830014": "خضار مشكل فريش 350 جرام",
            "2283957800017": "سوتيه فريش 350 جرام",
            "2283957650018": "بسلة مفصصة بالجزر فريش 350 جرام",
            "2283957590017": "بسلة مفصصة فريش 350 جرام",
            "2283957740016": "عنب اسود مستورد 500 جرام معبأ",
            "2283957920012": "موز مستورد 1ك",
            "2283957870010": "كيوي فاخر 250 جرام معبأ",
            "2283957660017": "تفاح اخضر امريكى 1ك معبأ",
            "2283957720018": "تفاح سكرى جالا 1ك معبأ",
            "2283957640019": "تفاح احمر مستورد 1ك معبأ",
            "2283957570019": "برقوق احمر مستورد 1ك",
            "2283958160046": "اناناس سكري فاخر معبأ",
            "2283957470012": "افوكادو 500 جرام",
            "2283958160039": "عنب ابيض مستورد 500 جرام",
            "2283958160037": "موز بلدي فاخر 1ك معبأ",
            "2283957840013": "كنتالوب 2ك معبأ",
            "2283958160028": "كزبرة معبأ",
            "2283958160027": "كرفس فرنساوي 250 جرام",
            "2283958160026": "شبت معبأ",
            "2283958160024": "زعتر فريش معبأ",
            "2283958160023": "ريحان اخضر معبأ",
            "2283958160022": "روزمارى فريش معبأ",
            "2283958160021": "جرجير معبأ",
            "2283958160020": "بقدونس معبأ",
            "2283957530013": "مشروم 200 جرام معبأ",
            "2283958140013": "كرنب احمر سلطة معبأ",
            "2283958120015": "كرنب ابيض سلطة معبأ",
            "2283958100017": "كابوتشى معبأ",
            "2283957780012": "زنجبيل 100 جرام معبأ",
            "2283957730017": "ذرة سكري 2 قطعه",
            "2283958160019": "خس بلدي فاخر معبأ",
            "2283958160017": "بصل اخضر معبأ",
            "2283957520014": "ليمون بلدى فاخر معبأ 250 جرام",
            "2283958160011": "ليمون اضاليا 250 جرام",
            "2283958150012": "كوسة معبأ 500 جرام",
            "2283958160016": "كرات 250 جرام",
            "2283958060014": "قرنبيط 500 جرام",
            "2283957950019": "فلفل اخضر حار معبأ 250 جرام",
            "2283958160014": "فجل احمر 500 جرام",
            "2283958160013": "طماطم فاخر معبأ 1ك",
            "2283957850012": "طماطم شيرى معبأ 250 جرام",
            "2283958160012": "خيار فاخر معبأ 1ك",
            "2283957670016": "جزر معبأ 500 جرام",
            "2283957610012": "بنجر احمر معبأ 500 جرام",
            "2283958020018": "بروكلي 500 جرام",
            "2283957940010": "باذنجان عروس اسود معبأ 500 جرام",
            "2283958160074": "عنب اسود 1ك",
            "2283957960018": "باذنجان عروس ابيض معبأ 500 جرام",
            "2283958160015": "فلفل حار احمر 250 جرام",
            "2283958160040": "فراوله 250 جرام",
            "2283958160056": "كمثري لبناني 500 جرام",
            "2283958160062": "حرنكش مقشر 250 جرام",
            "2283957740023": "برقوق اصفر مستورد 1ك",
            "2283958160071": "بلح عراقي 1ك",
            "2283957910070": "بطيخ",
            "2283957910071": "بطيخ احمر بدون بذر",
            "2283957910072": "بطيخ اصفر بدون بذر",
            "2283957910073": "خوخ سكرى",
            "2283957660071": "بسلة 500 جم",
            "2283957660072": "فاصوليا خضراء 500جم",
            "2283958160030": "جريب فروت ابيض 1كجم",
            "2283958160099": "جريب فروت احمر 1كجم",
            "2283957660107": "خوخ محلي 1كجم",
            "2283957660067": "يوسفي كريستينا 1كجم",
            "2283957900014": "شمام شهد 1ك معبأ",
            "": "باذنجان للحشو 350جم",
            "2283958160068": "تفاح مشكل 1كجم"
        }

    categories_dict = {
            "اعشاب": [
    "بصل اخضر معبأ",
    "بقدونس معبأ",
    "جرجير معبأ",
    "خس بلدي فاخر معبأ",
    "روزمارى فريش معبأ",
    "ريحان اخضر معبأ",
    "زعتر فريش معبأ",
    "شبت معبأ",
    "كابوتشى معبأ",
    "كرفس فرنساوي 250 جرام",
    "كرنب ابيض سلطة معبأ",
    "كرنب احمر سلطة معبأ",
    "كزبرة معبأ",
    "ملوخية جاهزة500جم",
    "كرات 250 جرام",
],
"جاهز": [
    "باذنجان للحشو 350جم",
    "بسلة مفصصة بالجزر فريش 350 جرام",
    "بسلة مفصصة فريش 350 جرام",
    "بطاطس شرائح 350 جم",
    "بطاطس شيبسى فريش 350 جرام",
    "بطاطس صوابع فريش 350 جرام",
    "جزر مقطع 350 جم",
    "خضار مشكل فريش 350 جرام",
    "رمان مفرط 350 جم",
    "سوتيه فريش 350 جرام",
    "عبوة ثوم مفصص 100 جرام",
    "فاصوليا مقطعة فريش 350 جرام",
    "فلفل مقور محشي 350 جم",
    "قرع مكعبات صافى 350 جرام",
    "قرنبيط 500 جرام",
    "قصب مقشر350جم",
    "قلقاس مكعبات فريش 350 جرام",
    "كابوتشا مقطع 350 جم",
    "كوسة حلقات 350جم",
    "كوسة مقورة فريش 350 جرام",
    "محشى مشكل فريش 350 جرام",
    "ملوخية جاهزة350جم",
    "ميكس كرنب سلطة مقطع فريش 350 جرام",
    "بامية جاهزة 350 جم",
],
"خضار": [
    "باذنجان عروس اسود معبأ 500 جرام",
    "باذنجان كوبى معبأ 1ك",
    "بصل ابيض معبأ 1ك",
    "بصل احمر معبأ 1ك",
    "بطاطا 1ك",
    "بطاطس معبأ 1ك",
    "بنجر احمر معبأ 500 جرام",
    "ثوم صيني ابيض 200جم",
    "جزر معبأ 500 جرام",
    "خيار فاخر معبأ 1ك",
    "طماطم فاخر معبأ 1ك",
    "فجل احمر 500 جرام",
    "فلفل اخضر حار معبأ 250 جرام",
    "فلفل اخضر كوبي 250جم",
    "فلفل الوان معبأ 500 جرام",
    "فلفل حلو 250 جرام",
    "كوسة معبأ 500 جرام",
    "ليمون اضاليا 250 جرام",
    "ليمون بلدى فاخر معبأ 250 جرام",
    "باذنجان عروس ابيض معبأ 500 جرام",
    "بسلة 500 جم",
    "فاصوليا خضراء 500جم",
    "فلفل حار احمر 250 جرام",
],
"فاكهه": [
    "ابوفروة 250جم",
    "افوكادو 500 جرام",
    "اناناس سكري فاخر معبأ",
    "برتقال عصير 2ك معبأ",
    "بروكلي 500 جرام",
    "بطيخ",
    "بطيخ احمر بدون بذر",
    "بطيخ اصفر بدون بذر",
    "تفاح احمر مستورد 1ك معبأ",
    "تفاح اخضر امريكى 1ك معبأ",
    "تفاح اصفر ايطالى 1ك معبأ",
    "تفاح سكرى جالا 1ك معبأ",
    "تفاح مشكل 1كجم",
    "جوز هند قطعة",
    "خوخ سكرى",
    "خوخ محلي 1كجم",
    "ذرة سكري 2 قطعه",
    "زنجبيل 100 جرام معبأ",
    "طماطم شيرى معبأ 250 جرام",
    "عنب اسود 1ك",
    "عنب اسود مستورد 500 جرام معبأ",
    "عنب ايرلي سويت ابيض 1ك",
    "فول سوداني 500جم",
    "كنتالوب 2ك معبأ",
    "كيوي فاخر 250 جرام معبأ",
    "مشروم 200 جرام معبأ",
    "موز بلدي فاخر 1ك معبأ",
    "موز مستورد 1ك",
    "يوسفي بلدي 1كجم",
    "يوسفي موركت 1ك",
    "برتقال بسرة 1ك",
    "برقوق احمر محلي 1ك",
    "برقوق احمر مستورد 1ك",
    "برقوق اصفر مستورد 1ك",
    "بلح برحي 500 جم",
    "بلح عراقي 1ك",
    "جريب فروت ابيض 1كجم",
    "جريب فروت احمر 1كجم",
    "جوافة 1ك معبأ",
    "حرنكش مقشر 250 جرام",
    "خوخ مستورد 500 جم",
    "رمان 1كجم",
    "شمام شهد 1ك معبأ",
    "عنب ابيض مستورد 500 جرام",
    "عنب بناتى 1ك",
    "عنب فليم احمر 1ك",
    "عنب كريمسون لبنانى 500 جرام معبأ",
    "فراوله 250 جرام",
    "كريز 250 جم",
    "كمثرى افريقي500 جرام",
    "كمثري لبناني 500 جرام",
    "مانجو زبدية 1ك",
    "مانجو صديقة 1ك",
    "مانجو عويس 1 ك",
    "مانجو فص عويس 500 جم",
    "مانجو فونس 1 ك",
    "نكتارين مستورد 500 جم",
    "يوسفي كريستينا 1كجم",
    "يوسفي كلمنتينا 1كجم",
]
}



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
            df.to_excel(writer, index=False, sheet_name="Orders")


            workbook = writer.book
            worksheet1 = writer.sheets["Orders"]

            number_format = workbook.add_format({'num_format': '0'})  # No scientific notation
            worksheet1.set_column("A:A", 20, number_format)      # Barcode
            worksheet1.set_column("B:B", 30)                     # Product Name
            worksheet1.set_column("C:C", 15)                     # pp
            worksheet1.set_column("D:D", 10)                     # Qty
            worksheet1.set_column("E:E", 20)                     # Total Cost
            worksheet1.set_column("F:F", 15)                     # Category


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
    delivery_date = st.date_input("Delivery Date", datetime.date.today())
    invoice_number = st.number_input("Invoice Number", min_value=1, step=1)
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
        st.info(invoice_number)
        # Provide download button
        st.download_button(
            label="Download Invoice Excel",
            data=excel_file,
            file_name=f"GoodsMart_Delivery_{delivery_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
if __name__ == "__main__":
    goodsmartInvoices()
