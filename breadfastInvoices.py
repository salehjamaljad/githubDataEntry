import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
import zipfile
import os
import tempfile
def breadfastInvoices():
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

    st.title("Extract Barcodes, Product Names & Quantity from PDF")
    action = st.selectbox(
        "اختر الفرع",
        [
            "الاسكندرية",
            "المنصورة"
        ],
    )
    if action == "الاسكندرية":
        invoice_num_loran = st.number_input("رقم الفاتورة - لوران", min_value=1, step=1)
        invoice_num_smouha = invoice_num_loran + 1
        delivery_date = st.date_input("تاريخ الاستلام")

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
                        worksheet1 = writer.sheets["Orders"]
                        number_format = workbook.add_format({'num_format': '0'})
                        quantity_format = workbook.add_format({'num_format': '0.00'})

                        worksheet1.set_column("A:A", 10)
                        worksheet1.set_column("B:B", 20, number_format)
                        worksheet1.set_column("C:C", 10, quantity_format)
                        worksheet1.set_column("D:D", 50)
                        worksheet1.set_column("E:E", 15)

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
                st.download_button(
                    label="Download ZIP with Both Branch Orders",
                    data=zip_buffer.getvalue(),
                    file_name=f"breadfast_alex_{delivery_date}.zip",
                    mime="application/zip"
                )
    elif action == 'المنصورة':
        # --- UI Input ---
        mansoura_invoice_num = st.number_input("رقم الفاتورة - المنصورة", min_value=1, step=1)
        delivery_date = st.date_input("تاريخ الاستلام")
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
            st.info(f"اخر رقم فاتورة هو:{mansoura_invoice_num}")
            st.download_button(
                label="Download ZIP - Mansoura Invoice",
                data=zip_buffer.getvalue(),
                file_name=f"mansoura_invoice_files_{delivery_date}.zip",
                mime="application/zip"
            )
if __name__ == "__main__":
    breadfastInvoices()
