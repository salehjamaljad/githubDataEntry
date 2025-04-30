import streamlit as st
import os
import pandas as pd
import pdfplumber
import zipfile
import tempfile
from io import BytesIO
from fuzzywuzzy import process
from datetime import datetime
import pytz
import io
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
import os
def pdfToExcel():
    # Define your standard column names
    columns = [
        "No.", 
        "SKU", 
        "Supplier SKU", 
        "Barcode", 
        "Product", 
        "Qty", 
        "Unit\nCost", 
        "Disc.\nAmt.", 
        "Amt.\nExcl.\nVAT", 
        "VAT\n%", 
        "VAT\nAmt.", 
        "Amt.\nIncl.\nVAT"
    ]
    standardized_columns = [col.replace("\n", "_") for col in columns]
    translation_dict = {926242: 'ملوخية جاهزة500جم',
        924881: 'فاصوليا مقطعة فريش 350 جرام',
        924880: 'كابوتشا مقطع 350 جم',
        924879: 'ميكس كرنب سلطة مقطع فريش 350 جرام',
        924878: 'بطاطس شيبسى فريش 350 جرام',
        924877: 'بطاطس صوابع فريش 350 جرام',
        924876: 'فلفل مقور محشي 350 جم',
        924875: 'بطاطس شرائح 350 جم',
        924874: 'جزر مقطع 350 جم',
        924873: 'كوسة حلقات 350جم',
        924871: 'فلفل حلو 250 جرام',
        924868: 'رمان 1كجم',
        924867: 'ابوفروة 250جم',
        924864: 'يوسفي كلمنتينا 1كجم',
        924862: 'يوسفي بلدي 1كجم',
        924861: 'بلح برحي 500 جم',
        924860: 'خوخ مستورد 500 جم',
        924859: 'نكتارين مستورد 500 جم',
        924858: 'فلفل اخضر كوبي 250جم',
        924857: 'ثوم صيني ابيض 200جم',
        924856: 'فول سوداني 500جم',
        913437: 'رمان مفرط 350 جم',
        912855: 'بامية جاهزة 350 جم',
        912854: 'ملوخية جاهزة350جم',
        912852: 'كريز 250 جم',
        912850: 'مانجو فونس 1 ك',
        912849: 'مانجو فص عويس 500 جم',
        912848: 'مانجو عويس 1 ك',
        912847: 'مانجو صديقة 1ك',
        912846: 'مانجو زبدية 1ك',
        912845: 'برقوق احمر محلي 1ك',
        912844: 'عنب بناتى 1ك',
        912843: 'عنب ايرلي سويت ابيض 1ك',
        912842: 'عنب فليم احمر 1ك',
        912841: 'قصب مقشر350جم',
        912840: 'قلقاس مكعبات فريش 350 جرام',
        912839: 'محشى مشكل فريش 350 جرام',
        912838: 'كوسة مقورة فريش 350 جرام',
        912837: 'كمثرى افريقي500 جرام',
        912836: 'فلفل الوان معبأ 500 جرام',
        911211: 'تفاح اصفر ايطالى 1ك معبأ',
        911045: 'برتقال عصير 2ك معبأ',
        911044: 'جوز هند قطعة',
        911043: 'يوسفي موركت 1ك',
        911042: 'جوافة 1ك معبأ',
        911041: 'بطاطا 1ك',
        911040: 'برتقال بسرة 1ك',
        911039: 'بطاطس معبأ 1ك',
        911038: 'بصل احمر معبأ 1ك',
        911037: 'بصل ابيض معبأ 1ك',
        911036: 'باذنجان كوبى معبأ 1ك',
        910161: 'عنب كريمسون لبنانى 500 جرام معبأ',
        910159: 'قرع مكعبات صافى 350 جرام',
        910158: 'عبوة ثوم مفصص 100 جرام',
        910157: 'خضار مشكل فريش 350 جرام',
        910156: 'سوتيه فريش 350 جرام',
        910155: 'بسلة مفصصة بالجزر فريش 350 جرام',
        910154: 'بسلة مفصصة فريش 350 جرام',
        910153: 'عنب اسود مستورد 500 جرام معبأ',
        910152: 'موز مستورد 1ك',
        910151: 'كيوي فاخر 250 جرام معبأ',
        910150: 'تفاح اخضر امريكى 1ك معبأ',
        910149: 'تفاح سكرى جالا 1ك معبأ',
        910148: 'تفاح احمر مستورد 1ك معبأ',
        910147: 'برقوق احمر مستورد 1ك',
        910146: 'اناناس سكري فاخر معبأ',
        910144: 'افوكادو 500 جرام',
        910142: 'عنب ابيض مستورد 500 جرام',
        910141: 'موز بلدي فاخر 1ك معبأ',
        910140: 'كنتالوب 2ك معبأ',
        910139: 'كزبرة معبأ',
        910138: 'كرفس فرنساوي 250 جرام',
        910137: 'شبت معبأ',
        910136: 'زعتر فريش معبأ',
        910135: 'ريحان اخضر معبأ',
        910134: 'روزمارى فريش معبأ',
        910133: 'جرجير معبأ',
        910132: 'بقدونس معبأ',
        910131: 'مشروم 200 جرام معبأ',
        910130: 'كرنب احمر سلطة معبأ',
        910129: 'كرنب ابيض سلطة معبأ',
        910128: 'كابوتشى معبأ',
        910127: 'زنجبيل 100 جرام معبأ',
        910126: 'ذرة سكري 2 قطعه',
        910125: 'خس بلدي فاخر معبأ',
        910124: 'بصل اخضر معبأ',
        910123: 'ليمون بلدى فاخر معبأ 250 جرام',
        910122: 'ليمون اضاليا 250 جرام',
        910121: 'كوسة معبأ 500 جرام',
        910120: 'كرات 250 جرام',
        910119: 'قرنبيط 500 جرام',
        910117: 'فلفل اخضر حار معبأ 250 جرام',
        910116: 'فجل احمر 500 جرام',
        910115: 'طماطم فاخر معبأ 1ك',
        910114: 'طماطم شيرى معبأ 250 جرام',
        910113: 'خيار فاخر معبأ 1ك',
        910112: 'جزر معبأ 500 جرام',
        910111: 'بنجر احمر معبأ 500 جرام',
        910110: 'بروكلي 500 جرام',
        910108: 'باذنجان عروس اسود معبأ 500 جرام',
        912853: 'عنب اسود 1ك',
        910109: 'باذنجان عروس ابيض معبأ 500 جرام',
        910118: 'فلفل حار احمر 250 جرام',
        910143: 'فراوله 250 جرام',
        910145: 'كمثري لبناني 500 جرام',
        910160: 'حرنكش مقشر 250 جرام',
        911046: 'برقوق اصفر مستورد 1ك',
        911047: 'بلح عراقي 1ك',
        911212: 'بطيخ',
        911213: 'بطيخ احمر بدون بذر',
        911214: 'بطيخ اصفر بدون بذر',
        911215: 'خوخ سكرى',
        924865: 'بسلة 500 جم',
        924866: 'فاصوليا خضراء 500جم',
        924869: 'جريب فروت ابيض 1كجم',
        924870: 'جريب فروت احمر 1كجم',
        924872: 'خوخ محلي 1كجم',
        924863: 'يوسفي كريستينا 1كجم',
        912835: 'شمام شهد 1ك معبأ'}
    categories_dict = {
            "فاكهة": [
                "افوكادو 500 جرام", "اناناس سكري فاخر معبأ", "برتقال عصير 2ك معبأ", "بروكلي 500 جرام", 
                "تفاح احمر مستورد 1ك معبأ", "تفاح اخضر امريكى 1ك معبأ", "تفاح اصفر ايطالى 1ك معبأ", 
                "تفاح سكرى جالا 1ك معبأ", "جوز هند قطعة", "زنجبيل 100 جرام معبأ", "طماطم شيرى معبأ 250 جرام", 
                "عنب اسود مستورد 500 جرام معبأ", "قصب مقشر350جم", "كنتالوب 2ك معبأ", "كيوي فاخر 250 جرام معبأ", 
                "مشروم 200 جرام معبأ", "موز بلدي فاخر 1ك معبأ", "موز مستورد 1ك", "يوسفي موركت 1ك", 
                "ابوفروة 250جم", "فول سوداني 500جم", "يوسفي بلدي 1كجم", "ذرة سكري 2 قطعه", "تفاح أصفر إيطالي", "بطيخ احمر بدون بذر", "بطيخ اصفر بدون بذر",
                "خوخ سكرى", "عنب اسود 1ك", "عنب ايرلي سويت ابيض 1ك", "بطيخ", "خوخ محلي 1كجم"
            ],
            "خضار": [
                "باذنجان عروس اسود معبأ 500 جرام", "باذنجان كوبى معبأ 1ك", "بصل ابيض معبأ 1ك", "بصل احمر معبأ 1ك", 
                "بطاطس معبأ 1ك", "بنجر احمر معبأ 500 جرام", "جزر معبأ 500 جرام", "خيار فاخر معبأ 1ك", 
                "طماطم فاخر معبأ 1ك", "فلفل اخضر حار معبأ 250 جرام", "فلفل الوان معبأ 500 جرام", "كوسة معبأ 500 جرام", 
                "ليمون بلدى فاخر معبأ 250 جرام", "بطاطا 1ك", "ثوم صيني ابيض 200جم", "فلفل اخضر كوبي 250جم", 
                "فلفل حلو 250 جرام", "فجل احمر 500 جرام", "ليمون اضاليا 250 جرام", "جزر  معبأ 500 جرام"
            ],
            "مجهز": [
                "سوتيه فريش 350 جرام", "بسلة مفصصة بالجزر فريش 350 جرام", "بسلة مفصصة فريش 350 جرام", 
                "خضار مشكل فريش 350 جرام", "عبوة ثوم مفصص 100 جرام", "قرع مكعبات صافى 350 جرام", 
                "قلقاس مكعبات فريش 350 جرام", "كوسة مقورة فريش 350 جرام", "محشى مشكل فريش 350 جرام", 
                "بطاطس شرائح 350 جم", "بطاطس شيبسى فريش 350 جرام", "بطاطس صوابع فريش 350 جرام", 
                "جزر مقطع 350 جم", "فاصوليا مقطعة فريش 350 جرام", "فلفل مقور محشي 350 جم", 
                "كابوتشا مقطع 350 جم", "كوسة حلقات 350جم", "ميكس كرنب سلطة مقطع فريش 350 جرام", 
                "رمان مفرط 350 جم", "قرنبيط 500 جرام", "خضار  مشكل فريش 350 جرام", "ملوخية جاهزة350جم", "ملوخية جاهزة500جم"
            ],
            "ورقيات وأعشاب": [
                "بصل اخضر معبأ", "بقدونس معبأ", "خس بلدي فاخر معبأ", "روزمارى فريش معبأ", "ريحان اخضر معبأ", 
                "زعتر فريش معبأ", "شبت معبأ", "كابوتشى معبأ", "كرنب ابيض سلطة معبأ", "كرنب احمر سلطة معبأ", 
                "كزبرة معبأ", "كرفس فرنساوي 250 جرام"
            ]}
    # Dictionary of branches
    branches_dict = {
        "EG_Alex East_DS_26": "سيدي بشر",
        "EG_Alex West_DS_27": "الابراهيميه",
        "EG_Alex_Wingat_DS_41": "وينجت",
        "EG_Cairo_DS_1": "المعادي لاسلكي",
        "EG_Cairo_DS_2": "الدقي",
        "EG_Cairo_DS_4": "زهراء المعادي",
        "EG_Cairo_DS_5": "ميدان لبنان",
        "EG_Cairo_DS_7": "العجوزة",
        "EG_Cairo_DS_9": "كورنيش المعادي",
        "EG_Zahraa Maadi 2_DS_49": "زهراء المعادي - 2",
        "EG_Assuit_DS_35": "اسيوط",
        "EG_Cairo_DS_10": "هيليوبليس",
        "EG_Cairo_DS_11": "هرم ترسا",
        "EG_Cairo_DS_12": "اكتوبر",
        "EG_Cairo_DS_17": "سيتي ستارز",
        "EG_Cairo_DS_19": "فرست مول",
        "EG_Cairo_DS_20": "فونت مول",
        "EG_Cairo_DS_21": "حدائق القبه",
        "EG_Cairo_DS_22": "الشيخ زايد",
        "EG_Cairo_DS_3": "الرحاب",
        "EG_Cairo_DS_31": "عين شمس",
        "EG_Cairo_DS_37_Tagamoa-Awal": "التجمع الاول",
        "EG_Cairo_DS_8": "مدينة نصر",
        "EG_faisal_DS_42": "فيصل",
        "EG_Hadayek October_DS_44": "حدائق اكتوبر",
        "EG_Ismailia_DS_34": "الاسماعيليه",
        "EG_Madinaty Craft_DS_39": "مدينتي كرافت",
        "EG_Madinaty_DS_23": "مدينتي",
        "EG_Mansoura gomhoreya_DS_48": "المنصورة جمهورية",
        "EG_Mansoura_DS_25": "المنصورة",
        "EG_Nasrcity 10th_DS_40": "الحي العاشر",
        "EG_Obour_DS_30": "العبور",
        "EG_October industrial_DS_47": "برايت مول",
        "EG_Palmhills_DS_36": "بالم هيلز",
        "EG_Portsaid_DS_32": "بورسعيد",
        "EG_Rehab_chillout_DS_50": "الرحاب تشيل اوت",
        "EG_Shrouk_DS_29": "الشروق",
        "EG_Tagamoa 5_Mahkama_DS_43": "التجمع محكمه",
        "EG_Tagamoa Golden Sq_DS_45": "التجمع جولدن سكوير",
        "EG_Tanta_DS_24": "طنطا",
        "EG_Zakazik_DS_33": "الزقازيق",
        "EG_Heliopolis_Sheraton_DS_52": "هيليوبليس شيراتون"
    }

    # Special EG_ codes that need to capture the next word too
    special_codes = {
        "EG_Alex East_DS_", "EG_Alex", "EG_Zahraa Maadi", "EG_Nasrcity", "EG_Mansoura", 
        "EG_Tagamoa Golden", "EG_Tagamoa", "EG_Madinaty", "EG_Hadayek", "EG_October"
    }

    def extract_eg_codes(pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + " "  # Space to avoid word sticking issues

            # Tokenize the text by splitting on whitespace
            words = text.split()
            i = 0
            results = []

            while i < len(words):
                word = words[i]
                
                # Check if word starts with EG_ and is in special_codes
                if word.startswith("EG_"):
                    # If the word is in special_codes, capture the next word
                    if any(word == code or word.startswith(code) for code in special_codes):
                        next_word = words[i + 1] if i + 1 < len(words) else ""
                        combined = f"{word} {next_word}"
                        
                        # Fuzzy match for the closest branch in branches_dict
                        closest_match, score = process.extractOne(combined, branches_dict.keys())
                        if score >= 80:  # You can adjust the score threshold if needed
                            results.append({
                                "filename": os.path.basename(pdf_path), 
                                "extracted": combined, 
                                "matched_key": closest_match, 
                                "arabic_name": branches_dict[closest_match]
                            })
                        else:
                            results.append({"filename": os.path.basename(pdf_path), "extracted": combined})
                        
                        i += 1  # Skip the next word because it's already included
                    else:
                        # Fuzzy match for a single EG_ code if not in special_codes
                        closest_match, score = process.extractOne(word, branches_dict.keys())
                        if score >= 80:  # You can adjust the score threshold if needed
                            results.append({
                                "filename": os.path.basename(pdf_path), 
                                "extracted": word, 
                                "matched_key": closest_match, 
                                "arabic_name": branches_dict[closest_match]
                            })
                        else:
                            results.append({"filename": os.path.basename(pdf_path), "extracted": word})
                i += 1
            
            return results

    def process_pdf(file_path):
        all_tables = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table)
                    all_tables.append(df)

        for i, table in enumerate(all_tables):
            non_null_counts = table.notnull().sum()
            threshold = non_null_counts.max() * 0.5
            columns_to_drop = non_null_counts[non_null_counts <= threshold].index
            if len(columns_to_drop) > 0:
                all_tables[i] = table.drop(columns=columns_to_drop)

        for i, df in enumerate(all_tables[1:]):  # Skip first table if needed
            df.columns = standardized_columns

        if len(all_tables) > 1:
            final_df = pd.concat(all_tables[1:], ignore_index=True)
        else:
            final_df = all_tables[0]

        df = final_df
        df = df.loc[~(df.applymap(lambda x: x == '').all(axis=1))]
        df = df.reset_index(drop=True)
        df = df[df['Qty'] != '']
        df = df[df['SKU'] != 'SKU']
        # Drop the specified columns
        df.drop(columns=[
            'Disc._Amt.', 
            'Amt._Excl._VAT', 
            'VAT_%', 
            'VAT_Amt.', 
            'Supplier SKU',
            'No.',
            'Product'
        ], inplace=True)

        # Rename the specified columns
        df.rename(columns={
            'Unit_Cost': 'PP',
            'Amt._Incl._VAT': 'Total'
        }, inplace=True)

        # Convert data types
        df['PP'] = df['PP'].astype(float)
        df['Total'] = df['Total'].astype(float)
        df['Qty'] = df['Qty'].astype(int)
        df['Barcode'] = df['Barcode'].astype(int)
        df['SKU'] = df['SKU'].astype(int)
        df["Item Name Ar"] = df["SKU"].map(translation_dict)
        df = df[['SKU', 'Barcode', 'Item Name Ar', 'PP', 'Qty', 'Total']]
        df = df.reset_index(drop=True)
        return df

    st.title("PDF to Excel Converter (Bulk)")
    selected_date = st.date_input('enter the delivery date')
    base_invoice_num = st.number_input("Enter base invoice number", min_value=0, step=1)
    uploaded_zip = st.file_uploader("Upload a ZIP file containing PDFs", type=["zip"])

    if uploaded_zip is not None:
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save the uploaded zip file
            zip_path = os.path.join(temp_dir, "uploaded.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.read())

            # Extract the zip
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # Create a temp folder for Excel outputs
            output_dir = os.path.join(temp_dir, "excels")
            os.makedirs(output_dir, exist_ok=True)

            # Process each PDF
            pos_with_filenames = {}
            for filename in os.listdir(temp_dir):
                if filename.endswith(".pdf"):
                    file_path = os.path.join(temp_dir, filename)
                    df = process_pdf(file_path)
                    po = filename.split("-")[0].strip()
                    pos_with_filenames[filename] = po
                    # Extract branch name for renaming
                    extracted_data = extract_eg_codes(file_path)
                    branch_name = None
                    if extracted_data:
                        branch_name = extracted_data[0].get("arabic_name", None)

                    # Use the branch name for renaming to Arabic only
                    if branch_name:
                        output_filename = f"{branch_name}_{po}.xlsx"
                    else:
                        output_filename = f"{os.path.splitext(filename)[0]}.xlsx"

                    output_path = os.path.join(output_dir, output_filename)
                    df["po"] = po
                    df.to_excel(output_path, index=False, columns=[col for col in df.columns if col != "po"])

            
            all_dfs = []

            for excel_file in os.listdir(output_dir):
                if excel_file.endswith(".xlsx"):
                    excel_path = os.path.join(output_dir, excel_file)
                    df = pd.read_excel(excel_path)

                    # Get branch name from file name (remove .xlsx)
                    base = os.path.splitext(excel_file)[0]
                    parts = base.rsplit("_", 1)
                    if len(parts) == 2:
                        branch_name, po = parts
                    df["branch"] = branch_name
                    df["po"] = po
                    

                    all_dfs.append(df)

            if all_dfs:
                combined_df = pd.concat(all_dfs, ignore_index=True)
                # Ensure SKU is integer
                combined_df["SKU"] = pd.to_numeric(combined_df["SKU"], errors="coerce").astype("Int64")

                # Replace Product column using SKU mapped through translation_dict
                combined_df["Product"] = combined_df["SKU"].map(translation_dict)
                # Create a reverse mapping: product -> category
                reverse_categories = {
                    item: category for category, items in categories_dict.items() for item in items
                }

                # Map the Product column to its category
                combined_df["category"] = combined_df["Product"].map(reverse_categories)

                st.dataframe(combined_df)



                # Pivot the dataframe
                pivot_df = combined_df.pivot_table(
                    index=["SKU", "Product", "category"],
                    columns="branch",
                    values="Qty",
                    aggfunc="sum",
                    fill_value=0
                ).reset_index()

                # Rename 'Product' to 'Product name' for consistency
                pivot_df = pivot_df.rename(columns={"Product": "Product name"})
                pivot_df[sorted(pivot_df.columns)]
                # Define column groups
                alexandria_columns = ['Product name', 'SKU', 'category', 'سيدي بشر', 'الابراهيميه', 'وينجت']
                ready_veg_columns = ['Product name', 'SKU', 'category', 'المعادي لاسلكي', 'الدقي', 'زهراء المعادي',
                                    'ميدان لبنان', 'العجوزة', 'كورنيش المعادي', 'زهراء المعادي - 2']
                
                
                # Always-include base columns
                base_columns = ['Product name', 'SKU', 'category']

                # Get all unique branch columns used in Alex and Ready Veg
                used_branch_columns = set(alexandria_columns + ready_veg_columns) - set(base_columns)

                # All other branch columns are considered Cairo
                cairo_branch_columns = [col for col in pivot_df.columns if col not in used_branch_columns and col not in base_columns]

                # Now construct the final Cairo columns list
                cairo_columns = base_columns + cairo_branch_columns


                # Create the split DataFrames
                alexandria_df = pivot_df[[col for col in alexandria_columns if col in pivot_df.columns]]
                ready_veg_df = pivot_df[[col for col in ready_veg_columns if col in pivot_df.columns]]
                cairo_df = pivot_df[[col for col in cairo_columns if col in pivot_df.columns]]
                alexandria_df[sorted(alexandria_df.columns)]
                ready_veg_df[sorted(ready_veg_df.columns)]
                cairo_df[sorted(cairo_df.columns)]


                # Define category sort order
                category_order = {
                    "فاكهة": 1,
                    "خضار": 2,
                    "مجهز": 3,
                    "ورقيات وأعشاب": 4
                }

                def add_total_and_sort(df):
                    # Identify branch columns by excluding fixed columns
                    fixed_cols = ['Product name', 'SKU', 'category']
                    branch_cols = [col for col in df.columns if col not in fixed_cols]

                    # Add total column
                    df["total"] = df[branch_cols].sum(axis=1)

                    # Map sort key for category
                    df["category_order"] = df["category"].map(category_order)

                    # Sort by category order and product name
                    df = df.sort_values(by=["category_order", "Product name"], ascending=[True, True])

                    # Drop helper column
                    df = df.drop(columns=["category_order"])

                    return df

                # Apply to each DataFrame
                alexandria_df = add_total_and_sort(alexandria_df)
                ready_veg_df = add_total_and_sort(ready_veg_df)
                cairo_df = add_total_and_sort(cairo_df)
                # Filter out rows where total is 0
                alexandria_df = alexandria_df[alexandria_df["total"] != 0]
                ready_veg_df = ready_veg_df[ready_veg_df["total"] != 0]
                cairo_df = cairo_df[cairo_df["total"] != 0]

                

                egypt_tz = pytz.timezone('Africa/Cairo')
                today_str = datetime.now(egypt_tz).strftime("%Y-%m-%d")  # Format: YYYY-MM-DD

                
                
                
                
                def set_paragraph_rtl(paragraph):
                    """Set paragraph direction to RTL."""
                    p = paragraph._p
                    pPr = p.get_or_add_pPr()
                    bidi = OxmlElement('w:bidi')
                    bidi.set(qn('w:val'), '1')
                    pPr.append(bidi)

                def create_docx_from_dfs(all_dfs, selected_date, base_invoice_num, branches_dict):
                    docx_files = {}

                    # Map from branch name to its df
                    branch_dfs = {}
                    for df in all_dfs:
                        if 'branch' not in df.columns:
                            continue
                        branch_key = df['branch'].iloc[0]
                        branch_name = branches_dict.get(branch_key, branch_key)
                        branch_dfs[branch_name] = df

                    # Priority branches first, then alphabetical
                    priority = ["الابراهيميه", "سيدي بشر", "وينجت"]
                    other_branches = sorted([b for b in branch_dfs if b not in priority])
                    sorted_branch_names = [b for b in priority if b in branch_dfs] + other_branches

                    # Create documents with padded invoice numbers
                    invoice_num = base_invoice_num
                    for branch_name in sorted_branch_names:
                        df = branch_dfs[branch_name]
                        customer_name = f"دليفيري هيرو ديمارت ايجيبت فرع {branch_name}"
                        po = df["po"].iloc[0] if "po" in df.columns else ""

                        df_to_save = df.copy()
                        if 'Qty' in df_to_save.columns:
                            df_to_save['Qty'] = ''
                        if 'Total' in df_to_save.columns:
                            df_to_save['Total'] = ''
                        df_to_save.drop(columns=['branch', 'po'], inplace=True, errors='ignore')

                        padded_invoice = str(invoice_num).zfill(8)
                        invoice_num += 1

                        doc = Document()

                        # Create a 2-row, 2-column table
                        image_table = doc.add_table(rows=2, cols=2)
                        image_table.autofit = False  # Disable autofit to control widths

                        # Adjust column widths: 3/4 for the left side (Pictures 3 & 4), 1/4 for the right side (Pictures 1 & 2)
                        total_width = Inches(6)  # Total width of the table, you can adjust as needed
                        left_col_width = total_width * 0.75  # 3/4 of the width
                        right_col_width = total_width * 0.25  # 1/4 of the width

                        # Apply column widths
                        for row in image_table.rows:
                            row.cells[0].width = left_col_width
                            row.cells[1].width = right_col_width

                        # Set picture paths
                        pictures = {
                            (0, 1): "Picture1.png",  # top-right (narrow)
                            (1, 1): "Picture2.png",  # bottom-right (narrow)
                            (0, 0): "Picture3.png",  # top-left (wide)
                            (1, 0): "Picture4.png"   # bottom-left (wide)
                        }

                        # Add pictures and minimize vertical spacing
                        for (row_idx, col_idx), img_path in pictures.items():
                            try:
                                cell = image_table.cell(row_idx, col_idx)
                                paragraph = cell.paragraphs[0]
                                run = paragraph.add_run()
                                # Control image size to reduce overall table height
                                image_width = right_col_width if col_idx == 1 else left_col_width
                                run.add_picture(img_path, width=image_width, height=Inches(1.25))

                                # Remove space above/below image
                                paragraph.paragraph_format.space_before = Pt(0)
                                paragraph.paragraph_format.space_after = Pt(0)
                            except Exception as e:
                                print(f"Error adding image {img_path}: {e}")

                        p0 = doc.add_paragraph(f"فاتورة مبيعات رقم/ {padded_invoice}")
                        set_paragraph_rtl(p0)

                        p1 = doc.add_paragraph(f"تحريرا في/ {selected_date}")
                        set_paragraph_rtl(p1)

                        p2 = doc.add_paragraph(f"اسم العميل/ {customer_name}")
                        set_paragraph_rtl(p2)

                        p3 = doc.add_paragraph(f"{po}/ امر شراء رقم ")
                        set_paragraph_rtl(p3)

                        table = doc.add_table(rows=1, cols=len(df_to_save.columns))
                        table.style = 'Table Grid'

                        hdr_cells = table.rows[0].cells
                        for j, column in enumerate(df_to_save.columns):
                            hdr_cells[j].text = str(column)

                        for _, row in df_to_save.iterrows():
                            row_cells = table.add_row().cells
                            for j, value in enumerate(row):
                                row_cells[j].text = str(value)

                        docx_buffer = BytesIO()
                        doc.save(docx_buffer)
                        docx_buffer.seek(0)

                        filename = f"{branch_name}.docx"
                        docx_files[filename] = docx_buffer.getvalue()

                    return docx_files


                docx_files = create_docx_from_dfs(all_dfs, selected_date, base_invoice_num, branches_dict)
                
                
                
                
                
                
                
                
                # Add Cairo DF
                cairo_buffer = BytesIO()
                with pd.ExcelWriter(cairo_buffer, engine='xlsxwriter') as writer:
                    cairo_df.to_excel(writer, index=False)
                
                # Add Ready Veg DF
                ready_buffer = BytesIO()
                with pd.ExcelWriter(ready_buffer, engine='xlsxwriter') as writer:
                    ready_veg_df.to_excel(writer, index=False)
                
                # Add Alexandria DF
                alex_buffer = BytesIO()
                with pd.ExcelWriter(alex_buffer, engine='xlsxwriter') as writer:
                    alexandria_df.to_excel(writer, index=False)
                
                # Create ZIP
                output_zip_buffer = BytesIO()
                with zipfile.ZipFile(output_zip_buffer, "w") as zipf:
                    # Add Excel files from output_dir
                    for excel_file in os.listdir(output_dir):
                        excel_path = os.path.join(output_dir, excel_file)
                        zipf.write(excel_path, arcname=excel_file)

                    # Add in-memory Excel dataframes
                    zipf.writestr(f"مجمع_طلبات_اسكندرية_{today_str}.xlsx", alex_buffer.getvalue())
                    zipf.writestr(f"مجمع_طلبات_الخضار_الجاهز_{today_str}.xlsx", ready_buffer.getvalue())
                    zipf.writestr(f"مجمع_طلبات_القاهرة_{today_str}.xlsx", cairo_buffer.getvalue())

                    # Add generated DOCX files
                    for filename, file_data in docx_files.items():
                        zipf.writestr(filename, file_data)

                output_zip_buffer.seek(0)

                st.success("Processing complete!")
                
                st.download_button(
                    label="Download All Files as ZIP",
                    data=output_zip_buffer.getvalue(),
                    file_name="documents.zip",
                    mime="application/zip"
                )
if __name__ == "__main__":
    pdfToExcel()
