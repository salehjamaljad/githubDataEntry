import streamlit as st
import os
import pandas as pd
import pdfplumber
import zipfile
import tempfile
from io import BytesIO
from fuzzywuzzy import process
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
        df = df.reset_index(drop=True)
        return df

    st.title("PDF to Excel Converter (Bulk)")

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
            for filename in os.listdir(temp_dir):
                if filename.endswith(".pdf"):
                    file_path = os.path.join(temp_dir, filename)
                    df = process_pdf(file_path)

                    # Extract branch name for renaming
                    extracted_data = extract_eg_codes(file_path)
                    branch_name = None
                    if extracted_data:
                        branch_name = extracted_data[0].get("arabic_name", None)

                    # Use the branch name for renaming to Arabic only
                    if branch_name:
                        output_filename = f"{branch_name}.xlsx"
                    else:
                        output_filename = f"{os.path.splitext(filename)[0]}.xlsx"

                    output_path = os.path.join(output_dir, output_filename)
                    df.to_excel(output_path, index=False)

            # Zip all the Excel files
            output_zip_buffer = BytesIO()
            with zipfile.ZipFile(output_zip_buffer, "w") as zipf:
                for excel_file in os.listdir(output_dir):
                    excel_path = os.path.join(output_dir, excel_file)
                    zipf.write(excel_path, arcname=excel_file)

            st.success("Processing complete!")

            # Offer download
            st.download_button(
                label="Download All Excels as ZIP",
                data=output_zip_buffer.getvalue(),
                file_name="excels.zip",
                mime="application/zip"
            )
if __name__ == "__main__":
    pdfToExcel()
