import streamlit as st
import pandas as pd
import re
from datetime import datetime
import pytz
import io
def pivot_app():
    egypt_tz = pytz.timezone('Africa/Cairo')
    today_str = datetime.now(egypt_tz).strftime("%Y-%m-%d")  # Format: YYYY-MM-DD
    st.title("Pivot 216 CSV File")

    # Upload the CSV file
    file_216 = st.file_uploader("Upload the 216.csv", type="csv")

    if file_216 is not None:
        # Read the uploaded CSV
        def clean_product_name(product):
            if isinstance(product, str):
                # Remove space between number and unit (kg/g)
                product = re.sub(r'(\d+)\s+(kg|g)', r'\1\2', product, flags=re.IGNORECASE)
                # Replace gmm or gm at the end with g
                product = re.sub(r'(gmm|gm)\b', 'g', product, flags=re.IGNORECASE)
            return product

        
        df = pd.read_csv(file_216)
        # Apply the cleaning function to the 'Product' column
        df['Product'] = df['Product'].apply(clean_product_name)
        # Raw Data (optional)
        st.subheader("Raw Data")
        st.dataframe(df)


        branches_dict = {
            "EG_Alex East_DS_ 26": "سيدي بشر",
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
            ]
        }
        df['Product code'] = df['Product code'].astype(int)
        # Create the pivot table
        pivot_df = df.pivot_table(
            index='Product code',
            columns='Store_name',
            values='Effective quantity',
            aggfunc='sum',
            fill_value=0
        )


        # Add a 'Product name' column using the translation_dict
        pivot_df.insert(0, 'Product name', pivot_df.index.map(lambda x: translation_dict.get(x, x)))

        pivot_df.columns = pivot_df.columns.map(lambda x: branches_dict.get(x, x))
        pivot_df = pivot_df[sorted(pivot_df.columns)]
        
        # Define the column groups
        alexandria_columns = ['Product name', 'سيدي بشر', 'الابراهيميه', 'وينجت']
        ready_veg_columns = ['Product name', 'المعادي لاسلكي', 'الدقي', 'زهراء المعادي', 'ميدان لبنان', 'العجوزة', 'كورنيش المعادي', 'زهراء المعادي - 2']
        cairo_columns = [col for col in pivot_df.columns if col not in alexandria_columns and col not in ready_veg_columns or col == 'Product name']


        # Ensure all values are numeric for summing
        pivot_df.loc[:, pivot_df.columns != 'Product name'] = pivot_df.loc[:, pivot_df.columns != 'Product name'].apply(pd.to_numeric, errors='coerce')
        print(pivot_df.columns)

        # Create the Alexandria DataFrame (keeping 'Product' as the index)
        alexandria_df = pivot_df[alexandria_columns].copy()
        alexandria_df = alexandria_df[sorted(alexandria_df.columns)]
        alexandria_df['Total'] = alexandria_df.loc[:, alexandria_df.columns != 'Product name'].sum(axis=1)

        # Get the list of available columns from ready_veg_columns that exist in pivot_df
        available_cols = [col for col in ready_veg_columns if col in pivot_df.columns]

        if available_cols:
            # Create the Ready Veg DataFrame with available columns
            ready_veg_df = pivot_df[available_cols].copy()
            ready_veg_df = ready_veg_df[sorted(ready_veg_df.columns)]
            ready_veg_df['Total'] = ready_veg_df.loc[:, ready_veg_df.columns != 'Product name'].sum(axis=1)
        else:
            # Create an empty DataFrame with the full expected columns
            ready_veg_df = pd.DataFrame(columns=sorted(ready_veg_columns) + ['Total'])


        # Create the Cairo DataFrame (keeping 'Product' as the index)
        cairo_df = pivot_df[cairo_columns].copy()
        cairo_df = cairo_df[sorted(cairo_df.columns)]
        cairo_df['Total'] = cairo_df.loc[:, cairo_df.columns != 'Product name'].sum(axis=1)

        # Filter each DataFrame to drop rows where the 'Total' column is 0
        alexandria_df = alexandria_df[alexandria_df['Total'] != 0]
        ready_veg_df = ready_veg_df[ready_veg_df['Total'] != 0]
        cairo_df = cairo_df[cairo_df['Total'] != 0]

        def get_category(product_name):
            for category, products in categories_dict.items():
                if product_name in products:
                    return category
            return 'غير محدد'  # Return 'غير محدد' (Undefined) if product not found in any category

        # Create a new 'category' column in pivot_df
        alexandria_df['category'] = alexandria_df['Product name'].map(get_category)
        ready_veg_df['category'] = ready_veg_df['Product name'].map(get_category)
        cairo_df['category'] = cairo_df['Product name'].map(get_category)

        category_order = {
            "فاكهة": 1,
            "خضار": 2,
            "مجهز": 3,
            "ورقيات وأعشاب": 4
        }

        # Sort each DataFrame
        def sort_df(df):
            # Add a temporary column for sorting based on the category order
            df['category_order'] = df['category'].map(category_order)
            
            # Sort by 'category_order' and then alphabetically by 'Product' (index)
            df_sorted = df.sort_values(by=['category_order', df['Product name'].name], ascending=[True, True])
            
            # Drop the temporary 'category_order' column after sorting
            df_sorted = df_sorted.drop(columns=['category_order'])
            
            return df_sorted

        # Sort the DataFrames
        alexandria_df = sort_df(alexandria_df)
        try:
            ready_veg_df = sort_df(ready_veg_df)
        except KeyError:
            print("لا يوجد فروع للخضار الجاهز")
        cairo_df = sort_df(cairo_df)

        def to_excel_download_button(df, filename, sheet_name, label):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=True, sheet_name=sheet_name)
            buffer.seek(0)
            st.download_button(
                label=label,
                data=buffer,
                file_name=filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        # Alexandria
        st.subheader("مجمع طلبات اسكندرية")
        st.dataframe(alexandria_df)
        to_excel_download_button(
            alexandria_df,
            f"مجمع_طلبات_اسكندرية_{today_str}.xlsx",
            "Alexandria",
            "تحميل مجمع اسكندرية"
        )

        # Ready Veg
        st.subheader("مجمع طلبات خضار الجاهز")
        st.dataframe(ready_veg_df)
        to_excel_download_button(
            ready_veg_df,
            f"مجمع_طلبات_الخضار_الجاهز_{today_str}.xlsx",
            "ReadyVeg",
            "تحميل مجمع الخضار الجاهز"
        )

        # Cairo
        st.subheader("مجمع طلبات القاهرة")
        st.dataframe(cairo_df)
        to_excel_download_button(
            cairo_df,
            f"مجمع_طلبات_القاهرة_{today_str}.xlsx",
            "Cairo",
            "تحميل مجمع القاهرة"
        )
    else:
        st.info("Please upload the 216.csv file.")
if __name__ == "__main__":
    pivot_app()
