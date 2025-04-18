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
            "EG_Zakazik_DS_33": "الزقازيق"
        }


        translation_dict = {
            "Khodar.Com Molokhia 500g": "ملوخية جاهزة500جم",
            "Khodar.Com Green Beans Sliced 350g": "فاصوليا مقطعة فريش 350 جرام",
            "Khodar.com Iceberg Lettuce Shredded 350g": "كابوتشا مقطع 350 جم",
            "Khodar.com Mix Salad Cabbage Shredded 350g": "ميكس كرنب سلطة مقطع فريش 350 جرام",
            "Khodar.Com Chips Potatos 350g": "بطاطس شيبسى فريش 350 جرام",
            "Khodar.com Potatos Fries Ready to Cook 350g": "بطاطس صوابع فريش 350 جرام",
            "Khodar.com Pepper Ready for Stuffing 350g": "فلفل مقور محشي 350 جم",
            "Khodar.com Sliced Potatos 350g": "بطاطس شرائح 350 جم",
            "Khodar.Com Carrots Sliced 350g": "جزر مقطع 350 جم",
            "Khodar.com Zucchini Slices 350g": "كوسة حلقات 350جم",
            "Khodar.Com Sweet pepper 250g": "فلفل حلو 250 جرام",
            "Khodar.com Pomegranate 1kg": "رمان 1كجم",
            "Khodar.com Chestnut 250g": "ابوفروة 250جم",
            "Khodar.Com Clementine 1Kg": "يوسفي كلمنتينا 1كجم",
            "Khodar.Com Tangerine 1Kg": "يوسفي بلدي 1كجم",
            "Khodar.com Barhi Dates 500g": "بلح برحي  500 جم",
            "Khodar.com Imported Peach 500g": "خوخ مستورد 500 جم",
            "Khodar.com Imported Nectarine 500g": "نكتارين مستورد 500 جم",
            "Khodar.com Cuban Green Pepper 250g": "فلفل اخضر كوبي 250جم",
            "Khodar.com Chinese Garlic 200g": "ثوم صيني ابيض 200جم",
            "Khodar.com Peanuts 500g": "فول سوداني 500جم",
            "Khodar.com Peeled Pomegranate 350g": "رمان مفرط 350 جم",
            "Khodar.Com Ready Okra 350g": "بامية جاهزة 350 جم",
            "Khodar.Com Molokhia Ready 350g": "ملوخية جاهزة350جم",
            "Khodar.Com Cherry 250g": "كريز 250 جم",
            "Khodar.Com Mango Fons 1Kg": "مانجو فونس 1 ك",
            "Khodar.Com Mango Fas Owais 500g": "مانجو فص عويس 500 جم",
            "Khodar.Com Mango Owais 1Kg": "مانجو عويس 1 ك",
            "Khodar.Com Mango Sadiqa 1Kg": "مانجو صديقة 1ك",
            "Khodar.Com Mango Zibdia 1Kg": "مانجو زبدية 1ك",
            "Khodar.Com Red Plum Local 1Kg": "برقوق احمر محلي 1ك",
            "Khodar.Com Banati Grapes 1Kg": "عنب بناتى 1ك",
            "Khodar.Com Early Sweet White Grapes 1Kg": "عنب ايرلي سويت ابيض 1ك",
            "Khodar.Com Red Flame Grapes 1Kg": "عنب فليم احمر 1ك",
            "Khodar.Com Peeled Cane 350g": "قصب مقشر350جم",
            "Khodar.Com Ready Taro 350g": "قلقاس مكعبات فريش 350 جرام",
            "Khodar.Com Mix Stuffed 350g": "محشى مشكل فريش 350 جرام",
            "Khodar.Com Ready Squash For Stuffing 350g": "كوسة مقورة فريش 350 جرام",
            "Khodar.Com Pears African 500g": "كمثرى افريقي500 جرام",
            "Khodar.Com Capsicum Mix 500g": "فلفل الوان معبأ 500 جرام",
            "Khodar Italian Golden Apple 1kg": "تفاح اصفر ايطالى 1ك معبأ",
            "Khodar.Com Italian Golden Apple 1Kg": "تفاح أصفر إيطالي",
            "Khodar.Com Orange For juice 2Kg": "برتقال عصير 2ك معبأ",
            "Khodar.Com Coconut Pc": "جوز هند قطعة",
            "Khodar.Com Guava 1Kg": "جوافة 1ك معبأ",
            "Khodar.Com Sweet Potato 1Kg": "بطاطا 1ك",
            "Khodar.Com Orange Navel 1Kg": "برتقال بسرة 1ك",
            "Khodar.Com Potato For Fried 1Kg": "بطاطس معبأ 1ك",
            "Khodar.Com Red Onion 1Kg": "بصل احمر معبأ 1ك",
            "Khodar.Com Golden Onion 1Kg": "بصل ابيض معبأ 1ك",
            "Khodar.Com Eggplant Romi 1Kg": "باذنجان كوبى معبأ 1ك",
            "Khodar.Com Grapes Red Lebanese 500g": "عنب كريمسون لبنانى 500 جرام معبأ",
            "Khodar.Com Ready Pumpkin 350g": "قرع مكعبات صافى 350 جرام",
            "Khodar.Com Peeled Garlic Balady 125g": "عبوة ثوم مفصص 100 جرام",
            "Khodar.Com Ready Mix Vegetables 350g": "خضار  مشكل فريش 350 جرام",
            "Khodar.Com Ready Soutee Vegetables 350g": "سوتيه فريش 350 جرام",
            "Khodar.Com Ready Sweet Peas+Carrots 350g": "بسلة مفصصة بالجزر فريش 350 جرام",
            "Khodar.Com Ready Sweet Peas 350g": "بسلة مفصصة فريش 350 جرام",
            "Khodar.Com Black Grapes Lebanese 500g": "عنب اسود مستورد 500 جرام معبأ",
            "Khodar.Com Imported Banana 1kg": "موز مستورد 1ك",
            "Khodar.Com Imported Kiwi 250g": "كيوي فاخر 250 جرام معبأ",
            "Khodar.Com Italian Green Apple 1Kg": "تفاح اخضر امريكى 1ك معبأ",
            "Khodar.Com Italian Royal Gala 1Kg": "تفاح سكرى جالا 1ك معبأ",
            "Khodar.Com Italian Red Apple 1Kg": "تفاح احمر مستورد 1ك معبأ",
            "Khodar.Com Imported Red Plum 1kg": "برقوق احمر مستورد 1ك",
            "Khodar.Com Sweet Pineapple Pc": "اناناس سكري فاخر معبأ",
            "Khodar.Com Imported Avocado 500g": "افوكادو 500 جرام",
            "Khodar.Com Imported White Grape  500g": "عنب ابيض مستورد 500 جرام",
            "Khodar.Com Banana Balady 1kg": "موز بلدي فاخر 1ك معبأ",
            "Khodar.Com Cantalope 2kg": "كنتالوب 2ك معبأ",
            "Khodar.Com Coriander 100g": "كزبرة معبأ",
            "Khodar.Com French Celery PC": "كرفس فرنساوي 250 جرام",
            "Khodar.Com DILL 100g": "شبت معبأ",
            "Khodar.Com Thyme 50g": "زعتر فريش معبأ",
            "Khodar.Com Basil 50g": "ريحان اخضر معبأ",
            "Khodar.Com Rosemary 50g": "روزمارى فريش معبأ",
            "Khodar.Com Watercress 100g": "جرجير معبأ",
            "Khodar.Com Parsley 100g": "بقدونس معبأ",
            "Khodar.Com Mushroom 200g": "مشروم 200 جرام معبأ",
            "Khodar.Com Red Cabbage Pc": "كرنب احمر سلطة معبأ",
            "Khodar.Com White Cabbage Pc": "كرنب ابيض سلطة معبأ",
            "Khodar.Com Iceberg Lettuce pc": "كابوتشى معبأ",
            "Khodar.Com Ginger 100g": "زنجبيل 100 جرام معبأ",
            "Khodar.Com Sweet Corn 2pc": "ذرة سكري 2 قطعه",
            "Khodar.Com Romaine Lettuce 1 Piece": "خس بلدي فاخر معبأ",
            "Khodar.Com Romaine Lettuce pc": "خس بلدي فاخر معبأ",
            "Khodar.Com Green onion 125g": "بصل اخضر معبأ",
            "Khodar.Com Lemon Balady 250g": "ليمون بلدى فاخر معبأ 250 جرام",
            "Khodar.Com Lemon Adalia 250g": "ليمون اضاليا 250 جرام",
            "Khodar.Com Zucchini 500g": "كوسة معبأ 500 جرام",
            "Khodar.Com Leek 250g": "كرات 250 جرام",
            "Khodar.Com Cauliflower 500g": "قرنبيط 500 جرام",
            "Khodar.Com Pepper Hot Green 250g": "فلفل اخضر حار معبأ 250 جرام",
            "Khodar.Com Red Radish 500g": "فجل احمر 500 جرام",
            "Khodar.Com Tomato 1kg": "طماطم فاخر معبأ 1ك",
            "Khodar.Com Cherry Tomato 250g": "طماطم شيرى معبأ 250 جرام",
            "Khodar.Com Cucumber 1kg": "خيار فاخر معبأ 1ك",
            "Khodar.Com Carrots 500g": "جزر  معبأ 500 جرام",
            "Khodar.Com Beet Root 500g": "بنجر احمر معبأ 500 جرام",
            "Khodar.Com Broccoli 500g": "بروكلي 500 جرام",
            "Khodar.Com Black Eggplant 500g": "باذنجان عروس اسود معبأ 500 جرام",
            "Khodar.Com Black Grapes 1Kg": "عنب اسود 1ك",
            "Khodar.Com White Eggplant 500g": "باذنجان عروس ابيض معبأ 500 جرام",
            "Khodar.Com Pepper hot Red 250g": "فلفل حار احمر 250 جرام",
            "Khodar.Com Strawberry 250g": "فراوله 250 جرام",
            "Khodar.Com Pears Lebanese 500g": "كمثري لبناني 500 جرام",
            "Khodar.Com Peeled Haranksh 250g": "حرنكش مقشر 250 جرام",
            "khodar.com Watermelon Pc": "بطيخ",
            "khodar.Com Sugar peach 1Kg": "خوخ سكرى",
            "Khodar.com Fresh Peas 500g": "بسلة 500 جم",
            "Khodar.com Fresh Green Beans 500g": "فاصوليا خضراء 500جم",
            "Khodar.com White Grape Fruit 1kg": "جريب فروت ابيض 1كجم",
            "Khodar.com Red Grape Fruit 1kg": "جريب فروت احمر 1كجم",
            "Khodar.com Local Peach 1kg": "خوخ محلي 1كجم",
            "Khodar.com Tangerine Christina 1kg": "يوسفي كريستينا 1كجم",
            "Khodar.Com Sweet Melon 1kg": "شمام شهد 1ك معبأ",
            "khodar.com Watermelon Pc":	"بطيخ",
            "khodar.com Watermelon 1Pc 6-8Kg": "بطيخ",
            "khodar.Com Red Watermelon Seedless Pc": "بطيخ احمر بدون بذر",
            "khodar.Com Yellow Watermelon Seedless Pc":	"بطيخ اصفر بدون بذر"
        }
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

        # Create the pivot table
        pivot_df = df.pivot_table(
            index='Product',
            columns='Store_name',
            values='Effective quantity',
            aggfunc='sum',
            fill_value=0
        )

        # Make a lowercase version of the translation dictionary keys
        lower_translation_dict = {k.lower(): v for k, v in translation_dict.items()}

        # Map using lowercase comparison
        pivot_df.index = pivot_df.index.map(lambda x: lower_translation_dict.get(x.lower(), x))

        pivot_df.columns = pivot_df.columns.map(lambda x: branches_dict.get(x, x))
        pivot_df = pivot_df[sorted(pivot_df.columns)]
        # Define the column groups
        alexandria_columns = ['سيدي بشر', 'الابراهيميه', 'وينجت']
        ready_veg_columns = ['المعادي لاسلكي', 'الدقي', 'زهراء المعادي', 'ميدان لبنان', 'العجوزة', 'كورنيش المعادي', 'زهراء المعادي - 2']
        cairo_columns = [col for col in pivot_df.columns if col not in alexandria_columns and col not in ready_veg_columns]

        # Ensure all values are numeric for summing
        pivot_df = pivot_df.apply(pd.to_numeric, errors='coerce')

        # Create the Alexandria DataFrame (keeping 'Product' as the index)
        alexandria_df = pivot_df[alexandria_columns].copy()
        alexandria_df = alexandria_df[sorted(alexandria_df.columns)]
        alexandria_df['Total'] = alexandria_df.sum(axis=1)

        # Create the Ready Veg DataFrame (keeping 'Product' as the index)
        ready_veg_df = pivot_df[ready_veg_columns].copy()
        ready_veg_df = ready_veg_df[sorted(ready_veg_df.columns)]
        ready_veg_df['Total'] = ready_veg_df.sum(axis=1)

        # Create the Cairo DataFrame (keeping 'Product' as the index)
        cairo_df = pivot_df[cairo_columns].copy()
        cairo_df = cairo_df[sorted(cairo_df.columns)]
        cairo_df['Total'] = cairo_df.sum(axis=1)

        # Filter each DataFrame to drop rows where all column values are 0
        alexandria_df = alexandria_df.loc[(alexandria_df != 0).any(axis=1)]
        
        ready_veg_df = ready_veg_df.loc[(ready_veg_df != 0).any(axis=1)]
        cairo_df = cairo_df.loc[(cairo_df != 0).any(axis=1)]
        def get_category(product_name):
            for category, products in categories_dict.items():
                if product_name in products:
                    return category
            return 'غير محدد'  # Return 'غير محدد' (Undefined) if product not found in any category

        # Create a new 'category' column in pivot_df
        alexandria_df['category'] = alexandria_df.index.map(get_category)
        ready_veg_df['category'] = ready_veg_df.index.map(get_category)
        cairo_df['category'] = cairo_df.index.map(get_category)

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
            df_sorted = df.sort_values(by=['category_order', df.index.name], ascending=[True, True])
            
            # Drop the temporary 'category_order' column after sorting
            df_sorted = df_sorted.drop(columns=['category_order'])
            
            return df_sorted

        # Sort the DataFrames
        alexandria_df = sort_df(alexandria_df)
        ready_veg_df = sort_df(ready_veg_df)
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
