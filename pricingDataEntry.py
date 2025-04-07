import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd


def pricing_app():
    st.title("Pricing Data Entry")

    # Constants
    products = [
        "تفاح احمر مستورد", "شمام شهد", "عنب ايرلى سويت", "عنب ابيض افريقى", "فول سادة - مطبخ",
        "برتقال عصير تصدير", "بطيخ", "جوز هند", "افوكادو", "ليمون بلدى", "مشروم 200 جرام", "بصل احمر",
        "برتقال بسرة", "خوخ سكرى", "برقوق احمر مستورد", "بطاطس تحمير", "بسلة", "بطاطا", "بنجر", "جزر",
        "تفاح اخضر دايت", "يوسفى موركت", "يوسفى موركيت اسبانى", "قلقاس", "رمان", "نعناع", "ثوم بلدى بدون عرش",
        "فاصوليا خضراء", "ملوخية", "كوسة", "بروكلى", "روزمارى فريش", "كرفس", "كرفس فرنساوى", "تفاح اصفر مستورد",
        "تفاح سكرى جالا", "ذرة سكرى", "كزبرة", "سلق", "عنب اسود لبنانى", "خوخ فلوريدا", "كمثرى افريقى فاخر",
        "بامية تركى", "زنجبيل فريش", "كنتالوب", "طماطم شيرى", "كيوى", "سبانخ", "فلفل الوان", "موز مستورد فاخر",
        "باذنجان عروس اسود", "فلفل حار", "باذنجان عروس ابيض", "فلفل حلو", "باذنجان كوبى", "فلفل اخضر كوبى",
        "فلفل كاليفورنيا", "بصل ابيض", "قرنبيط", "كابوتشى", "طماطم", "خيار", "كرنب ابيض سلطة", "كرنب احمر سلطة",
        "ليمون اضاليا", "فجل احمر بدون عرش", "فلفل احمر حار", "كرات فرنساوى", "بصل اخضر", "ثوم مستورد باللفة",
        "خس بلدى", "شبت", "بقدونس", "جرجير", "ريحان", "زعتر فريش", "جريب فروت ابيض", "جوافة بلدى فاخر",
        "موز بلدى", "فول سودانى محمص بقشره بالملح", "فراولة", "يوسفى بلدى", "ابو فروة", "اناناس سكرى",
        "عنب فليم احمر", "مشمش بلدى", "برقوق احمر محلى", "عنب كريمسون", "كمثرى لبنانى فاخر", "برقوق اصفر محلى",
        "قرع عسل", "حرنكش", "مانجو صديقة", "برقوق هوليود", "مانجو زبدة", "قصب مقشر 350 جرام", "مشمش لبنانى",
        "كريز", "مانجو عويس", "عنب اسود محلى", "عنب اصفر بناتى", "مانجو فص", "مانجو الفونس", "مانجو نعومى",
        "مانجو تيمور", "كمثرى بلدى خشابى", "بلح برحى", "مانجو كيت", "خوخ مستورد", "خوخ نكتارين مستورد",
        "قشطة عبد الرازق", "قشطة بلدى", "كاكا", "جريب فروت احمر", "خرشوف", "ورق عنب", "كرنب بلدى",
        "تين شوكى بالواحدة", "يوسفي كلمنتينا", "برقوق اصفر مستورد", "بلح عراقي", "خوخ محلي", "يوسفي كريستينا"
    ]

    providers = [
        "ابانوب يحيي", "محمد هلال ( ابو كريم )", "احمد التوام", "احمد بسيوني", "احمد كمال",
        "اسماعيل خضر", "جمال عبد الحكيم", "حاج هلال", "حسين عبد المغنى", "حماده الديب",
        "حمدى احمد", "عبد الله جمال", "علي زهران", "مجدى", "محمد ذكى", "سعيد ورقيات",
        "سيد ابو ياسين", "عبد الله العمدة", "عادل عبدالمغني", "احمد دهشان", "حافظ ابو حليمه",
        "رمضان العمدة", "اجرينال", "محمد عبد الكريم", "شركة الانوار", "نقدى", "احمد بسيونى",
        "احمد زلط", "اشرف ابو عنيزة", "اشرف عبد الجواد", "البدرى توفيق", "الشيخ محمد",
        "ايمن دنقل فرع", "ثروت وطلعت", "حمدى احمد بطيخ", "خالد محمد السيد", "خلف ابو الليل",
        "سكر", "شركة الحمد محمد زغلوله", "شعبان بسيونى", "عادل برعى", "عثمان", "مؤمن ومصطفى",
        "محمد احمد عبد الغنى", "محمد الخياط", "وليد عزوز", "شركة الفتح مشروم", "شركة اجرينال"
    ]

    branches = ['Cairo', "Alexandria"]
    # Establishing a Google Sheets connection
    conn = st.connection("gsheets", type=GSheetsConnection)


    branch = st.selectbox("اختر الفرع", ['Cairo', 'Alexandria'])
    # 🛠 Read existing data
    existing_data = conn.read(worksheet=branch, usecols=list(range(10)), ttl=5).dropna(how="all")

    action = st.selectbox(
        "Choose an Action",
        [
            "تحديث السعر",
            "عرض كل الأصناف",
            "حذف الصنف",
        ],
    )

    if action == "تحديث السعر":
        st.markdown("### اختر الصنف والمورد وتاريخ الشراء لتحديث السعر")

        # Read and clean existing data
        existing_data = conn.read(worksheet=branch, usecols=list(range(10)), ttl=5).dropna(how="all")

        # Ensure correct date format
        existing_data["تاريخ الشراء"] = pd.to_datetime(     existing_data["تاريخ الشراء"], dayfirst=True, errors="coerce" ).dt.strftime("%Y-%m-%d")

        # Select product to update
        product_to_update = st.selectbox("اختر الصنف", options=existing_data["اسم الصنف"].unique().tolist(), index=None)

        if product_to_update:
            provider_to_update = st.selectbox(
                "اختر المورد", 
                options=existing_data[existing_data["اسم الصنف"] == product_to_update]["مورد الشركة"].unique().tolist(), 
                index=None
            )

            if provider_to_update:
                purchase_date_to_update = st.selectbox(
                    "اختر تاريخ الشراء", 
                    options=existing_data[
                        (existing_data["اسم الصنف"] == product_to_update) & 
                        (existing_data["مورد الشركة"] == provider_to_update)
                    ]["تاريخ الشراء"].unique().tolist(),
                    index=None
                )

                if purchase_date_to_update:
                    # Retrieve selected row for editing
                    selected_row = existing_data[
                        (existing_data["اسم الصنف"] == product_to_update) & 
                        (existing_data["مورد الشركة"] == provider_to_update) & 
                        (existing_data["تاريخ الشراء"] == purchase_date_to_update)
                    ]

                    if not selected_row.empty:
                        selected_row = selected_row.iloc[0]  # Get the first matching row

                        # Display current row data dynamically before editing
                        st.markdown("### البيانات الحالية:")
                        st.table(selected_row[["اسم الصنف", "مورد الشركة", "تاريخ الشراء", "الاجمالى", "وزن صافي", "تكلفة الوحدة"]])

                        with st.form(key="price_update_form"):
                            total_price = st.number_input(
                                "الاجمالى", min_value=0.0, value=float(selected_row["الاجمالى"])
                            )

                            update_button = st.form_submit_button(label="تحديث السعر")

                            if update_button:
                                # Update the selected row with new price
                                existing_data.loc[selected_row.name, "الاجمالى"] = total_price
                                net_weight = existing_data.loc[selected_row.name, "وزن صافي"]
                                existing_data.loc[selected_row.name, "تكلفة الوحدة"] = total_price / net_weight

                                # Update the worksheet
                                conn.update(worksheet=branch, data=existing_data)
                                st.success("تم تحديث السعر بنجاح!")
                    else:
                        st.warning("لم يتم العثور على بيانات مطابقة.")


    # عرض كل الأصناف
    elif action == "عرض كل الأصناف":
        st.dataframe(existing_data)
    # حذف الصنف
    elif action == "حذف الصنف":
        st.markdown("اختر الصنف والمورد وتاريخ الشراء لحذفه")

        # Read and clean existing data
        existing_data = conn.read(worksheet=branch, usecols=list(range(10)), ttl=5).dropna(how="all")

        # Ensure correct date format
        existing_data["تاريخ الشراء"] = pd.to_datetime(     existing_data["تاريخ الشراء"], dayfirst=True, errors="coerce" ).dt.strftime("%Y-%m-%d")

        # Select product to delete
        product_to_delete = st.selectbox("اختر الصنف", options=existing_data["اسم الصنف"].unique().tolist(), index=None)
        
        if product_to_delete:
            provider_to_delete = st.selectbox(
                "اختر المورد", 
                options=existing_data[existing_data["اسم الصنف"] == product_to_delete]["مورد الشركة"].unique().tolist(), 
                index=None
            )

            if provider_to_delete:
                purchase_date_to_delete = st.selectbox(
                    "اختر تاريخ الشراء", 
                    options=existing_data[
                        (existing_data["اسم الصنف"] == product_to_delete) & 
                        (existing_data["مورد الشركة"] == provider_to_delete)
                    ]["تاريخ الشراء"].unique().tolist(),
                    index=None
                )

                if purchase_date_to_delete:
                    # Display confirmation message
                    st.warning(f"هل أنت متأكد أنك تريد حذف {product_to_delete} من {provider_to_delete} بتاريخ {purchase_date_to_delete}؟")

                    if st.button("حذف الصنف"):
                        # Remove the selected row
                        existing_data = existing_data[
                            ~(
                                (existing_data["اسم الصنف"] == product_to_delete) &
                                (existing_data["مورد الشركة"] == provider_to_delete) &
                                (existing_data["تاريخ الشراء"] == purchase_date_to_delete)
                            )
                        ]

                        # Update worksheet
                        conn.update(worksheet=branch, data=existing_data)
                        st.success("تم حذف الصنف بنجاح!")
if __name__ == "__main__":
    pricing_app()
