import os
import re
import zipfile
import datetime
import tempfile
import pytz
import pandas as pd
import smtplib
from streamlit_gsheets import GSheetsConnection
import streamlit as st
import gspread
from docx import Document
from email.message import EmailMessage

def dashboardApp():
    multi_items_products = [{
            'talabat product': 'خضار.كوم ميكس كرنب سلطة مقطع فريش 350 جرام',
            'data entry product': ['كرنب احمر سلطة', 'كرنب ابيض سلطة'],
            'نسبة الفرزة': 0.0,
            'التعبئه': 0.0,
            'فرق وزن': 0.0,
            'نسبة مرتجع': 0.0,
            'منتج المواد': ['كرنب احمر سلطة', 'كرنب ابيض سلطة'],
            'كمية المنتج في المواد': [0.25, 0.25] 
        },
    {
    'talabat product': 'محشى مشكل فريش 350 جرام',
    'data entry product': ['باذنجان عروس اسود', 'كوسة', 'فلفل اخضر كوبى'],
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': ['باذنجان عروس اسود', 'كوسة', 'فلفل اخضر كوبى'],
    'كمية المنتج في المواد': [0.85, 0.75, 0.1]},
    {
    'talabat product': 'خضار مشكل فريش 350 جرام',
    'data entry product': ['بسلة', 'كوسة', 'جزر'],
    'نسبة الفرزة': 0.2,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': ['بسلة', 'كوسة', 'جزر'],
    'كمية المنتج في المواد': [0.08, 0.15, 0.15]},
    {
    'talabat product': 'سوتيه فريش 350 جرام',
    'data entry product': ['بسلة', 'كوسة', 'جزر', 'فاصوليا خضراء', 'بروكلى', 'روزمارى فريش', 'مشروم 200 جرام'],
    'نسبة الفرزة': 0.2,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': ['بسلة', 'كوسة', 'جزر', 'فاصوليا خضراء', 'بروكلى', 'روزمارى فريش', 'مشروم 200 جرام'],
    'كمية المنتج في المواد': [0.08, 0.15, 0.15, 0.08, 0.05, 1, 0.25]},
    {
    'talabat product': 'بسلة مفصصة بالجزر فريش 350 جرام',
    'data entry product': ['جزر', 'بسلة'],
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': ['جزر', 'بسلة'],
            'كمية المنتج في المواد': [0.15, 0.5]}]

    single_item_products = [{
    'talabat product': 'ملوخية جاهزة500جم',
    'data entry product': 'ملوخية',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'ملوخية مخرطة فريش 500جرام',
    'كمية المنتج في المواد': 0.65},
    {
    'talabat product': 'خضار.كوم فاصوليا مقطعة فريش 350 جرام',
    'data entry product': 'فاصوليا خضراء',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'فاصوليا مقطعة فريش 350 جرام',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'خضار.كوم كابوتشا مقطع 350 جم',
    'data entry product': 'كابوتشى',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'كابوتشي مقطع فريش 350جرام',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'خضار.كوم بطاطس شيبسى فريش 350 جرام',
    'data entry product': 'بطاطس تحمير',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'بطاطس شيبسى فريش 350 جرام',
    'كمية المنتج في المواد': 0.45},
    {
    'talabat product': 'خضار.كوم بطاطس صوابع فريش 350 جرام',
    'data entry product': 'بطاطس تحمير',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'بطاطس صوابع فريش 350 جرام',
    'كمية المنتج في المواد': 0.45},
    {
    'talabat product': 'خضار.كوم فلفل مقور محشي 350 جم',
    'data entry product': 'فلفل اخضر كوبى',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'فلفل اخضر مقور محشى 350جم',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'خضار.كوم بطاطس شرائح 350 جم',
    'data entry product': 'بطاطس تحمير',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'بطاطس شرائح 350 جرام',
    'كمية المنتج في المواد': 0.45},
    {
    'talabat product': 'خضار.كوم جزر مقطع 350 جم',
    'data entry product': 'جزر',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'جزر مقطع فريش 350 جم',
    'كمية المنتج في المواد': 0.65},
    {
    'talabat product': 'خضار.كوم كوسة حلقات 350جم',
    'data entry product': 'كوسة',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'كوسة حلقات فريش 350جرام',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'خضار.كوم فلفل حلو 250 جرام',
    'data entry product': 'فلفل حلو',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'فلفل حلو معبأ 250 جرام',
    'كمية المنتج في المواد': 0.3},
    {
    'talabat product': 'خضار.كوم رمان 1كجم',
    'data entry product': 'رمان',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'رمان 1ك معبأ',
    'كمية المنتج في المواد': 1.1},
    {
    'talabat product': 'خضار.كوم ابوفروة 250جم',
    'data entry product': 'ابو فروة',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'ابوفروة 250 جرام معبأ',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'خضار.كوم يوسفي كلمنتينا 1كجم',
    'data entry product': 'يوسفي كلمنتينا',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': '',
    'كمية المنتج في المواد': ''},
    {
    'talabat product': 'خضار.كوم يوسفي بلدي 1كجم',
    'data entry product': 'يوسفى بلدى',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'يوسفي بلدى 1ك معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'خضار.كوم بلح برحي 500 جم',
    'data entry product': 'بلح برحى',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'بلح برحي 500 جرام معبأ',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'خضار.كوم نكتارين مستورد 500 جم',
    'data entry product': 'خوخ نكتارين مستورد',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'خوخ مستورد 500جرام معبأ',
    'كمية المنتج في المواد': 0.55},
    {
    'talabat product': 'خضار.كوم فلفل اخضر كوبي 250جم',
    'data entry product': 'فلفل اخضر كوبى',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'فلفل اخضر كوبى معبأ 250 جرام',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'خضار.كوم فول سوداني 500جم',
    'data entry product': 'فول سودانى محمص بقشره بالملح',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'فول سودانى محمص بقشره بالملح 500 جرام معبأ',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'رمان مفرط 350 جم',
    'data entry product': 'رمان',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'رمان مفرط 350 جرام',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'بامية جاهزة 350 جم',
    'data entry product': 'بامية تركى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بامية تركى مقمعة فريش 350 جرام',
    'كمية المنتج في المواد': 0.7},
    {
    'talabat product': 'ملوخية جاهزة350جم',
    'data entry product': 'ملوخية',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'ملوخية متقطفة فريش 350 جرام',
    'كمية المنتج في المواد': 1.1},
    {
    'talabat product': 'كريز 250 جم',
    'data entry product': 'كريز',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'كريز250 جرام معبأ',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'مانجو فونس 1 ك',
    'data entry product': 'مانجو الفونس',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'مانجو الفونس 1ك معبأ',
    'كمية المنتج في المواد': 1.05},
    {
    'talabat product': 'مانجو فص عويس 500 جم',
    'data entry product': 'مانجو فص',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'مانجو فص 500 جرام معبأ',
    'كمية المنتج في المواد': 0.6},
    {
    'talabat product': 'مانجو عويس 1 ك',
    'data entry product': 'مانجو عويس',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'مانجو عويس 1ك معبأ',
    'كمية المنتج في المواد': 1.1},
    {
    'talabat product': 'مانجو صديقة 1ك',
    'data entry product': 'مانجو صديقة',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'مانجو صديقة 1ك معبأ',
    'كمية المنتج في المواد': 1.2},
    {
    'talabat product': 'مانجو زبدية 1ك',
    'data entry product': 'مانجو زبدة',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'مانجو زبدة 1ك معبأ',
    'كمية المنتج في المواد': 1.2},
    {
    'talabat product': 'برقوق احمر محلي 1ك',
    'data entry product': 'برقوق احمر محلى',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'برقوق احمرمصرى 1ك معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'عنب بناتى 1ك',
    'data entry product': 'عنب اصفر بناتى',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'عنب اصفر بناتى 1ك معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'عنب فليم احمر 1ك',
    'data entry product': 'عنب فليم احمر',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'عنب فليم احمر 1ك معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'قصب مقشر350جم',
    'data entry product': 'قصب مقشر 350 جرام',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'قصب مقشر 350 جرام معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'قلقاس مكعبات فريش 350 جرام',
    'data entry product': 'قلقاس',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'قلقاس مكعبات فريش 350 جرام',
    'كمية المنتج في المواد': 0.6},
    {
    'talabat product': 'كوسة مقورة فريش 350 جرام',
    'data entry product': 'كوسة',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كوسة مقورة فريش 350 جرام',
    'كمية المنتج في المواد': 0.6},
    {
    'talabat product': 'كمثرى افريقي500 جرام',
    'data entry product': 'كمثرى افريقى فاخر',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كمثرى افريقى 500 جرام معبأ',
    'كمية المنتج في المواد': 0.6},
    {
    'talabat product': 'فلفل الوان معبأ 500 جرام',
    'data entry product': 'فلفل الوان',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'فلفل الوان معبأ 500 جرام',
    'كمية المنتج في المواد': 0.6},
    {
    'talabat product': 'تفاح اصفر ايطالى 1ك معبأ',
    'data entry product': 'تفاح اصفر مستورد',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'تفاح اصفر ايطالى 1ك معبأ',
    'كمية المنتج في المواد': 1.07},
    {
    'talabat product': 'برتقال عصير 2ك معبأ',
    'data entry product': 'برتقال عصير تصدير',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'برتقال عصير 2ك معبأ',
    'كمية المنتج في المواد': 2.1},
    {
    'talabat product': 'جوز هند قطعة',
    'data entry product': 'جوز هند',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'جوزهند معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'يوسفي موركت 1ك',
    'data entry product': 'يوسفى موركت',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'يوسفى موركيت اسبانى',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'جوافة 1ك معبأ',
    'data entry product': 'جوافة بلدى فاخر',
    'نسبة الفرزة': 0.1,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'جوافة 1ك معبأ',
    'كمية المنتج في المواد': 1.05},
    {
    'talabat product': 'بطاطا 1ك',
    'data entry product': 'بطاطا',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بطاطا 1ك معبأ',
    'كمية المنتج في المواد': 1.15},
    {
    'talabat product': 'برتقال بسرة 1ك',
    'data entry product': 'برتقال بسرة',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'برتقال بسرة 1ك معبأ',
    'كمية المنتج في المواد': 1.07},
    {
    'talabat product': 'بطاطس معبأ 1ك',
    'data entry product': 'بطاطس تحمير',
    'نسبة الفرزة': 0.07,
    'التعبئه': 1.34,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بطاطس معبأ 1ك',
    'كمية المنتج في المواد': 1.05},
    {
    'talabat product': 'بصل احمر معبأ 1ك',
    'data entry product': 'بصل احمر',
    'نسبة الفرزة': 0.07,
    'التعبئه': 1.34,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بصل احمر معبأ 1ك',
    'كمية المنتج في المواد': 1.05},
    {
    'talabat product': 'بصل ابيض معبأ 1ك',
    'data entry product': 'بصل ابيض',
    'نسبة الفرزة': 0.07,
    'التعبئه': 1.34,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بصل ابيض معبأ 1ك',
    'كمية المنتج في المواد': 1.05},
    {
    'talabat product': 'باذنجان كوبى معبأ 1ك',
    'data entry product': 'باذنجان كوبى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'باذنجان كوبى معبأ 1ك',
    'كمية المنتج في المواد': 1.1},
    {
    'talabat product': 'قرع مكعبات صافى 350 جرام',
    'data entry product': 'قرع عسل',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'قرع مكعبات صافى 350 جرام',
    'كمية المنتج في المواد': 0.7},
    {
    'talabat product': 'عبوة ثوم مفصص 100 جرام',
    'data entry product': 'ثوم بلدى بدون عرش',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'عبوة ثوم مفصص 100 جرام',
    'كمية المنتج في المواد': 0.1},
    {
    'talabat product': 'بسلة مفصصة فريش 350 جرام',
    'data entry product': 'بسلة',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'بسلة مفصصة فريش 350 جرام',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'عنب اسود مستورد 500 جرام معبأ',
    'data entry product': 'عنب اسود لبنانى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'عنب اسود مستورد 500 جرام معبأ',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'موز مستورد 1ك',
    'data entry product': 'موز مستورد فاخر',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'موز اكوادورى فاخر 1ك معبأ',
    'كمية المنتج في المواد': 1.1},
    {
    'talabat product': 'كيوي فاخر 250 جرام معبأ',
    'data entry product': 'كيوى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 1.75,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كيوي فاخر 250 جرام معبأ',
    'كمية المنتج في المواد': 0.27},
    {
    'talabat product': 'تفاح اخضر امريكى 1ك معبأ',
    'data entry product': 'تفاح اخضر دايت',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'تفاح اخضر امريكى 1ك معبأ',
    'كمية المنتج في المواد': 1.07},
    {
    'talabat product': 'تفاح سكرى جالا 1ك معبأ',
    'data entry product': 'تفاح سكرى جالا',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'تفاح سكرى جالا 1ك معبأ',
    'كمية المنتج في المواد': 1.07},
    {
    'talabat product': 'تفاح احمر مستورد 1ك معبأ',
    'data entry product': 'تفاح احمر مستورد',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'تفاح احمر مستورد 1ك معبأ',
    'كمية المنتج في المواد': 1.07},
    {
    'talabat product': 'برقوق احمر مستورد 1ك',
    'data entry product': 'برقوق احمر مستورد',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'برقوق احمر مستورد 1ك معبأ',
    'كمية المنتج في المواد': 1.03},
    {
    'talabat product': 'اناناس سكري فاخر معبأ',
    'data entry product': 'اناناس سكرى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.25,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'اناناس سكري فاخر معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'افوكادو 500 جرام',
    'data entry product': 'افوكادو',
    'نسبة الفرزة': 0.07,
    'التعبئه': 2.0,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'افوكادو مستورد 500 جرام معبأ',
    'كمية المنتج في المواد': 0.6},
    {
    'talabat product': 'عنب ابيض مستورد 500 جرام',
    'data entry product': 'عنب ابيض افريقى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'عنب ابيض مستورد 500 جرام معبأ',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'موز بلدي فاخر 1ك معبأ',
    'data entry product': 'موز بلدى',
    'نسبة الفرزة': 0.15,
    'التعبئه': 2.25,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'موز بلدي فاخر 1ك معبأ',
    'كمية المنتج في المواد': 1.1},
    {
    'talabat product': 'كنتالوب 2ك معبأ',
    'data entry product': 'كنتالوب',
    'نسبة الفرزة': 0.07,
    'التعبئه': 1.34,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كنتالوب 2ك معبأ',
    'كمية المنتج في المواد': 2.2},
    {
    'talabat product': 'كزبرة معبأ',
    'data entry product': 'كزبرة',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كزبرة معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'كرفس فرنساوي 250 جرام',
    'data entry product': 'كرفس فرنساوى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كرفس فرنساوى معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'شبت معبأ',
    'data entry product': 'شبت',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'شبت معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'زعتر فريش معبأ',
    'data entry product': 'زعتر فريش',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'زعتر فريش معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'ريحان اخضر معبأ',
    'data entry product': 'ريحان',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'ريحان اخضر معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'روزمارى فريش معبأ',
    'data entry product': 'روزمارى فريش',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'روزمارى فريش معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'جرجير معبأ',
    'data entry product': 'جرجير',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'جرجير معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'بقدونس معبأ',
    'data entry product': 'بقدونس',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بقدونس معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'مشروم 200 جرام معبأ',
    'data entry product': 'مشروم 200 جرام',
    'نسبة الفرزة': 0.15,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'مشروم 200 جرام معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'كرنب احمر سلطة معبأ',
    'data entry product': 'كرنب احمر سلطة',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كرنب احمر سلطة معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'كرنب ابيض سلطة معبأ',
    'data entry product': 'كرنب ابيض سلطة',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كرنب ابيض سلطة معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'كابوتشى معبأ',
    'data entry product': 'كابوتشى',
    'نسبة الفرزة': 0.15,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كابوتشى معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'زنجبيل 100 جرام معبأ',
    'data entry product': 'زنجبيل فريش',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'زنجبيل 100 جرام معبأ',
    'كمية المنتج في المواد': 0.1},
    {
    'talabat product': 'ذرة سكري 2 قطعه',
    'data entry product': 'ذرة سكرى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'ذرة سكري 2 قطعة معبأ',
    'كمية المنتج في المواد': 2},
    {
    'talabat product': 'خس بلدي فاخر معبأ',
    'data entry product': 'خس بلدى',
    'نسبة الفرزة': 0.1,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'خس بلدي فاخر معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'بصل اخضر معبأ',
    'data entry product': 'بصل اخضر',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.73,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بصل اخضر معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'ليمون بلدى فاخر معبأ 250 جرام',
    'data entry product': 'ليمون بلدى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'ليمون بلدى فاخر معبأ 250 جرام',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'ليمون اضاليا 250 جرام',
    'data entry product': 'ليمون اضاليا',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'ليمون اضاليا معبأ 250جرام',
    'كمية المنتج في المواد': 0.4},
    {
    'talabat product': 'كوسة معبأ 500 جرام',
    'data entry product': 'كوسة',
    'نسبة الفرزة': 0.05,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'كوسة معبأ 500 جرام',
    'كمية المنتج في المواد': 0.6},
    {
    'talabat product': 'فلفل اخضر حار معبأ 250 جرام',
    'data entry product': 'فلفل حار',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'فلفل اخضر حار معبأ 250 جرام',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'فجل احمر 500 جرام',
    'data entry product': 'فجل احمر بدون عرش',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'فجل احمر معبأ 500 جرام',
    'كمية المنتج في المواد': 0.55},
    {
    'talabat product': 'طماطم فاخر معبأ 1ك',
    'data entry product': 'طماطم',
    'نسبة الفرزة': 0.15,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'طماطم فاخر معبأ 1ك',
    'كمية المنتج في المواد': 1.07},
    {
    'talabat product': 'طماطم شيرى معبأ 250 جرام',
    'data entry product': 'طماطم شيرى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'طماطم شيرى معبأ 250 جرام',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'خيار فاخر معبأ 1ك',
    'data entry product': 'خيار',
    'نسبة الفرزة': 0.2,
    'التعبئه': 2.5,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'خيار فاخر معبأ 1ك',
    'كمية المنتج في المواد': 1.1},
    {
    'talabat product': 'جزر معبأ 500 جرام',
    'data entry product': 'جزر',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'جزر معبأ 500 جرام',
    'كمية المنتج في المواد': 0.55},
    {
    'talabat product': 'بنجر احمر معبأ 500 جرام',
    'data entry product': 'بنجر',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بنجر احمر معبأ 500 جرام',
    'كمية المنتج في المواد': 0.55},
    {
    'talabat product': 'بروكلي 500 جرام',
    'data entry product': 'بروكلى',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'بروكلي 500 جرام',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'باذنجان عروس اسود معبأ 500 جرام',
    'data entry product': 'باذنجان عروس اسود',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'باذنجان عروس اسود معبأ 500 جرام',
    'كمية المنتج في المواد': 0.55},
    {
    'talabat product': 'عنب اسود 1ك',
    'data entry product': 'عنب اسود محلى',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'عنب اسود مصري 1ك معبأ',
    'كمية المنتج في المواد': 1},
    {
    'talabat product': 'باذنجان عروس ابيض معبأ 500 جرام',
    'data entry product': 'باذنجان عروس ابيض',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.95,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'باذنجان عروس ابيض معبأ 500 جرام',
    'كمية المنتج في المواد': 0.55},
    {
    'talabat product': 'فلفل حار احمر 250 جرام',
    'data entry product': 'فلفل احمر حار',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'فلفل احمر حار معبأ 250 جرام',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'فراوله 250 جرام',
    'data entry product': 'فراولة',
    'نسبة الفرزة': 0.07,
    'التعبئه': 0.87,
    'فرق وزن': 0.07,
    'نسبة مرتجع': 0.04,
    'منتج المواد': 'فراوله فاخر 250 جرام معبأ',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'كمثري لبناني 500 جرام',
    'data entry product': 'كمثرى لبنانى فاخر',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'كمثري لبنانى 500 جرام معبأ',
    'كمية المنتج في المواد': 0.5},
    {
    'talabat product': 'حرنكش مقشر 250 جرام',
    'data entry product': 'حرنكش',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'حرنكش مقشر 250 جرام',
    'كمية المنتج في المواد': 0.25},
    {
    'talabat product': 'برقوق اصفر مستورد 1ك',
    'data entry product': 'برقوق اصفر مستورد',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': '',
    'كمية المنتج في المواد': ''},
    {
    'talabat product': 'بلح عراقي 1ك',
    'data entry product': 'بلح عراقي',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': '',
    'كمية المنتج في المواد': ''},
    {
    'talabat product': 'خضار.كوم بسلة 500 جم',
    'data entry product': 'بسلة',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'بسلة 500 جم معبأ',
    'كمية المنتج في المواد': 0.55},
    {
    'talabat product': 'خضار.كوم فاصوليا خضراء 500جم',
    'data entry product': 'فاصوليا خضراء',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'فاصوليا خضراء 500جم معبأ',
    'كمية المنتج في المواد': 1.03},
    {
    'talabat product': 'خضار.كوم جريب فروت ابيض 1كجم',
    'data entry product': 'جريب فروت ابيض',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'جريب فروت ابيض 1ك معبأ',
    'كمية المنتج في المواد': 1.07},
    {
    'talabat product': 'خضار.كوم جريب فروت احمر 1كجم',
    'data entry product': 'جريب فروت احمر',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'جريب فروت احمر 1ك معبأ',
    'كمية المنتج في المواد': 1.07},
    {
    'talabat product': 'خضار.كوم خوخ محلي 1كجم',
    'data entry product': 'خوخ محلي',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': '',
    'كمية المنتج في المواد': ''},
    {
    'talabat product': 'خضار.كوم يوسفي كريستينا 1كجم',
    'data entry product': 'يوسفي كريستينا',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'يوسفى كريستينا 1ك معبأ',
    'كمية المنتج في المواد': 1.05},
    {
    'talabat product': 'شمام شهد 1ك معبأ',
    'data entry product': 'شمام شهد',
    'نسبة الفرزة': 0.0,
    'التعبئه': 0.0,
    'فرق وزن': 0.0,
    'نسبة مرتجع': 0.0,
    'منتج المواد': 'شمام شهد 1ك معبأ',
    'كمية المنتج في المواد': 1.15}]



    conn = st.connection("gsheets", type=GSheetsConnection)

    # Open the spreadsheet by its key
    spreadsheet_key = "1yp-1Zswwwf3ZjNWGrks0R3OFqZWlF30jwAwb-PZV1XM"




    def extract_description_and_price(file_path):
        """Extract Arabic description and price from a Word document."""
        doc = Document(file_path)
        extracted_data = []
        
        file_name = os.path.basename(file_path)
        price_column_index = 3 if "بريد فاست" in file_name else 2  # 4th column for بريد فاست, 3rd otherwise

        # Extract date from filename using regex
        date_match = re.search(r"\d{2}-\d{2}-\d{4}", file_name)
        invoice_date = pd.NaT
        if date_match:
            invoice_date = pd.to_datetime(date_match.group(), format="%d-%m-%Y")

        for table in doc.tables:  
            for row in table.rows:  
                cells = [cell.text.strip() for cell in row.cells]
                if len(cells) > price_column_index:  # Ensure enough columns exist
                    arabic_description = cells[1]  
                    price = cells[price_column_index]  
                    extracted_data.append({
                        "Product": arabic_description,
                        "Price": price,
                        "File": file_name,
                        "Invoice Date": invoice_date
                    })

        return pd.DataFrame(extracted_data)

    global df_invoices
    def process_zip(zip_file):
        """Extract files from ZIP and process .docx invoices."""
        global df_invoices  # Declare the global variable

        extracted_folder = tempfile.mkdtemp()
        with zipfile.ZipFile(zip_file, 'r') as z:
            z.extractall(extracted_folder)
        
        docx_files = [os.path.join(extracted_folder, f) for f in os.listdir(extracted_folder) if f.endswith(".docx") and "طلبات" in f]
        if not docx_files:
            df_invoices = pd.DataFrame()  # Assign an empty DataFrame to the global variable
            return df_invoices, "No matching .docx files found."
        
        df_invoices = pd.concat([extract_description_and_price(file) for file in docx_files], ignore_index=True)
        df_invoices.replace(to_replace={'': None}, inplace=True)
        df_invoices = df_invoices[df_invoices["Product"] != "Arabic Description"].reset_index(drop=True)
        df_invoices.dropna(inplace=True)
        df_invoices = df_invoices[pd.to_numeric(df_invoices["Price"], errors="coerce").notna()]
        
        return df_invoices, None

    st.title("Invoice Data Extractor")
    uploaded_file = st.file_uploader("Upload a ZIP file containing .docx invoices", type="zip")

    if uploaded_file:
        with st.spinner("Processing..."):
            df, error = process_zip(uploaded_file)
        dfAlex = conn.read(worksheet="Alexandria", usecols=list(range(10)), ttl=5).dropna(how="all")
        dfAlex["تاريخ الشراء"] = pd.to_datetime(dfAlex["تاريخ الشراء"], format="%d/%m/%Y", errors="coerce")
        dfAlex['تكلفة الوحدة'] = dfAlex['تكلفة الوحدة'].round(2)
        dfCairo = conn.read(worksheet="Cairo", usecols=list(range(10)), ttl=5).dropna(how="all")
        dfCairo["تاريخ الشراء"] = pd.to_datetime(dfCairo["تاريخ الشراء"], format="%d/%m/%Y", errors="coerce")
        dfCairo['تكلفة الوحدة'] = dfCairo['تكلفة الوحدة'].round(2)
        cairo_tz = pytz.timezone("Africa/Cairo")

        # Get today's date in Cairo time and remove timezone for comparison
        today_cairo = pd.Timestamp.now(tz=cairo_tz).normalize().tz_localize(None)

        # Calculate the last three days (including today)
        three_days_ago = today_cairo - pd.Timedelta(days=2)  # Today and the 2 previous days

        # Ensure "تاريخ الشراء" is in datetime format
        dfAlex["تاريخ الشراء"] = pd.to_datetime(dfAlex["تاريخ الشراء"])
        dfCairo["تاريخ الشراء"] = pd.to_datetime(dfCairo["تاريخ الشراء"])

        # Filter each DataFrame to include only data from today and the last 2 days
        dfAlex = dfAlex[dfAlex["تاريخ الشراء"].between(three_days_ago, today_cairo)]
        dfCairo = dfCairo[dfCairo["تاريخ الشراء"].between(three_days_ago, today_cairo)]

        # Concatenate the filtered data
        df = pd.concat([dfAlex, dfCairo])

        # Group by "اسم الصنف" and aggregate required columns
        df = df.groupby("اسم الصنف").agg(
            عدد_العبوات=("عدد العبوات", "sum"),
            وزن_العبوات=("وزن العبوات", "sum"),
            وزن_قائم=("وزن قائم", "sum"),
            وزن_صافي=("وزن صافي", "sum"),
            تكلفة_الوحدة=("تكلفة الوحدة", "max"),  # Keep the highest unit cost
            الاجمالى=("الاجمالى", "sum")
        )

        # Reset index for final result
        df.reset_index(inplace=True)
        df.columns = df.columns.str.replace("_", " ")
        # Merge df with dfAlex and dfCairo separately
        df_merged_alex = df.merge(dfAlex[['اسم الصنف', 'تكلفة الوحدة', 'تاريخ الشراء', 'مورد الشركة']], 
                                on=['اسم الصنف', 'تكلفة الوحدة'], how='left')

        df_merged_cairo = df.merge(dfCairo[['اسم الصنف', 'تكلفة الوحدة', 'تاريخ الشراء', 'مورد الشركة']], 
                                on=['اسم الصنف', 'تكلفة الوحدة'], how='left')

        # Concatenate both merged DataFrames
        df = pd.concat([df_merged_alex, df_merged_cairo])

        # Drop duplicates to keep unique rows
        df.dropna(inplace=True)

        # Reset index
        df.reset_index(drop=True, inplace=True)
        # Convert single_item_products to a mapping
        single_map = {item["talabat product"]: item["data entry product"] for item in single_item_products}

        # Create df_single_mapping to merge
        df_single_mapping = pd.DataFrame(list(single_map.items()), columns=["Product", "اسم الصنف"])

        # Drop duplicates from df_invoices
        df_invoices = df_invoices.drop_duplicates(subset=["Product"])

        # First merge with single item products
        df_merged = df_invoices.merge(df_single_mapping, on="Product", how="left")

        # Find unmatched products (i.e., those that didn't match single_item_products)
        unmatched_products = df_merged[df_merged["اسم الصنف"].isna()]["Product"].tolist()

        # Multi-item mapping: Convert talabat product -> data entry product list
        multi_map = {item["talabat product"]: item["data entry product"] for item in multi_items_products}

        # Process multi-item products
        for product in unmatched_products:
            if product in multi_map:
                entry_products = multi_map[product]  
                if isinstance(entry_products, str):  
                    entry_products = [entry_products]  # Convert to list if it's a string

                # Assign the list directly in df_merged
                df_merged.loc[df_merged["Product"] == product, "اسم الصنف"] = str(entry_products)  # Store as string to avoid errors

        # Convert lists stored as strings back to actual lists
        df_merged["اسم الصنف"] = df_merged["اسم الصنف"].apply(lambda x: eval(x) if isinstance(x, str) and x.startswith("[") else x)

        # Reset index
        df_merged.reset_index(drop=True, inplace=True)
        df_merged.drop("File", axis=1, inplace=True)
        df_merged = df_merged.rename(columns={'Product': 'اسم الصنف طلبات', 'Price': 'اخر سعر بيع لطلبات', 'Invoice Date' : 'تاريخ فاتورة طلبات'})
        df_merged.dropna(subset=['اسم الصنف'], inplace=True)
        # Function to calculate تكلفة الوحدة for single-item products
        def get_single_item_cost(row, cost_df, single_map):
            talabat_product = row["اسم الصنف طلبات"]
            internal_product = row["اسم الصنف"]
            
            if pd.isna(internal_product):  # Skip if no matching product
                return None

            # Get the unit cost from df
            unit_cost = cost_df.loc[cost_df["اسم الصنف"] == internal_product, "تكلفة الوحدة"]
            
            if unit_cost.empty:
                return None  # Skip if no matching cost found
            
            unit_cost = unit_cost.values[0]  # Extract the cost
            
            # Get كمية المنتج في المواد from single_item_products
            quantity = next((item["كمية المنتج في المواد"] for item in single_item_products if item["talabat product"] == talabat_product), None)
            
            if quantity is None:
                return None  # Skip if no quantity found

            return unit_cost * quantity  # Calculate total cost

        # Function to calculate تكلفة الوحدة for multi-item products
        def get_multi_item_cost(row, cost_df, multi_map):
            talabat_product = row["اسم الصنف طلبات"]
            internal_products = row["اسم الصنف"]

            if not isinstance(internal_products, list):  # Skip if not a list
                return None
            
            # Get the corresponding مواد details from multi_items_products
            matching_item = next((item for item in multi_items_products if item["talabat product"] == talabat_product), None)
            
            if matching_item is None:
                return None
            
            total_cost = 0

            for i, product in enumerate(matching_item["منتج المواد"]):  
                # Get unit cost from df
                unit_cost = cost_df.loc[cost_df["اسم الصنف"] == product, "تكلفة الوحدة"]
                
                if unit_cost.empty:
                    continue  # Skip if no cost found
                
                unit_cost = unit_cost.values[0]  # Extract the cost
                quantity = matching_item["كمية المنتج في المواد"][i]  # Get quantity

                total_cost += unit_cost * quantity  # Add to total cost

            return total_cost

        # Apply the functions to df_merged
        df_merged["تكلفة الوحدة"] = df_merged.apply(
            lambda row: get_single_item_cost(row, df, single_map) if isinstance(row["اسم الصنف"], str) 
            else get_multi_item_cost(row, df, multi_map), axis=1
        )
        def get_single_item_attributes(row, single_map):
            talabat_product = row["اسم الصنف طلبات"]
            
            # Search for the product in single_item_products
            item = next((item for item in single_item_products if item["talabat product"] == talabat_product), None)
            
            if item is None:
                return None, None, None, None  # Return None for all attributes if not found

            return item.get("نسبة الفرزة"), item.get("التعبئه"), item.get("فرق وزن"), item.get("نسبة مرتجع")

        # Function to get additional attributes from multi-item products
        def get_multi_item_attributes(row, multi_map):
            talabat_product = row["اسم الصنف طلبات"]

            # Search for the product in multi_items_products
            item = next((item for item in multi_items_products if item["talabat product"] == talabat_product), None)
            
            if item is None:
                return None, None, None, None  # Return None for all attributes if not found

            return item.get("نسبة الفرزة"), item.get("التعبئه"), item.get("فرق وزن"), item.get("نسبة مرتجع")

        # Apply the functions to extract the new columns
        df_merged[["نسبة الفرزة", "التعبئه", "فرق وزن", "نسبة مرتجع"]] = df_merged.apply(
            lambda row: get_single_item_attributes(row, single_map) if isinstance(row["اسم الصنف"], str) 
            else get_multi_item_attributes(row, multi_map), axis=1, result_type="expand"
        )
        df_merged["اخر سعر بيع لطلبات"] = pd.to_numeric(df_merged["اخر سعر بيع لطلبات"], errors="coerce")
        df_merged["نسبة الفرزة"] = pd.to_numeric(df_merged["نسبة الفرزة"], errors="coerce")
        df_merged["تكلفة الوحدة"] = pd.to_numeric(df_merged["تكلفة الوحدة"], errors="coerce")
        df_merged["فرق وزن"] = pd.to_numeric(df_merged["فرق وزن"], errors="coerce")
        df_merged["التعبئه"] = pd.to_numeric(df_merged["التعبئه"], errors="coerce")
        df_merged["نسبة مرتجع"] = pd.to_numeric(df_merged["نسبة مرتجع"], errors="coerce")
        df_merged.dropna(inplace=True)
        df_merged.reset_index(drop=True, inplace=True)
        df_merged["اخر سعر بيع لطلبات"] = df_merged["اخر سعر بيع لطلبات"] * 0.9
        df_merged["قيمه الفرزه"] = df_merged["نسبة الفرزة"] * df_merged["تكلفة الوحدة"]
        df_merged["قيمه فرق الوزن"] = df_merged["فرق وزن"] * df_merged["تكلفة الوحدة"]
        df_merged["اجمالي التكلفة"] = df_merged["قيمه الفرزه"] + df_merged["قيمه فرق الوزن"] + df_merged["التعبئه"] + df_merged["تكلفة الوحدة"]
        df_merged["قيمه المرتجع"] = df_merged["نسبة مرتجع"] * df_merged["تكلفة الوحدة"]
        df_merged["سعر البيع قبل العمولة"] = df_merged["اجمالي التكلفة"] + df_merged["قيمه المرتجع"]
        df_merged["هامش ربح طلبات 8%"] = df_merged["سعر البيع قبل العمولة"] * 0.08
        df_merged["سعر البيع بهامش ربح 10%"] = (df_merged["هامش ربح طلبات 8%"] + df_merged["سعر البيع قبل العمولة"]) *1.1
        

        # Generate today's date
        todays_date = datetime.datetime.today().strftime('%Y-%m-%d')

        # Define filename
        filename = f"khodar_dashboard_{todays_date}.xlsx"

        # Save df_merged to an Excel file
        df_merged.to_excel(filename, index=False)

        # Email configuration
        sender_email = "salehgamalgad@gmail.com"  # Replace with your email
        receiver_email = "salehgamalgad@gmail.com"
        password = "xcli ahkk rulq btbs"  # Use an app password if using Gmail

        # Create the email
        msg = EmailMessage()
        msg["Subject"] = "Khodar Dashboard Report"
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg.set_content("Attached is the latest Khodar dashboard report.")

        # Attach the file
        with open(filename, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=filename)

        # Send the email
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(sender_email, password)
                server.send_message(msg)
            print("Email sent successfully!")
        except Exception as e:
            print(f"Error sending email: {e}")
        try:
            dashboard_sheet = conn.read(worksheet="Dashboard", usecols=list(range(25)), ttl=5).dropna(how="all")
        except gspread.exceptions.WorksheetNotFound:
            dashboard_sheet = conn.read(worksheet="CaDashboardiro", usecols=list(range(25)), ttl=5).dropna(how="all")

        # Retrieve existing data from the dashboard sheet
        existing_data = dashboard_sheet
        old_df = existing_data

        # Ensure correct data types for merging
        df_merged["تاريخ فاتورة طلبات"] = df_merged["تاريخ فاتورة طلبات"].astype(str)

        # Flatten lists in "اسم الصنف" column and keep numbers as numbers
        df_merged["اسم الصنف"] = df_merged["اسم الصنف"].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)

        # Concatenate new data with existing data (new data on top)
        if not old_df.empty:
            new_df = pd.concat([df_merged, old_df], ignore_index=True)
        else:
            new_df = df_merged

        # Convert DataFrame to list of lists (header + values)
        data_to_upload = [new_df.columns.tolist()] + new_df.values.tolist()

        # Upload to Google Sheets
        conn.update(worksheet="Dashboard", data=new_df)

        print("Data appended successfully to the 'Dashboard' sheet.")
        if error:
            st.error(error)
        elif not df.empty:
            st.success("Run successful!")
            st.dataframe(df_merged)
            
            csv = df.to_csv(index=False).encode("utf-8")
            st.download_button("Download CSV", csv, "invoices.csv", "text/csv")
        else:
            st.warning("No valid data extracted.")
