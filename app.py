import streamlit as st
from pricingDataEntry import pricing_app
from stockKeepingDataEntry import stock_app
from dashboardApp import dashboardApp
from pivotApp import pivot_app
from pdfsToExcels import pdfToExcel
from breadfastInvoices import breadfastInvoices
from rabbitInvoices import rabbitInvoices
from goodsmartInvoices import goodsmartInvoices
# Demo user credentials
users = {
    "khodar1": {"password": "pricing", "access": "pricing"},
    "khodar2": {"password": "stock", "access": "stock"},
    "khodar3": {"password": "dashboard", "access": "dashboard"},
    "khodar4": {"password": "pivot", "access": "pivot"},
    "khodar5": {"password": "pdfToExcel", "access": "pdfToExcel"},
    "khodar6": {"password": "breadfastInvoices", "access": "breadfastInvoices"},
    "khodar8": {"password": "rabbitInvoices", "access": "rabbitInvoices"},
    "khodar9": {"password": "GoodsMartInvoices", "access": "goodsmartInvoices"}
}

def main():
    st.title("Login Form")
    
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    
    if st.button("Login"):
        if username in users and users[username]["password"] == password:
            st.session_state["logged_in"] = True
            st.session_state["access"] = users[username]["access"]
            st.success("Login successful!")
            st.rerun()
        else:
            st.error("Invalid username or password.")

if "logged_in" not in st.session_state or not st.session_state["logged_in"]:
    main()
else:
    if st.session_state["access"] == "pricing":
        pricing_app()
    elif st.session_state["access"] == "stock":
        stock_app()
    elif st.session_state["access"] == "dashboard":
        dashboardApp()
    elif st.session_state["access"] == "pivot":
        pivot_app()
    elif st.session_state["access"] == "pdfToExcel":
        pdfToExcel()
    elif st.session_state["access"] == "breadfastInvoices":
        breadfastInvoices()
    elif st.session_state["access"] == "rabbitInvoices":
        rabbitInvoices()
    elif st.session_state["access"] == "goodsmartInvoices":
        goodsmartInvoices()
