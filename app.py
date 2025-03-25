import streamlit as st
from pricingDataEntry import pricing_app
from stockKeepingDataEntry import stock_app

# Demo user credentials
users = {
    "pricing_user": {"password": "pricing123", "access": "pricing"},
    "stock_user": {"password": "stock123", "access": "stock"},
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
