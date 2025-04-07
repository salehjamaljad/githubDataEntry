import streamlit as st
import pandas as pd

def pivot_app():
    st.title("Pivot 216 CSV File")

    # Upload the CSV file
    file_216 = st.file_uploader("Upload the 216.csv", type="csv")

    if file_216 is not None:
        # Read the uploaded CSV
        df = pd.read_csv(file_216)

        # Show raw data (optional)
        st.subheader("Raw Data")
        st.dataframe(df)

        # Pivot the table
        pivot_df = df.pivot_table(
            index='Product',
            columns='Store_name',
            values='Effective quantity',
            aggfunc='sum',
            fill_value=0
        )
        pivot_df["Total"] = pivot_df.sum(axis=1)

        # Show pivoted data
        st.subheader("Pivoted Data")
        st.dataframe(pivot_df)
    else:
        st.info("Please upload the 216.csv file.")
if __name__ == "__main__":
    pivot_app()