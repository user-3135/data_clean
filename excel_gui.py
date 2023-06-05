import streamlit as st
import pandas as pd
from io import BytesIO
import excel_file
st.header('Upload Budget Comp')
uploaded_file = st.file_uploader("Choose a file")
dataframe_budget_comp = pd.read_excel(uploaded_file)
output = BytesIO()
prop_name = excel_file.create_excel(dataframe_budget_comp, output)
name = prop_name + 'Budget Comp.xlsx'
st.download_button(
        label="Download Excel workbook",
        data=output.getvalue(),
        file_name=name,
        mime="application/vnd.ms-excel"
)
