import streamlit as st
import pandas as pd
from io import BytesIO
import excel_file
st.header('Upload Budget Comp')
uploaded_file = st.file_uploader("Choose a file")
dataframe_budget_comp = pd.read_excel(uploaded_file)
output = BytesIO()
excel_file.create_excel(dataframe_budget_comp, output)
st.download_button(
        label="Download Excel workbook",
        data=output.getvalue(),
        file_name='Budget_Comp.xlsx',
        mime="application/vnd.ms-excel"
)