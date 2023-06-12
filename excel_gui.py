import streamlit as st
import pandas as pd
from io import BytesIO
import excel_file
st.header('Upload Budget Comp')
uploaded_file = st.file_uploader("Choose a file")
# income_statement_1
st.header('Upload Income Statement')
uploaded_file_income_statemnt_1 = st.file_uploader("Upload the Income Statement")
# bal_sheet_1
st.header('Upload Balance Sheet Period Change')
uploaded_file_balance_sheet = st.file_uploader("Upload the Balance Sheet Period Change")
# cash_flow_1_df
st.header('Upload Statement of Cashflow')
uploaded_file_cash_flow = st.file_uploader("Upload the Statement of Cash Flow")
# trail_balance_df
st.header('Upload Trial Balance')
uploaded_file_trial_balance = st.file_uploader("Upload the Trial Balance")
# payment_register_df
st.header('Upload Payment Register')
uploaded_file_payment_register = st.file_uploader("Upload teh Payment Register")
dataframe_budget_comp = pd.read_excel(uploaded_file)
dataframe_is = pd.read_excel(uploaded_file_income_statemnt_1)
dataframe_bs = pd.read_excel(uploaded_file_balance_sheet)
dataframe_cf = pd.read_excel(uploaded_file_cash_flow)
dataframe_tb = pd.read_excel(uploaded_file_trial_balance)
dataframe_pr = pd.read_excel(uploaded_file_payment_register)
try:
        output = BytesIO()
        prop_name = excel_file.create_excel_2(dataframe_budget_comp, output, dataframe_is, dataframe_bs, dataframe_cf, dataframe_tb, dataframe_pr)
        name = prop_name + 'Budget Comp.xlsx'
        st.download_button(
                label="Download Excel workbook",
                data=output.getvalue(),
                file_name=name,
                mime="application/vnd.ms-excel"
        )
except:
        pass
