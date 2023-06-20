import streamlit as st
import pandas as pd
from io import BytesIO
import excel_file
data = 1
st.header('import file')
uploaded_file_2 = st.file_uploader("Upload a Single File")

# data read
try:
    tb_df = pd.read_excel(uploaded_file_2, sheet_name = 'TB') #BS
except:
    pass
try:
    cash_flow_df = pd.read_excel(uploaded_file_2, sheet_name = 'CashFlow') #BS
except:
    st.write('Cash Flow Error')
try:
    bs_df = pd.read_excel(uploaded_file_2, sheet_name = 'BS') #BS
except:
    st.write('Balance Sheet Error')
try:
    is_df = pd.read_excel(uploaded_file_2, sheet_name = 'IS') #BS
except:
    st.write('Income Statement Error')
try:
    actual_budget = pd.read_excel(uploaded_file_2, sheet_name = 'Actual-Budget')
except:
    st.write('Actual - Budget Error')
try:
    data_ar_detail = pd.read_excel(uploaded_file_2, sheet_name = 'AR Detail')
except:
    st.write('AR Detail Error')
try:
    data_12_month = pd.read_excel(uploaded_file_2, sheet_name = 'IS 12 Month Actual')
except:
    st.write('IS 12 Month Actual')
try:    
    data_ten_sched = pd.read_excel(uploaded_file_2, sheet_name = 'TenSched1')
except:
    st.write('Ten Sched Error')
# -------------------------------------
output_2 = BytesIO()
excel_file_2.create_excel(actual_budget, 'test_clean.xlsx', is_df, bs_df, cash_flow_df, tb_df, data_ar_detail, data_12_month, data_ten_sched)
name = 'Property Workbook Test.xlsx'
st.download_button(
                    label="Download Excel workbook",
                    data=output_2.getvalue(),
                    file_name=name,
                    mime="application/vnd.ms-excel"
    )
#except:
#        pass
if data == 2:
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
                excel_file.create_excel_2(dataframe_budget_comp, output, dataframe_is, dataframe_bs, dataframe_cf, dataframe_tb, dataframe_pr)
                name = 'Property Workbook Test.xlsx'
                st.download_button(
                        label="Download Excel workbook",
                        data=output.getvalue(),
                        file_name=name,
                        mime="application/vnd.ms-excel"
                )
        except:
                pass
