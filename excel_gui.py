import streamlit as st
import pandas as pd
from io import BytesIO
import excel_file
import excel_file_2
import excel_file_3
import excel_file_4 as writer
import comb
import xlsxwriter as xl
import openpyxl
data = 1
st.header('import file')
uploaded_file_2 = st.file_uploader("Upload a Single File")
# data read
try:
    tb_df_1 = pd.read_excel(uploaded_file_2, sheet_name = 'TB') #BS
except:
    st.write('Trial Balance Error')
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
# ------------------------------------
try:
    data_payment_register = pd.read_excel(uploaded_file_2, sheet_name = 'Pymnt Register')
except:
    st.write('Pymnt Register Error')
try:
    data_ap_detail = pd.read_excel(uploaded_file_2, sheet_name = 'AP Detail')
except:
    st.write('AP Detail Error')
try:
    data_mth_gl = pd.read_excel(uploaded_file_2, sheet_name = 'MTH GL')
except:
    st.write('MTH GL Error')
try:    
    je_register_data = pd.read_excel(uploaded_file_2, sheet_name = 'JE Register')
except:
    st.write('JE Register Error')
output_4 = BytesIO()
try:
    workbook_not_func = xl.Workbook(output_4)
    #----------------------------------------------------------------------------------- 1
    act_bud_worksheet_func = workbook_not_func.add_worksheet('test')
    Income_Statement_wb = workbook_not_func.add_worksheet('Income Statement')
    worksheet_12_mo_actual = workbook_not_func.add_worksheet('IS 12 Month Actual')
    bs_wb = workbook_not_func.add_worksheet('BS')
    cf_worksheet_func = workbook_not_func.add_worksheet('cash flow')
    tb_worksheet_func = workbook_not_func.add_worksheet('trial balance')
    worksheet_tenancy_sched = workbook_not_func.add_worksheet('TenSched1')
    ap_detail_sheet = workbook_not_func.add_worksheet('AP Detail')
    worksheet_pay_reg = workbook_not_func.add_worksheet('Payment Register')
    aging_detail_sheet = workbook_not_func.add_worksheet('AR Detail')
    je_register_worksheet = workbook_not_func.add_worksheet('JE Register')
    mnth_gl_worksheet = workbook_not_func.add_worksheet('Mnth GL')
    #----------------------------------------------------------------------------------- 2
    # procs
    try:
        h = writer.JE_REGISTER_SHEET(workbook_not_func, je_register_data, je_register_worksheet)
    except:
        print(1)
    try:
        f = writer.mnth_gl_sheet(workbook_not_func, data_mth_gl, mnth_gl_worksheet) ##mnth_gl_sheet
    except:
        print(2)
    try:
        e = writer.ap_detail_sheet_def(workbook_not_func, data_ap_detail, ap_detail_sheet)
    except:
        print(3)
    try:
        d = writer.payment_register_sheet(workbook_not_func, data_payment_register, worksheet_pay_reg)
    except:
        print(4)
    #----------------------------------------------------------------------------------- 2
    if 1 == 1:
        # procs
        try:
            a = writer.twelve_month_actual_budget(workbook_not_func, data_12_month, worksheet_12_mo_actual)
        except:
            print(5)
        try:
            b = writer.ten_sched_1(workbook_not_func, data_ten_sched, worksheet_tenancy_sched)
        except:
            print(6)
        try:
            c = writer.aging_detail(workbook_not_func, data_ar_detail, aging_detail_sheet)
        except:
            print(7)
    #----------------------------------------------------------------------------------- 3
    try:
        aa = writer.create_xl_cf(workbook_not_func, cash_flow_df, cf_worksheet_func)
    except:
        print(8)
    try:
        ba = writer.create_xl_tb(workbook_not_func, tb_df_1, tb_worksheet_func)
    except:
        print(9)
    try:
        ca = writer.create_xl_balance_sheet(workbook_not_func, bs_df, bs_wb)
    except:
        print(10)
    try:
        fa = writer.budget_comp_sheet_creation(workbook_not_func, actual_budget, act_bud_worksheet_func)
    except:
        print(11)
    try:
        writer.income_statement(workbook_not_func, is_df, Income_Statement_wb)
    except:
        print(12)
    column_width_list = [
                [11.5, 0, 0, mnth_gl_worksheet]
                ,[16, 1, 1, mnth_gl_worksheet]
                ,[10.8, 2, 2, mnth_gl_worksheet]
                ,[11.3, 3, 3, mnth_gl_worksheet]
                ,[39.6, 4, 4, mnth_gl_worksheet]
                ,[9.3, 5, 5, mnth_gl_worksheet]
                ,[23.3, 6, 6, mnth_gl_worksheet]
                ,[11.4, 7, 8, mnth_gl_worksheet]
                ,[14.4, 9, 9, mnth_gl_worksheet]
                ,[50.7, 10, 10, mnth_gl_worksheet]
                ,[71, 0, 0, Income_Statement_wb]
                ,[17.7, 1, 2, Income_Statement_wb]
                ,[42.3,0,0, worksheet_12_mo_actual]
                ,[11.9,1,16, worksheet_12_mo_actual]
                ,[58.8, 0, 0, bs_wb]
                ,[17.9, 1, 3, bs_wb]
                ,[71, 0, 0, cf_worksheet_func ] ## cashflow
                ,[17.7, 1, 2, cf_worksheet_func] ## cashflow
                ,[38.7, 0, 0, tb_worksheet_func] ## Trial Balance
                ,[16.4, 1, 4, tb_worksheet_func] ## Trial Balance
                ,[12.7, 0, 16, worksheet_tenancy_sched] ## Tenancy Schedule
                ,[10, 0, 0, aging_detail_sheet] ## AR Detail
                ,[9.8, 1, 1, aging_detail_sheet] ## AR Detail
                ,[42.2, 2, 2, aging_detail_sheet] ## AR Detail
                ,[9.8, 3, 3, aging_detail_sheet] ## AR Detail
                ,[7.8, 4, 4, aging_detail_sheet] ## AR Detail
                ,[8.8, 5, 5, aging_detail_sheet] ## AR Detail
                ,[10.3, 6, 6, aging_detail_sheet] ## AR Detail
                ,[10.3, 7, 7, aging_detail_sheet] ## AR Detail
                ,[12, 8, 14, aging_detail_sheet] ## AR Detail
                ,[14.4, 0, 0, worksheet_pay_reg] ## Payment Register
                ,[6.4, 1, 1, worksheet_pay_reg] ## Payment Register
                ,[9.5, 2, 2, worksheet_pay_reg] ## Payment Register
                ,[10.7, 3, 3, worksheet_pay_reg] ## Payment Register
                ,[22.9, 4, 4, worksheet_pay_reg] ## Payment Register
                ,[10.2, 5, 5, worksheet_pay_reg] ## Payment Register
                ,[9.5, 6, 6, worksheet_pay_reg] ## Payment Register
                ,[8, 7, 7, worksheet_pay_reg] ## Payment Register
                ,[7.5, 8, 8, worksheet_pay_reg] ## Payment Register
                ,[8.2, 9, 9, worksheet_pay_reg] ## Payment Register
                ,[10, 10, 10, worksheet_pay_reg ] ## Payment Register
                ,[16, 11, 11, worksheet_pay_reg] ## Payment Register
                ,[10, 12, 12, worksheet_pay_reg] ## Payment Register
                ,[47.3, 13, 13, worksheet_pay_reg] ## Payment Register
                ,[13.8, 0, 0, ap_detail_sheet] ## AP Detail
                ,[12.2, 1, 0, ap_detail_sheet] ## AP Detail
                ,[25.2, 2, 0, ap_detail_sheet] ## AP Detail
                ,[6.4, 3, 0, ap_detail_sheet] ## AP Detail
                ,[7, 4, 0, ap_detail_sheet] ## AP Detail
                ,[8.2, 5, 0, ap_detail_sheet] ## AP Detail
                ,[10.8, 6, 0, ap_detail_sheet] ## AP Detail
                ,[16.3, 7, 0, ap_detail_sheet] ## AP Detail
                ,[17.3, 8, 0, ap_detail_sheet] ## AP Detail
                ,[10.3, 9, 14, ap_detail_sheet] ## AP Detail
                ,[27.4, 15, 15, ap_detail_sheet] ## AP Detail
                ,[9.8, 0, 0, je_register_worksheet] ## JE Register
                ,[10.2, 1, 3, je_register_worksheet] ## JE Register
                ,[9.5, 4, 4, je_register_worksheet] ## JE Register
                ,[12.8, 5, 5, je_register_worksheet] ## JE Register
                ,[42.2, 6, 6, je_register_worksheet] ## JE Register
                ,[16, 7, 7, je_register_worksheet] ## JE Register
                ,[10, 8, 8, je_register_worksheet] ## JE Register
                ,[12.5, 9, 10, je_register_worksheet] ## JE Register
                ,[6.3, 12, 12, je_register_worksheet] ## JE Register
                ,[37.5, 13, 13, je_register_worksheet] ## JE Register
                # ,[11.5, 0, 0, ] ## General Ledger
                # ,[16, 0, 0, ] ## General Ledger
                # ,[10.8, 0, 0, ] ## General Ledger
                # ,[11.3, 0, 0, ] ## General Ledger
                # ,[39.7, 0, 0, ] ## General Ledger
                # ,[9.3, 0, 0, ] ## General Ledger
                # ,[23.3, 0, 0, ] ## General Ledger
                # ,[11.4, 0, 0, ] ## General Ledger
                # ,[14.4, 0, 0, ] ## General Ledger
                # ,[50.7, 0, 0, ] ## General Ledger
    
        ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass
    
    workbook_not_func.close()
except: 
    st.write('error in base')
try:
    wb_base = openpyxl.load_workbook(output_4)
    vals_to_rip = ['Summary'
                   , 'FS Checklist'
                   , 'Financial Summary Chart'
                   , 'Rent Roll'
                   , 'Mngt Fee '
                   , 'RET & Insurance'
                   , 'Depr Schedule'
                   , 'Amort Schedule'
                   , 'Mortgage Statement'
                   , 'Bank Recon'
                  ]
    wb_add = openpyxl.load_workbook(uploaded_file_2)
    final_output = BytesIO()
    for i in vals_to_rip:    
        target_sheet = wb_base.create_sheet(i)
        try:
            source_sheet = wb_add[i]
            comb.copy_sheet(source_sheet, target_sheet)
        except:
            print(i)
        
    wb_base.save(final_output)
except:
    pass
# -------------------------------------
try:
    try:
        name = 'Combined Property Workbook Test.xlsx'
        st.download_button(
                            label="Download Excel Combined Workbook",
                            data=final_output.getvalue(),
                            file_name=name,
                            mime="application/vnd.ms-excel"
        )
    except:
        try:
            st.download_button(
                                label="Download Excel Workbook",
                                data=output_4.getvalue(),
                                file_name=name,
                                mime="application/vnd.ms-excel"
            )
        except:
            output_2 = BytesIO()
            excel_file_3.create_excel_v3(actual_budget, output_2, is_df, bs_df, cash_flow_df, tb_df_1 , data_ar_detail, data_12_month, data_ten_sched
                                        , je_register_data, data_mth_gl,data_ap_detail, data_payment_register)
            name = 'Property Workbook Test.xlsx'
            st.download_button(
                                label="Download Excel workbook",
                                data=output_2.getvalue(),
                                file_name=name,
                                mime="application/vnd.ms-excel"
            )
except:
        pass
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
