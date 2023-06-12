import pandas as pd
import xlsxwriter as xl
import math
def create_excel(df, xlfile):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    prop_name = df.columns[0]
    try:
        prop_name = prop_name.split('(', 1)[0]
    except:
        pass
    df = df.rename(columns={df.columns[0]: 'Col1'
                           , df.columns[1]: 'Col2'
                           , df.columns[2]: 'Col3'
                           , df.columns[3]: 'Col4'
                           , df.columns[4]: 'Col5'
                           , df.columns[5]: 'Col6'
                           , df.columns[6]: 'Col7'
                           , df.columns[7]: 'Col8'
                           , df.columns[8]: 'Col9'
                           })
    header_2 = df['Col1'][0]
    header_3 = df['Col1'][1]
    header_4 = df['Col1'][2]
    header_cols = ['MTD Actual'
                   ,'MTD Budget'
                   ,'Variance'
                   ,'% Var'
                   ,'YTD Actual'
                   ,'YTD Budget'
                   ,'Variance'
                   ,'% Var'
                   , 'Variance Explanations (5% and $2,500)']
    # clean data
    df=df.dropna(subset=['Col1']).reset_index(drop=True)
    def clean_text_col_1(text_val):
        return_text_val = ''
        index_val = 0
        for i in text_val:
            if index_val > 2:
                return_text_val = return_text_val + str(i)
            index_val += 1
        return return_text_val
    df['Col1'] = df.apply(lambda x: clean_text_col_1(x['Col1']), axis=1)
    def flag_zero_vals(val_1, val_2):
        flag_zero_val = 0
        if(math.isnan(val_1)):
            if(math.isnan(val_2)):
                flag_zero_val = 1
        return flag_zero_val
    df['Nan_Var_Check'] = df.apply(lambda x: flag_zero_vals(x['Col5'], x['Col9']), axis=1)
    def flag_total_rows(val_1):
        flag_total_val = 0
        if 'TOTAL' in val_1:
            flag_total_val = 1
        return flag_total_val
    df['Total_Check'] = df.apply(lambda x: flag_total_rows(x['Col1']), axis=1)
    def flag_header_rows(val_1):
        flag_total_val = 0
        if val_1[0] != ' ':
            flag_total_val = 1
        return flag_total_val
    df['Header_Check'] = df.apply(lambda x: flag_header_rows(x['Col1']), axis=1)
    # wirte excel
    workbook = xl.Workbook(xlfile)
    worksheet = workbook.add_worksheet('Budget Comparison')
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A1:I1", prop_name, header_format_1)
    worksheet.merge_range("A2:I2", header_2, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A3:I3", header_3, header_format_2)
    worksheet.merge_range("A4:I4", header_4, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    for row in range(10):
        if row == 0:
            worksheet.write_blank(4, row, '', header_format_3)
        else:
            worksheet.write_string(4, row, header_cols[row - 1], header_format_3)
    worksheet.merge_range(5, 0, 5, 9, '', header_format_2)
    worksheet.set_row(5,7.5)
    row_write_val = 6
    row_val_format_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item = workbook.add_format({'font_color': dark_gray_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_num = workbook.add_format({'font_color': dark_gray_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':44
                                     })
    row_val_format_sub_item_percent = workbook.add_format({'font_color': dark_gray_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                     })
    row_val_format_total_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    #-------------------------------------------------------------------------
    for i in range(4, df.shape[0] - 1):
        if df['Header_Check'][i] == 1:
            try:
                next_header_val = df['Header_Check'][i+1]
                next_total_val = df['Total_Check'][i+1]
                next_nan_val = df['Nan_Var_Check'][i+1]
            except:
                next_header_val = 0
                next_total_val = 0
                next_nan_val = 0
            # try to get next vals for logic
            new_row_needed = 0
            if df['Col1'][i] in ['OPERATING INCOME', 'OPERATING EXPENSES', 'RECOVERABLE', 'NON-RECOVERABLE']:
                worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                worksheet.merge_range(row_write_val, 0, row_write_val, 9, '', header_format_2)
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                if df['Total_Check'][i] == 1:
                    if next_total_val == 1:
                        worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                        worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num)
                        worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num)
                        worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num)
                        try:
                            worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent)
                        except:
                            worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent)
                        worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num)
                        worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num)
                        worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num)
                        try:
                            worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_total_item_percent)
                        except:
                            worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent)
                        row_write_val = row_write_val + 1
                    else:
                        if df['Nan_Var_Check'][i] == 0:
                            worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                            worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num)
                            try:
                                worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent)
                            except:
                                worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent)
                            worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num)
                            try:
                                worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_total_item_percent)
                            except:
                                worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent)
                            row_write_val = row_write_val + 1
                            if next_header_val == 1:
                                new_row_needed = 1
                            else:
                                pass
                else:
                    if next_nan_val == 0:
                        worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                        row_write_val = row_write_val + 1
                    elif next_header_val == 0:
                        row_val_check = i
                        value_needed = 0
                        next_header_found = 0
                        while row_val_check <= df.shape[0] - 1 and next_header_found == 1:
                            if df['Header_Check'][row_val_check] == 1:
                                next_header_found = 1
                            elif df['Nan_Var_Check'][row_val_check] == 1:
                                value_needed = 1
                            else:
                                row_val_check = row_val_check + 1
                        if value_needed == 1:
                            worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                            row_write_val = row_write_val + 1
                    else:
                        pass
            # ------------------------------------------
            # add a row or not
            if new_row_needed == 1:
                worksheet.merge_range(row_write_val, 0, row_write_val, 9, '', header_format_2)
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                pass
        else:
            if df['Nan_Var_Check'][i] == 1:
                pass
            else:
                worksheet.write_string(row_write_val, 0, df['Col1'][i],row_val_format_sub_item)
                worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_sub_item_num)
                try:
                    worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_sub_item_percent)
                except:
                    worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent)
                worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_sub_item_num)
                try:
                    worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_sub_item_percent)
                except:
                    worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent)
                row_write_val = row_write_val + 1
    worksheet.set_column(0,0,49.29)
    worksheet.set_column(9,9,49.29)
    worksheet.set_column(1,8,15)
    worksheet.print_area(0,0, row_write_val - 1, 9)
    worksheet.fit_to_pages(1, 1)
    worksheet.set_landscape()
    workbook.close()
    return prop_name
def create_excel_2(df, xlfile, income_statement_1, bal_sheet_1, cash_flow_1_df, trail_balance_df, payment_register_df):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    payment_register_gray = '#F2F2F2'
    prop_name = df.columns[0]
    prop_name_is = income_statement_1.columns[0]
    try:
        prop_name = prop_name.split('(', 1)[0]
    except:
        pass
    try:
        prop_name_is = prop_name_is('(', 1)[0]
    except:
        pass
    df = df.rename(columns={df.columns[0]: 'Col1'
                           , df.columns[1]: 'Col2'
                           , df.columns[2]: 'Col3'
                           , df.columns[3]: 'Col4'
                           , df.columns[4]: 'Col5'
                           , df.columns[5]: 'Col6'
                           , df.columns[6]: 'Col7'
                           , df.columns[7]: 'Col8'
                           , df.columns[8]: 'Col9'
                           })
    # income statement 1
    income_statement_1 = income_statement_1.rename(columns={income_statement_1.columns[0]: 'Col1'
                           , income_statement_1.columns[1]: 'Col2'
                           , income_statement_1.columns[2]: 'Col3'
                           , income_statement_1.columns[3]: 'Col4'
                           , income_statement_1.columns[4]: 'Col5'
                           })
    header_2 = df['Col1'][0]
    header_3 = df['Col1'][1]
    header_4 = df['Col1'][2]
    header_cols = ['MTD Actual'
                   ,'MTD Budget'
                   ,'Variance'
                   ,'% Var'
                   ,'YTD Actual'
                   ,'YTD Budget'
                   ,'Variance'
                   ,'% Var'
                   , 'Variance Explanations (5% and $2,500)']
    header_cols_is = ['Month to Date'
                  ,'Year to Date']
    header_2_is = income_statement_1['Col1'][0]
    header_3_is = income_statement_1['Col1'][1]
    header_4_is = income_statement_1['Col1'][2]
    # clean data
    df=df.dropna(subset=['Col1']).reset_index(drop=True)
    income_statement_1=income_statement_1.dropna(subset=['Col1']).reset_index(drop=True)
    def clean_text_col_1(text_val):
        return_text_val = ''
        index_val = 0
        for i in text_val:
            if index_val > 2:
                return_text_val = return_text_val + str(i)
            index_val += 1
        return return_text_val
    df['Col1'] = df.apply(lambda x: clean_text_col_1(x['Col1']), axis=1)
    income_statement_1['Col1'] = income_statement_1.apply(lambda x: clean_text_col_1(x['Col1']), axis=1)
    def flag_zero_vals(val_1, val_2):
        flag_zero_val = 0
        if(math.isnan(val_1)):
            if(math.isnan(val_2)):
                flag_zero_val = 1
        return flag_zero_val
    def flag_zero_vals_2(val_1, val_2):
        flag_zero_val = 0
        if(val_1 == 0):
            if(val_2 == 0):
                flag_zero_val = 1
        return flag_zero_val
    def flag_zero_vals_3(val_1, val_2):
        flag_zero_val = 0
        if type(val_1) == str:
            if type(val_2) == str:
                flag_zero_val = 1
        elif(math.isnan(val_1)):
            if(math.isnan(val_2)):
                flag_zero_val = 1
        return flag_zero_val
    df['Nan_Var_Check'] = df.apply(lambda x: flag_zero_vals(x['Col5'], x['Col9']), axis=1)
    income_statement_1['Nan_Var_Check'] = income_statement_1.apply(lambda x: flag_zero_vals_2(x['Col2'], x['Col4']), axis=1)
    def flag_total_rows(val_1):
        flag_total_val = 0
        try:
            if 'TOTAL' in val_1:
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    df['Total_Check'] = df.apply(lambda x: flag_total_rows(x['Col1']), axis=1)
    income_statement_1['Total_Check'] = income_statement_1.apply(lambda x: flag_total_rows(x['Col1']), axis=1)
    def flag_header_rows(val_1):
        flag_total_val = 0
        try:
            if val_1[0] != ' ':
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    df['Header_Check'] = df.apply(lambda x: flag_header_rows(x['Col1']), axis=1)
    income_statement_1['Header_Check'] = income_statement_1.apply(lambda x: flag_header_rows(x['Col1']), axis=1)
    # balance sheet Period Change
    prop_name_bs_change = bal_sheet_1.columns[0]
    try:
        prop_name_bs_change = prop_name_bs_change('(', 1)[0]
    except:
        pass
    bal_sheet_1 = bal_sheet_1.rename(columns={bal_sheet_1.columns[0]: 'Col1'
                           , bal_sheet_1.columns[1]: 'Col2'
                           , bal_sheet_1.columns[2]: 'Col3'
                           , bal_sheet_1.columns[3]: 'Col4'
                        })
    header_cols_bs_change_1 = ['Balance'
                  ,'Beginning'
                  ,'Net']
    header_cols_bs_change_2 = ['Current Period'
                  ,'Balance'
                  ,'Change']
    header_2_bal_sheet_1 = bal_sheet_1['Col1'][0]
    header_3_bal_sheet_1 = bal_sheet_1['Col1'][1]
    header_4_bal_sheet_1 = bal_sheet_1['Col1'][2]
    bal_sheet_1=bal_sheet_1.dropna(subset=['Col1']).reset_index(drop=True)
    bal_sheet_1['Col1'] = bal_sheet_1.apply(lambda x: clean_text_col_1(x['Col1']), axis=1)
    bal_sheet_1['Nan_Var_Check'] = bal_sheet_1.apply(lambda x: flag_zero_vals_2(x['Col2'], x['Col3']), axis=1)
    bal_sheet_1['Total_Check'] = bal_sheet_1.apply(lambda x: flag_total_rows(x['Col1']), axis=1)
    bal_sheet_1['Header_Check'] = bal_sheet_1.apply(lambda x: flag_header_rows(x['Col1']), axis=1)
    # Cash Flow 1
    prop_name_cf_1 = cash_flow_1_df.columns[0]
    try:
        prop_name_cf_1 = prop_name_cf_1('(', 1)[0]
    except:
        pass
    cash_flow_1_df = cash_flow_1_df.rename(columns={cash_flow_1_df.columns[0]: 'Col1'
                           , cash_flow_1_df.columns[1]: 'Col2'
                           , cash_flow_1_df.columns[2]: 'Col3'
                           , cash_flow_1_df.columns[3]: 'Col4'
                           , cash_flow_1_df.columns[4]: 'Col5'
                        })
    header_cols_bs_change_1 = ['Balance'
                  ,'Beginning'
                  ,'Net']
    header_2_cf_1 = cash_flow_1_df['Col1'][0]
    header_3_cf_1 = cash_flow_1_df['Col1'][1]
    header_4_cf_1 = cash_flow_1_df['Col1'][2]
    cash_flow_1_df=cash_flow_1_df.dropna(subset=['Col1']).reset_index(drop=True)
    cf_df_1_extra_row_val = cash_flow_1_df[cash_flow_1_df['Col1']=='   TOTAL OF ALL'].index.values.astype(int)[0] + 1
    cash_flow_1_df_end_of_page = cash_flow_1_df.iloc[cf_df_1_extra_row_val:].reset_index(drop=True)
    cash_flow_1_df['Col1'] = cash_flow_1_df.apply(lambda x: clean_text_col_1(x['Col1']), axis=1)
    cash_flow_1_df['Nan_Var_Check'] = cash_flow_1_df.apply(lambda x: flag_zero_vals_2(x['Col2'], x['Col4']), axis=1)
    cash_flow_1_df['Total_Check'] = cash_flow_1_df.apply(lambda x: flag_total_rows(x['Col1']), axis=1)
    cash_flow_1_df['Header_Check'] = cash_flow_1_df.apply(lambda x: flag_header_rows(x['Col1']), axis=1)
    # wirte excel
    workbook = xl.Workbook(xlfile)
    worksheet = workbook.add_worksheet('Budget Comparison')
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A1:I1", prop_name, header_format_1)
    worksheet.merge_range("A2:I2", header_2, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A3:I3", header_3, header_format_2)
    worksheet.merge_range("A4:I4", header_4, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    for row in range(10):
        if row == 0:
            worksheet.write_blank(4, row, '', header_format_3)
        else:
            worksheet.write_string(4, row, header_cols[row - 1], header_format_3)
    worksheet.merge_range(5, 0, 5, 9, '', header_format_2)
    worksheet.set_row(5,7.5)
    row_write_val = 6
    row_val_format_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item = workbook.add_format({'font_color': dark_gray_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_2 = workbook.add_format({'font_color': dark_gray_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_num = workbook.add_format({'font_color': dark_gray_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':44
                                     })
    row_val_format_sub_item_percent = workbook.add_format({'font_color': dark_gray_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                     })
    row_val_format_total_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    #-------------------------------------------------------------------------
    for i in range(4, df.shape[0] - 1):
        if df['Header_Check'][i] == 1:
            try:
                next_header_val = df['Header_Check'][i+1]
                next_total_val = df['Total_Check'][i+1]
                next_nan_val = df['Nan_Var_Check'][i+1]
            except:
                next_header_val = 0
                next_total_val = 0
                next_nan_val = 0
            # try to get next vals for logic
            new_row_needed = 0
            if df['Col1'][i] in ['OPERATING INCOME', 'OPERATING EXPENSES', 'RECOVERABLE', 'NON-RECOVERABLE']:
                worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                worksheet.merge_range(row_write_val, 0, row_write_val, 9, '', header_format_2)
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                if df['Total_Check'][i] == 1:
                    if next_total_val == 1:
                        worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                        worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num)
                        worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num)
                        worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num)
                        try:
                            worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent)
                        except:
                            worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent)
                        worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num)
                        worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num)
                        worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num)
                        try:
                            worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_total_item_percent)
                        except:
                            worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent)
                        row_write_val = row_write_val + 1
                    else:
                        if df['Nan_Var_Check'][i] == 0:
                            worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                            worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num)
                            try:
                                worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent)
                            except:
                                worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent)
                            worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num)
                            try:
                                worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_total_item_percent)
                            except:
                                worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent)
                            row_write_val = row_write_val + 1
                            if next_header_val == 1:
                                new_row_needed = 1
                            else:
                                pass
                else:
                    if next_nan_val == 0:
                        worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                        row_write_val = row_write_val + 1
                    elif next_header_val == 0:
                        row_val_check = i
                        value_needed = 0
                        next_header_found = 0
                        while row_val_check <= df.shape[0] - 1 and next_header_found == 1:
                            if df['Header_Check'][row_val_check] == 1:
                                next_header_found = 1
                            elif df['Nan_Var_Check'][row_val_check] == 1:
                                value_needed = 1
                            else:
                                row_val_check = row_val_check + 1
                        if value_needed == 1:
                            worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                            row_write_val = row_write_val + 1
                    else:
                        pass
            # ------------------------------------------
            # add a row or not
            if new_row_needed == 1:
                worksheet.merge_range(row_write_val, 0, row_write_val, 9, '', header_format_2)
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                pass
        else:
            if df['Nan_Var_Check'][i] == 1:
                pass
            else:
                worksheet.write_string(row_write_val, 0, df['Col1'][i],row_val_format_sub_item)
                worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_sub_item_num)
                try:
                    worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_sub_item_percent)
                except:
                    worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent)
                worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_sub_item_num)
                try:
                    worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_sub_item_percent)
                except:
                    worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent)
                row_write_val = row_write_val + 1
    worksheet.set_column(0,0,49.29)
    worksheet.set_column(9,9,49.29)
    worksheet.set_column(1,8,15)
    worksheet.print_area(0,0, row_write_val - 1, 9)
    num_pages_budget_1 = math.ceil(row_write_val/65)
    worksheet.fit_to_pages(1, num_pages_budget_1)
    worksheet.set_landscape()
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    Income_Statement = workbook.add_worksheet('Income Statement')
    Income_Statement.merge_range("A1:C1", prop_name_is, header_format_1)
    Income_Statement.merge_range("A2:C2", header_2_is, header_format_1)
    Income_Statement.merge_range("A3:C3", header_3_is, header_format_2)
    Income_Statement.merge_range("A4:C4", header_4_is, header_format_2)
    for row in range(3):
        if row == 0:
            Income_Statement.write_blank(4, row, '', header_format_3)
            Income_Statement.write_blank(5, row, '', header_format_3)
        else:
            Income_Statement.write_string(4, row, header_cols_is[row - 1], header_format_3)
    Income_Statement.write_formula(5, 1, '=+TEXT(RIGHT(A3,8),"mmmm yyyy")', header_format_3)
    Income_Statement.write_formula(5, 2, '=+B6', header_format_3)
    Income_Statement.merge_range(6, 0, 6, 9, '', header_format_2)
    Income_Statement.set_row(6,7.5)
    row_write_val = 7
    for i in range(4, income_statement_1.shape[0] - 1):
        if income_statement_1['Header_Check'][i] == 1:
            try:
                next_header_val = income_statement_1['Header_Check'][i+1]
                next_total_val = income_statement_1['Total_Check'][i+1]
                next_nan_val = income_statement_1['Nan_Var_Check'][i+1]
            except:
                next_header_val = 0
                next_total_val = 0
                next_nan_val = 0
            # try to get next vals for logic
            new_row_needed = 0
            if income_statement_1['Col1'][i] in ['OPERATING INCOME', 'OPERATING EXPENSES', 'RECOVERABLE', 'NON-RECOVERABLE']:
                Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                Income_Statement.merge_range(row_write_val, 0, row_write_val, 9, '', header_format_2)
                Income_Statement.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                if income_statement_1['Total_Check'][i] == 1:
                    if next_total_val == 1:
                        Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header)
                        Income_Statement.write_number(row_write_val, 1, income_statement_1['Col2'][i],row_val_format_total_item_num)
                        Income_Statement.write_number(row_write_val, 2, income_statement_1['Col4'][i],row_val_format_total_item_num)
                        row_write_val = row_write_val + 1
                    else:
                        if income_statement_1['Nan_Var_Check'][i] == 0:
                            Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header)
                            Income_Statement.write_number(row_write_val, 1, income_statement_1['Col2'][i],row_val_format_total_item_num)
                            Income_Statement.write_number(row_write_val, 2, income_statement_1['Col4'][i],row_val_format_total_item_num)
                            row_write_val = row_write_val + 1
                            if next_header_val == 1:
                                new_row_needed = 1
                            else:
                                pass
                else:
                    if next_nan_val == 0:
                        Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header)
                        row_write_val = row_write_val + 1
                    elif next_header_val == 0:
                        row_val_check = i
                        value_needed = 0
                        next_header_found = 0
                        while row_val_check <= income_statement_1.shape[0] - 1 and next_header_found == 1:
                            if income_statement_1['Header_Check'][row_val_check] == 1:
                                next_header_found = 1
                            elif income_statement_1['Nan_Var_Check'][row_val_check] == 1:
                                value_needed = 1
                            else:
                                row_val_check = row_val_check + 1
                        if value_needed == 1:
                            Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header)
                            row_write_val = row_write_val + 1
                    else:
                        pass
            # ------------------------------------------
            # add a row or not
            if new_row_needed == 1:
                Income_Statement.merge_range(row_write_val, 0, row_write_val, 9, '', header_format_2)
                Income_Statement.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                pass
        else:
            if income_statement_1['Nan_Var_Check'][i] == 1:
                pass
            else:
                Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i],row_val_format_sub_item)
                Income_Statement.write_number(row_write_val, 1, income_statement_1['Col2'][i],row_val_format_sub_item_num)
                Income_Statement.write_number(row_write_val, 2, income_statement_1['Col4'][i],row_val_format_sub_item_num)
                row_write_val = row_write_val + 1
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    BS_Change_Sheet = workbook.add_worksheet('Balance Sheet Period Change')
    BS_Change_Sheet.merge_range("A1:D1", prop_name_is, header_format_1)
    BS_Change_Sheet.merge_range("A2:D2", header_2_is, header_format_1)
    BS_Change_Sheet.merge_range("A3:D3", header_3_is, header_format_2)
    BS_Change_Sheet.merge_range("A4:D4", header_4_is, header_format_2)
    for row in range(4):
        if row == 0:
            BS_Change_Sheet.write_blank(4, row, '', header_format_3)
            BS_Change_Sheet.write_blank(5, row, '', header_format_3)
            BS_Change_Sheet.write_blank(6, row, '', header_format_3)
        else:
            BS_Change_Sheet.write_string(4, row, header_cols_bs_change_1[row - 1], header_format_3)
            BS_Change_Sheet.write_string(5, row, header_cols_bs_change_2[row - 1], header_format_3)
    BS_Change_Sheet.write_formula(6, 1, '=+EOMONTH(RIGHT(A3,8),0)', header_format_3)
    BS_Change_Sheet.write_formula(6, 2, '=+EOMONTH(B7,-1)', header_format_3)
    BS_Change_Sheet.write_blank(6, 3, '', header_format_3)
    BS_Change_Sheet.merge_range(7, 0, 7, 4, '', header_format_2)
    BS_Change_Sheet.set_row(7,7.5)
    row_write_val = 8
    for i in range(4, bal_sheet_1.shape[0] - 1):
        if bal_sheet_1['Header_Check'][i] == 1:
            try:
                next_header_val = bal_sheet_1['Header_Check'][i+1]
                next_total_val = bal_sheet_1['Total_Check'][i+1]
                next_nan_val = bal_sheet_1['Nan_Var_Check'][i+1]
            except:
                next_header_val = 0
                next_total_val = 0
                next_nan_val = 0
            # try to get next vals for logic
            new_row_needed = 0
            if bal_sheet_1['Col1'][i] in ['OPERATING INCOME', 'OPERATING EXPENSES', 'RECOVERABLE', 'NON-RECOVERABLE']:
                BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                BS_Change_Sheet.merge_range(row_write_val, 0, row_write_val, 4, '', header_format_2)
                BS_Change_Sheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                if bal_sheet_1['Total_Check'][i] == 1:
                    if next_total_val == 1:
                        BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
                        BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i],row_val_format_total_item_num)
                        BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i],row_val_format_total_item_num)
                        BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i],row_val_format_total_item_num)
                        row_write_val = row_write_val + 1
                    else:
                        if bal_sheet_1['Nan_Var_Check'][i] == 0:
                            BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
                            BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i],row_val_format_total_item_num)
                            BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i],row_val_format_total_item_num)
                            BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i],row_val_format_total_item_num)
                            row_write_val = row_write_val + 1
                            if next_header_val == 1:
                                new_row_needed = 1
                            else:
                                pass
                else:
                    if next_nan_val == 0:
                        BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
                        row_write_val = row_write_val + 1
                    elif next_header_val == 0:
                        row_val_check = i
                        value_needed = 0
                        next_header_found = 0
                        while row_val_check <= bal_sheet_1.shape[0] - 1 and next_header_found == 1:
                            if bal_sheet_1['Header_Check'][row_val_check] == 1:
                                next_header_found = 1
                            elif bal_sheet_1['Nan_Var_Check'][row_val_check] == 1:
                                value_needed = 1
                            else:
                                row_val_check = row_val_check + 1
                        if value_needed == 1:
                            Income_Statement.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
                            row_write_val = row_write_val + 1
                    else:
                        pass
            # ------------------------------------------
            # add a row or not
            if new_row_needed == 1:
                BS_Change_Sheet.merge_range(row_write_val, 0, row_write_val, 4, '', header_format_2)
                BS_Change_Sheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                pass
        else:
            if bal_sheet_1['Nan_Var_Check'][i] == 1:
                pass
            else:
                BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i],row_val_format_sub_item)
                BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i],row_val_format_sub_item_num)
                BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i],row_val_format_sub_item_num)
                BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i],row_val_format_sub_item_num)
                row_write_val = row_write_val + 1
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    Cash_Flow_1 = workbook.add_worksheet('Cash Flow')
    Cash_Flow_1.merge_range("A1:C1", prop_name_cf_1, header_format_1)
    Cash_Flow_1.merge_range("A2:C2", header_2_cf_1, header_format_1)
    Cash_Flow_1.merge_range("A3:C3", header_3_cf_1, header_format_2)
    Cash_Flow_1.merge_range("A4:C4", header_4_cf_1, header_format_2)
    for row in range(3):
        if row == 0:
            Cash_Flow_1.write_blank(4, row, '', header_format_3)
            Cash_Flow_1.write_blank(5, row, '', header_format_3)
        else:
            Cash_Flow_1.write_string(4, row, header_cols_is[row - 1], header_format_3)
    Cash_Flow_1.write_formula(5, 1, '=+TEXT(RIGHT(A3,8),"mmmm yyyy")', header_format_3)
    Cash_Flow_1.write_formula(5, 2, '=+B6', header_format_3)
    Cash_Flow_1.merge_range(6, 0, 6, 9, '', header_format_2)
    Cash_Flow_1.set_row(6,7.5)
    row_write_val = 7
    for i in range(4, cash_flow_1_df.shape[0] - 1):
        if cash_flow_1_df['Col1'][i] == 'TOTAL OF ALL':
            break
        elif cash_flow_1_df['Header_Check'][i] == 1:
            try:
                next_header_val = cash_flow_1_df['Header_Check'][i+1]
                next_total_val = cash_flow_1_df['Total_Check'][i+1]
                next_nan_val = cash_flow_1_df['Nan_Var_Check'][i+1]
            except:
                next_header_val = 0
                next_total_val = 0
                next_nan_val = 0
            # try to get next vals for logic
            new_row_needed = 0
            if cash_flow_1_df['Col1'][i] in ['OPERATING INCOME', 'OPERATING EXPENSES', 'RECOVERABLE', 'NON-RECOVERABLE']:
                Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 9, '', header_format_2)
                Cash_Flow_1.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                if cash_flow_1_df['Total_Check'][i] == 1:
                    if next_total_val == 1:
                        Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                        Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df['Col2'][i],row_val_format_total_item_num)
                        Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col4'][i],row_val_format_total_item_num)
                        row_write_val = row_write_val + 1
                    else:
                        if cash_flow_1_df['Nan_Var_Check'][i] == 0:
                            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                            Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df['Col2'][i],row_val_format_total_item_num)
                            Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col4'][i],row_val_format_total_item_num)
                            row_write_val = row_write_val + 1
                            if next_header_val == 1:
                                new_row_needed = 1
                            else:
                                pass
                else:
                    if next_nan_val == 0:
                        Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                        row_write_val = row_write_val + 1
                    elif next_header_val == 0:
                        row_val_check = i
                        value_needed = 0
                        next_header_found = 0
                        while row_val_check <= cash_flow_1_df.shape[0] - 1 and next_header_found == 1:
                            if cash_flow_1_df['Header_Check'][row_val_check] == 1:
                                next_header_found = 1
                            elif cash_flow_1_df['Nan_Var_Check'][row_val_check] == 1:
                                value_needed = 1
                            else:
                                row_val_check = row_val_check + 1
                        if value_needed == 1:
                            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                            row_write_val = row_write_val + 1
                    else:
                        pass
            # ------------------------------------------
            # add a row or not
            if new_row_needed == 1:
                Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
                Cash_Flow_1.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                pass
        else:
            if cash_flow_1_df['Nan_Var_Check'][i] == 1:
                pass
            else:
                Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i],row_val_format_sub_item)
                Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df['Col2'][i],row_val_format_sub_item_num)
                Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col4'][i],row_val_format_sub_item_num)
                row_write_val = row_write_val + 1
    cf_bottom_format_1 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    cf_header_format_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    cf_header_format_3 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    cf_header_format_4 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    })
    cf_header_format_5 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    })
    cf_bottom_format_6 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    cf_header_format_7 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    cf_header_format_8 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    for i in range(cash_flow_1_df_end_of_page.shape[0]):
        if cash_flow_1_df_end_of_page['Col1'][i] in ['Year to Date', 'Period to Date']:
            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df_end_of_page['Col1'][i],cf_bottom_format_6)
            Cash_Flow_1.write_string(row_write_val, 1, 'Beginning Balance',cf_bottom_format_1)
            Cash_Flow_1.write_string(row_write_val, 2, 'Difference',cf_bottom_format_1)
            row_write_val = row_write_val + 1
        elif cash_flow_1_df_end_of_page['Col1'][i] == 'Cash Flow':
            Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
            Cash_Flow_1.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df_end_of_page['Col1'][i],cf_header_format_7)
            Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df_end_of_page['Col2'][i],cf_header_format_3)
            Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df_end_of_page['Col4'][i],cf_header_format_3)
            row_write_val = row_write_val + 1
            Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
            Cash_Flow_1.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
        elif cash_flow_1_df_end_of_page['Col1'][i] == 'Total Cash':
            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df_end_of_page['Col1'][i],cf_header_format_8)
            Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df_end_of_page['Col2'][i],cf_header_format_2)
            Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df_end_of_page['Col4'][i],cf_header_format_2)
            row_write_val = row_write_val + 1
            Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
            Cash_Flow_1.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
        else:
            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df_end_of_page['Col1'][i],cf_header_format_4)
            Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df_end_of_page['Col2'][i],cf_header_format_5)
            Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df_end_of_page['Col4'][i],cf_header_format_5)
            row_write_val = row_write_val + 1
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    prop_name_tb = trail_balance_df.columns[0]
    try:
        prop_name_tb = prop_name_bs_change('(', 1)[0]
    except:
        pass
    trail_balance_df = trail_balance_df.rename(columns={trail_balance_df.columns[0]: 'Col1'
                           , trail_balance_df.columns[1]: 'Col2'
                           , trail_balance_df.columns[2]: 'Col3'
                           , trail_balance_df.columns[3]: 'Col4'
                           , trail_balance_df.columns[4]: 'Col5'
                        })
    header_cols_tb_1 = ['Forward'
                  ,''
                  ,''
                  , 'Ending']
    header_cols_tb_2 = ['Balance'
                  ,'Debit'
                  ,'Credit'
                  ,'Balance']
    header_2_tb = trail_balance_df['Col1'][0]
    header_3_tb = trail_balance_df['Col1'][1]
    header_4_tb = trail_balance_df['Col1'][2]
    trail_balance_df=trail_balance_df.dropna(subset=['Col1']).reset_index(drop=True)
    trail_balance_df['Nan_Var_Check'] = trail_balance_df.apply(lambda x: flag_zero_vals_2(x['Col2'], x['Col3']), axis=1)
    # trial balance
    Trial_Balance = workbook.add_worksheet('Trial Balance')
    Trial_Balance.merge_range("A1:E1", prop_name_is, header_format_1)
    Trial_Balance.merge_range("A2:E2", header_2_is, header_format_1)
    Trial_Balance.merge_range("A3:E3", header_3_is, header_format_2)
    Trial_Balance.merge_range("A4:E4", header_4_is, header_format_2)
    for row in range(5):
        if row == 0:
            Trial_Balance.write_blank(4, row, '', header_format_3)
            Trial_Balance.write_blank(5, row, '', header_format_3)
        else:
            Trial_Balance.write_string(4, row, header_cols_tb_1[row - 1], header_format_3)
            Trial_Balance.write_string(5, row, header_cols_tb_2[row - 1], header_format_3)
    Trial_Balance.write_blank(6, 0, '', header_format_3)
    Trial_Balance.write_formula(6, 1, '=+EOMONTH(RIGHT(A3,8),-1)', header_format_3)
    Trial_Balance.write_blank(6, 2, '', header_format_3)
    Trial_Balance.write_blank(6, 3, '', header_format_3)
    Trial_Balance.write_formula(6, 4, '=+EOMONTH(B7,1)', header_format_3)
    Trial_Balance.merge_range(7, 0, 7, 4, '', header_format_2)
    Trial_Balance.set_row(7,7.5)
    row_write_val = 8
    for i in range(4, trail_balance_df.shape[0]):
        if trail_balance_df['Col1'][i] == 'Total':
            Trial_Balance.write_string(row_write_val, 0, trail_balance_df['Col1'][i], row_val_format_sub_item)
            Trial_Balance.write_number(row_write_val, 1, trail_balance_df['Col2'][i],row_val_format_total_item_num)
            Trial_Balance.write_number(row_write_val, 2, trail_balance_df['Col3'][i],row_val_format_total_item_num)
            Trial_Balance.write_number(row_write_val, 3, trail_balance_df['Col4'][i],row_val_format_total_item_num)
            Trial_Balance.write_number(row_write_val, 4, trail_balance_df['Col5'][i],row_val_format_total_item_num)
            row_write_val = row_write_val + 1
        else:
            if trail_balance_df['Nan_Var_Check'][i] == 1:
                pass
            else:
                Trial_Balance.write_string(row_write_val, 0, trail_balance_df['Col1'][i], row_val_format_sub_item_2)
                Trial_Balance.write_number(row_write_val, 1, trail_balance_df['Col2'][i],row_val_format_sub_item_num)
                Trial_Balance.write_number(row_write_val, 2, trail_balance_df['Col3'][i],row_val_format_sub_item_num)
                Trial_Balance.write_number(row_write_val, 3, trail_balance_df['Col4'][i],row_val_format_sub_item_num)
                Trial_Balance.write_number(row_write_val, 4, trail_balance_df['Col5'][i],row_val_format_sub_item_num)
                row_write_val = row_write_val + 1
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    #-------------------------------------------------------------------------
    header_1_pay_reg = payment_register_df.columns[0]
    payment_register_df = payment_register_df.rename(columns={payment_register_df.columns[0]: 'Col1'
                           , payment_register_df.columns[1]: 'Col2'
                           , payment_register_df.columns[2]: 'Col3'
                           , payment_register_df.columns[3]: 'Col4'
                           , payment_register_df.columns[4]: 'Col5'                           
                           , payment_register_df.columns[5]: 'Col6'
                           , payment_register_df.columns[6]: 'Col7'
                           , payment_register_df.columns[7]: 'Col8'
                           , payment_register_df.columns[8]: 'Col9'                           
                           , payment_register_df.columns[9]: 'Col10'
                           , payment_register_df.columns[10]: 'Col11'
                           , payment_register_df.columns[11]: 'Col12'
                           , payment_register_df.columns[12]: 'Col13'
                           , payment_register_df.columns[13]: 'Col14'
                        })
    header_cols_payment_register = ['Check #'
                        ,'Check'
                        ,'Bank Code'
                        ,'Payee Code'
                        ,'Payee Name'
                        ,'Check Date'
                        ,'Post Month'
                        ,'Payment Method'
                        ,'Payable'
                        ,'Property'
                        ,'Amount'
                        ,'Due To / Due From'
                        ,'Department'
                        ,'Notes'
                       ]
    header_cols_payment_register_2 = [''
                        ,'Control'
                        ,''
                        ,''
                        ,''
                        ,''
                        ,''
                        ,''
                        ,'Control#'
                        ,''
                        ,''
                        ,''
                        ,''
                        ,''
                       ]
    header_2_pay_reg = payment_register_df['Col1'][0]
    header_3_pay_reg = payment_register_df['Col1'][1]
    payment_register_df=payment_register_df.dropna(subset=['Col1']).reset_index(drop=True)
    payment_register_df['Nan_Var_Check'] = payment_register_df.apply(lambda x: flag_zero_vals_3(x['Col3'], x['Col11']), axis=1)
    pay_reg_df_meat = payment_register_df[payment_register_df['Nan_Var_Check'] == 0].reset_index(drop=True)
    # payment register payment_register_df payment_register_gray
    Payment_Register_Sheet =  workbook.add_worksheet('Payment Register')
    Payment_Register_Sheet.merge_range("A1:E1", header_1_pay_reg, header_format_1)
    Payment_Register_Sheet.merge_range("A2:E2", header_2_pay_reg, header_format_1)
    Payment_Register_Sheet.merge_range("A3:E3", header_3_pay_reg, header_format_2)
    for row in range(15):
        if row == 0:
            Payment_Register_Sheet.write_blank(3, row, '', header_format_3)
            Payment_Register_Sheet.write_blank(4, row, '', header_format_3)
        else:
            Payment_Register_Sheet.write_string(3, row, header_cols_payment_register[row - 1], header_format_3)
            Payment_Register_Sheet.write_string(4, row, header_cols_payment_register_2[row - 1], header_format_3)
    Payment_Register_Sheet.merge_range(5, 0, 5, 14, '', header_format_2)
    Payment_Register_Sheet.set_row(5,7.5)
    row_write_val = 6
    pr_header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':payment_register_gray
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'Center'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    })
    pr_header_format_2 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':payment_register_gray
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'Center'
                                    , 'num_format':'m/d/yyyy'
                                    , 'border_color':black_color
                                    })
    pr_header_format_3 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':payment_register_gray
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'Center'
                                    , 'num_format':'mm-yyyy'
                                    , 'border_color':black_color
                                    })
    pr_header_format_4 = workbook.add_format({'font_color': black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'Center'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    })
    
    pr_header_format_total = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'Center'
                                    , 'num_format':'_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    })
    for i in pay_reg_df_meat['Col1'].unique():
        if i == 'Grand Total ':
            print('nice')
        else:    
            loop_pay_reg_df = payment_register_df[payment_register_df['Col1'] == i].reset_index(drop=True)
            Payment_Register_Sheet.write_string(row_write_val, 0, loop_pay_reg_df['Col1'][0], pr_header_format_1)
            Payment_Register_Sheet.write_string(row_write_val, 1, loop_pay_reg_df['Col2'][0], pr_header_format_1)
            Payment_Register_Sheet.write_string(row_write_val, 2, loop_pay_reg_df['Col3'][0], pr_header_format_1)
            Payment_Register_Sheet.write_string(row_write_val, 3, loop_pay_reg_df['Col4'][0], pr_header_format_1)
            Payment_Register_Sheet.write_string(row_write_val, 4, loop_pay_reg_df['Col5'][0], pr_header_format_1)
            Payment_Register_Sheet.write_datetime(row_write_val, 5, loop_pay_reg_df['Col6'][0], pr_header_format_2)
            Payment_Register_Sheet.write_datetime(row_write_val, 6, loop_pay_reg_df['Col7'][0], pr_header_format_3)
            Payment_Register_Sheet.write_string(row_write_val, 7, loop_pay_reg_df['Col8'][0], pr_header_format_1)
            Payment_Register_Sheet.write_string(row_write_val, 8, loop_pay_reg_df['Col9'][0], pr_header_format_1)
            Payment_Register_Sheet.write_blank(row_write_val, 9, '', pr_header_format_1)
            Payment_Register_Sheet.write_blank(row_write_val, 10, '', pr_header_format_1)
            Payment_Register_Sheet.write_blank(row_write_val, 11, '', pr_header_format_1)
            Payment_Register_Sheet.write_blank(row_write_val, 12, '', pr_header_format_1)
            Payment_Register_Sheet.write_blank(row_write_val, 13, '', pr_header_format_1)
            Payment_Register_Sheet.write_blank(row_write_val, 14, '', pr_header_format_1)
            row_write_val = row_write_val + 1
            for i in range(loop_pay_reg_df.shape[0]):
                Payment_Register_Sheet.write_string(row_write_val, 9, loop_pay_reg_df['Col9'][i], pr_header_format_4)
                Payment_Register_Sheet.write_string(row_write_val, 10, loop_pay_reg_df['Col10'][i], pr_header_format_4)
                Payment_Register_Sheet.write_number(row_write_val, 11, loop_pay_reg_df['Col11'][i], pr_header_format_4)
                try:
                    Payment_Register_Sheet.write_string(row_write_val, 12, loop_pay_reg_df['Col12'][i], pr_header_format_4)
                except:
                    pass
                    #Payment_Register_Sheet.write(row_write_val, 12, loop_pay_reg_df['Col12'][i], pr_header_format_1)
                try:
                    Payment_Register_Sheet.write_string(row_write_val, 13, loop_pay_reg_df['Col13'][i], pr_header_format_4)
                except:
                    pass
                try:
                    Payment_Register_Sheet.write_string(row_write_val, 14, loop_pay_reg_df['Col14'][i], pr_header_format_4)
                except:
                    pass
                row_write_val = row_write_val + 1
            total_loop = loop_pay_reg_df['Col11'].sum()
            total_loop_string = 'Total '+ str(total_loop)
            Payment_Register_Sheet.write_string(row_write_val, 0, total_loop_string, pr_header_format_total)
            Payment_Register_Sheet.write_number(row_write_val, 11, total_loop, pr_header_format_total)
            row_write_val = row_write_val + 1
            Payment_Register_Sheet.merge_range(row_write_val, 0, row_write_val, 14, '', header_format_2)
            row_write_val = row_write_val + 1
            
    workbook.close()
