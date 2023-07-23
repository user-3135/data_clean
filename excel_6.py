import pandas as pd
import xlsxwriter as xl
import openpyxl
import math
# ------------------------------------------------
def rename_prop(property_name):
    try:
        property_name = property_name('(', 1)[0]
    except:
        try:
            name = ''
            for i in property_name:
                if i == '(':
                    break
                else:
                    name = name + i
            property_name = name
        except:
            pass
    return property_name
def clean_text_col_1(text_val):
    return_text_val = ''
    index_val = 0
    try:
        for i in text_val:
            if index_val > 2:
                return_text_val = return_text_val + str(i)
            index_val += 1
    except: pass
        
    return return_text_val
def flag_zero_vals(val_1, val_2):
    flag_zero_val = 0
    try:
        if(math.isnan(val_1)):
            if(math.isnan(val_2)):
                flag_zero_val = 1
    except:
        pass
    return flag_zero_val
def flag_zero_vals_2(val_1, val_2):
    flag_zero_val = 0
    if(val_1 == 0):
        if(val_2 == 0):
            flag_zero_val = 1
    return flag_zero_val
def flag_total_rows(val_1):
    flag_total_val = 0
    if 'TOTAL' in val_1:
        flag_total_val = 1
    return flag_total_val
def flag_header_rows(val_1):
    flag_total_val = 0
    try:
        val_1 = str(val_1)
        if val_1[0] != ' ':
            flag_total_val = 1
    except:
        pass
    return flag_total_val
def flag_header_rows_2(val_1):
    flag_total_val = 0
    try:
        if(math.isnan(val_1)):
            flag_total_val = 1
    except:
        pass
    return flag_total_val
def flag_box(val_1, val_2):
    flag_total_val = 0
    try:
        if(math.isnan(val_2)):
            if type(val_1) == str:
                flag_total_val = 1
    except:
        pass
    return flag_total_val
def check_header_ten_sch(val_1):
    return_val = 1
    try:
        if math.isnan(val_1):
            return_val = 0
    except:
        pass
    return return_val
def check_header_ten_sch(val_1):
    return_val = 1
    try:
        if math.isnan(val_1):
            return_val = 0
    except:
        pass
    return return_val
def flag_total_rows_2(val_1):
    flag_total_val = 0
    try:
        if 'Total' in val_1:
            flag_total_val = 1
    except:
        pass
    return flag_total_val
def flag_total_rows_3(val_1):
    flag_total_val = 0
    try:
        if 'Net Change' in val_1:
            flag_total_val = 1
    except:
        pass
    return flag_total_val
# ----------------------------------------------------------------------------------------------------------------------
def income_statement_v2(workbook, income_statement_1, Income_Statement):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    payment_register_gray = '#F2F2F2'
    prop_name_is = income_statement_1.columns[0]
    prop_name_is = str(prop_name_is)
    def rename_prop(property_name):
        try:
            property_name = property_name('(', 1)[0]
        except:
            try:
                name = ''
                for i in property_name:
                    if i == '(':
                        break
                    else:
                        name = name + i
                property_name = name
            except:
                pass
        return property_name
    try:
        prop_name_is = rename_prop(prop_name_is)
    except:
        pass
    income_statement_1 = income_statement_1.rename(columns={income_statement_1.columns[0]: 'Col1'
                           , income_statement_1.columns[1]: 'Col2'
                           , income_statement_1.columns[2]: 'Col3'
                           , income_statement_1.columns[3]: 'Col4'
                           , income_statement_1.columns[4]: 'Col5'
                           })
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    header_format_2 = workbook.add_format({'font_color': black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    header_cols_is = ['Month to Date'
                  ,'Year to Date']
    header_2_is = income_statement_1['Col1'][0]
    header_3_is = income_statement_1['Col1'][1]
    header_4_is = income_statement_1['Col1'][2]
    income_statement_1=income_statement_1.dropna(subset=['Col1']).reset_index(drop=True)
    def clean_text_col_1(text_val):
        return_text_val = ''
        index_val = 0
        for i in text_val:
            if index_val > 2:
                return_text_val = return_text_val + str(i)
            index_val += 1
        return return_text_val
    income_statement_1['Col1'] = income_statement_1.apply(lambda x: clean_text_col_1(x['Col1']), axis=1)
    def flag_zero_vals_2(val_1, val_2):
        flag_zero_val = 0
        if(val_1 == 0):
            if(val_2 == 0):
                flag_zero_val = 1
        return flag_zero_val
    income_statement_1['Nan_Var_Check'] = income_statement_1.apply(lambda x: flag_zero_vals_2(x['Col2'], x['Col4']), axis=1)
    def flag_total_rows(val_1):
        flag_total_val = 0
        try:
            if 'TOTAL' in val_1:
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    income_statement_1['Total_Check'] = income_statement_1.apply(lambda x: flag_total_rows(x['Col1']), axis=1)
    def flag_header_rows(val_1):
        flag_total_val = 0
        try:
            if val_1[0] != ' ':
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    income_statement_1['Header_Check'] = income_statement_1.apply(lambda x: flag_header_rows(x['Col1']), axis=1)
    row_val_format_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    row_val_format_sub_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                     })
    row_val_format_total_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_item_num_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    , 'top':1
                                    , 'bottom':1
                                    })
    row_val_format_header_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'top':1
                                    , 'bottom':1
                                     })
    row_val_format_total_item_num_3 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    })
    row_val_format_header_3 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                     })
    row_val_format_total_item_num_4 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    , 'bottom':1
                                    })
    row_val_format_header_4 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'bottom':1
                                     })
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
    Income_Statement.merge_range(6, 0, 6, 2, '', header_format_2)
    Income_Statement.set_row(6,7.5)
    row_write_val = 7
    for i in range(3, income_statement_1.shape[0]):
        try:
            next_row_val = income_statement_1['Col1'][i + 1]
        except:
            next_row_val = ''
        if income_statement_1['Col1'][i] == 'AL OF ALL':
            pass
        elif income_statement_1['Col1'][i] == 'TOTAL OF ALL':
            pass
        elif income_statement_1['Col1'][i] == 'TOTAL NET INCOME':
            Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header_3)
            Income_Statement.write_number(row_write_val, 1, income_statement_1['Col2'][i], row_val_format_total_item_num_3)
            Income_Statement.write_number(row_write_val, 2, income_statement_1['Col4'][i], row_val_format_total_item_num_3)
            row_write_val = row_write_val + 1
        elif income_statement_1['Col1'][i] == 'AL OF ALL':
            pass
        elif next_row_val == 'TOTAL NET INCOME':
            Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header_4)
            Income_Statement.write_number(row_write_val, 1, income_statement_1['Col2'][i], row_val_format_total_item_num_4)
            Income_Statement.write_number(row_write_val, 2, income_statement_1['Col4'][i], row_val_format_total_item_num_4)
            row_write_val = row_write_val + 1
        elif income_statement_1['Col1'][i] in ['NON-OPERATING EXPENSES', 'OPERATING EXPENSES']:
            if income_statement_1['Col1'][i] == 'OPERATING EXPENSES':
                Income_Statement.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
                Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                Income_Statement.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                Income_Statement.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
                Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
        elif income_statement_1['Header_Check'][i] == 1:
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
                Income_Statement.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
                Income_Statement.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            elif income_statement_1['Col1'][i] in ['TOTAL OPERATING INCOME', 'TOTAL OPERATING EXPENSES', 'TOTAL NET OPERATING INCOME']:
                Income_Statement.write_string(row_write_val, 0, income_statement_1['Col1'][i], row_val_format_header_2)
                Income_Statement.write_number(row_write_val, 1, income_statement_1['Col2'][i], row_val_format_total_item_num_2)
                Income_Statement.write_number(row_write_val, 2, income_statement_1['Col4'][i], row_val_format_total_item_num_2)
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
                Income_Statement.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
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
    column_width_list = [
                [71, 0, 0, Income_Statement]
                ,[17.7, 1, 2, Income_Statement]
    ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass
    Income_Statement.set_portrait()
    Income_Statement.set_margins(.5,.5,.5,.5)
    Income_Statement.repeat_rows(0, 6)
    Income_Statement.print_area(0,0, row_write_val, 2)
    Income_Statement.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    Income_Statement.fit_to_pages(1, total_pages)
    return income_statement_1
# ----------------------------------------------------------------- budget comp
def budget_comp_sheet_creation_v2(workbook, df, worksheet):
    prop_name = df.columns[0]
    def rename_prop(property_name):
        try:
            property_name = property_name('(', 1)[0]
        except:
            try:
                name = ''
                for i in property_name:
                    if i == '(':
                        break
                    else:
                        name = name + i
                property_name = name
            except:
                pass
        return property_name
    try:
        prop_name = rename_prop(prop_name)
    except:
        pass
    # -------------------------------------
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    payment_register_gray = '#F2F2F2'
    def clean_text_col_1(text_val):
        return_text_val = ''
        index_val = 0
        for i in text_val:
            if index_val > 2:
                return_text_val = return_text_val + str(i)
            index_val += 1
        return return_text_val
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
    def flag_total_rows(val_1):
        flag_total_val = 0
        try:
            if 'TOTAL' in val_1:
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    def flag_header_rows(val_1):
        flag_total_val = 0
        try:
            if val_1[0] != ' ':
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    header_format_2 = workbook.add_format({'font_color': black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'text_wrap':True
                                     })
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    row_val_format_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_sub_item = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_sub_item_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_sub_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_sub_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_total_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'valign':'vcenter'
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
                                    , 'valign':'vcenter'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    row_val_format_header_top_bottom = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'valign':'vcenter'
                                     })
    row_val_format_total_item_num_top_bottom = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'valign':'vcenter'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent_top_bottom = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'valign':'vcenter'
                                     })
    row_val_format_header_top_bottom_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'valign':'vcenter'
                                     })
    row_val_format_total_item_num_top_bottom_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'valign':'vcenter'
                                    })
    row_val_format_total_item_percent_top_bottom_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'border_color':black_color
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'valign':'vcenter'
                                     })
    row_val_format_header_top_bottom_3 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bottom':1
                                    , 'valign':'vcenter'
                                     })
    #--------------------------------------------------------------------------------------------------------------------
    header_format_1_base = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    header_format_2_base = workbook.add_format({'font_color': black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'text_wrap':True
                                     })
    header_format_3_base = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    row_val_format_header_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_sub_item_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_sub_item_2_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_sub_item_num_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_sub_item_percent_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_total_item_num_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'valign':'vcenter'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'valign':'vcenter'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    row_val_format_header_top_bottom_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'valign':'vcenter'
                                     })
    row_val_format_total_item_num_top_bottom_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'valign':'vcenter'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent_top_bottom_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'valign':'vcenter'
                                     })
    row_val_format_header_top_bottom_2_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'valign':'vcenter'
                                     })
    row_val_format_total_item_num_top_bottom_2_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'valign':'vcenter'
                                    })
    row_val_format_total_item_percent_top_bottom_2_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'border_color':black_color
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'valign':'vcenter'
                                     })
    row_val_format_header_top_bottom_3_base = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bottom':1
                                    , 'valign':'vcenter'
                                     })
    df = df.rename(columns={df.columns[0]: 'Col1'
                           , df.columns[1]: 'Col2'
                           , df.columns[2]: 'Col3'
                           , df.columns[3]: 'Col4'
                           , df.columns[4]: 'Col5'
                           , df.columns[5]: 'Col6'
                           , df.columns[6]: 'Col7'
                           , df.columns[7]: 'Col8'
                           , df.columns[8]: 'Col9'
                           , df.columns[9]: 'Col10'
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
    df=df.dropna(subset=['Col1']).reset_index(drop=True)
    df['Col1'] = df.apply(lambda x: clean_text_col_1(x['Col1']), axis=1)
    df['Nan_Var_Check'] = df.apply(lambda x: flag_zero_vals(x['Col5'], x['Col9']), axis=1)
    df['Total_Check'] = df.apply(lambda x: flag_total_rows(x['Col1']), axis=1)
    df['Header_Check'] = df.apply(lambda x: flag_header_rows(x['Col1']), axis=1)
    worksheet.merge_range("A1:I1", prop_name, header_format_1)
    worksheet.merge_range("A2:I2", header_2, header_format_1)
    worksheet.merge_range("A3:I3", header_3, header_format_2)
    worksheet.merge_range("A4:I4", header_4, header_format_2)
    for row in range(10):
        if row == 0:
            worksheet.write_blank(4, row, '', header_format_3)
        else:
            worksheet.write_string(4, row, header_cols[row - 1], header_format_3)
    worksheet.merge_range(5, 0, 5, 9, '', header_format_2)
    worksheet.set_row(5,7.5)
    row_write_val = 6
    for i in range(3, df.shape[0]):
        #print(df['Col1'][i])
        if df['Col1'][i] == 'TOTAL NET INCOME':
            if 1 == 1:
                worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header_top_bottom_2_base)
                worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num_top_bottom_2_base)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num_top_bottom_2_base)
                worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num_top_bottom_2)
                try:
                    worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent_top_bottom_2)
                except:
                    worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent_top_bottom_2)
                worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num_top_bottom_2_base)
                worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num_top_bottom_2_base)
                worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num_top_bottom_2)
                try:
                    worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_total_item_percent_top_bottom_2)
                except:
                    worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent_top_bottom_2)
                row_write_val = row_write_val + 1
        elif df['Header_Check'][i] == 1:
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
            try:
                next_col_1 = df['Col1'][i + 1]
            except:
                next_col_1 = ''
            if next_col_1 == 'TOTAL NET INCOME':
                worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header_top_bottom_base)
                worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num_top_bottom_base)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num_top_bottom_base)
                worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num_top_bottom)
                try:
                    worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent_top_bottom)
                except:
                    worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent_top_bottom)
                worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num_top_bottom_base)
                worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num_top_bottom_base)
                worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num_top_bottom)
                try:
                    worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_total_item_percent_top_bottom)
                except:
                    worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent_top_bottom)
                row_write_val = row_write_val + 1
            elif df['Col1'][i] == 'AL OF ALL':
                pass
            elif df['Col1'][i] in ['OPERATING INCOME', 'OPERATING EXPENSES', 'RECOVERABLE', 'NON-RECOVERABLE']:
                worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                worksheet.merge_range(row_write_val, 0, row_write_val, 9, '', header_format_2)
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            elif df['Col1'][i] in ['TOTAL OPERATING INCOME', 'TOTAL OPERATING EXPENSES', 'TOTAL NET OPERATING INCOME']:
                worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header_top_bottom_base)
                worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num_top_bottom_base)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num_top_bottom_base)
                worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num_top_bottom)
                try:
                    worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent_top_bottom)
                except:
                    worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent_top_bottom)
                worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num_top_bottom_base)
                worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num_top_bottom_base)
                worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num_top_bottom)
                try:
                    worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_total_item_percent_top_bottom)
                except:
                    worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent_top_bottom)
                row_write_val = row_write_val + 1
                if df['Col1'][i] != 'TOTAL OPERATING EXPENSES':
                    worksheet.set_row(row_write_val,7.5)
                    row_write_val = row_write_val + 1
            else:
                if df['Total_Check'][i] == 1:
                    if next_total_val == 1:
                        worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header_base)
                        worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num_base)
                        worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num_base)
                        worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num)
                        try:
                            worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent)
                        except:
                            worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent)
                        worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num_base)
                        worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num_base)
                        worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num)
                        try:
                            worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_total_item_percent)
                        except:
                            worksheet.write_number(row_write_val, 8, 0,row_val_format_total_item_percent)
                        row_write_val = row_write_val + 1
                    else:
                        if df['Nan_Var_Check'][i] >= 0:
                            worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header_base)
                            worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num_base)
                            worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num_base)
                            worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num)
                            try:
                                worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_total_item_percent)
                            except:
                                worksheet.write_number(row_write_val, 4, 0,row_val_format_total_item_percent)
                            worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num_base)
                            worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num_base)
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
                    if next_nan_val >= 0:
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
                            #worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
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
            #print(df['Col1'][i])
            if df['Nan_Var_Check'][i] == 2:
                pass
            else:
                worksheet.write_string(row_write_val, 0, df['Col1'][i],row_val_format_sub_item_base)
                worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_sub_item_num_base)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_sub_item_num_base)
                worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_sub_item_num)
                try:
                    worksheet.write_number(row_write_val, 4, df['Col5'][i]/100,row_val_format_sub_item_percent)
                except:
                    worksheet.write_number(row_write_val, 4, 0,row_val_format_sub_item_percent)
                worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_sub_item_num_base)
                worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_sub_item_num_base)
                worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_sub_item_num)
                try:
                    worksheet.write_number(row_write_val, 8, df['Col9'][i]/100,row_val_format_sub_item_percent)
                except:
                    worksheet.write_number(row_write_val, 8, 0,row_val_format_sub_item_percent)
                try:
                    worksheet.write_string(row_write_val, 9, df['Col10'][i], header_format_2)
                except:
                    worksheet.write_blank(row_write_val, 9, '', header_format_2)
                row_write_val = row_write_val + 1
    worksheet.set_column(0,0,49.29)
    worksheet.set_column(9,9,49.29)
    worksheet.set_column(1,8,15)
    worksheet.print_area(0,0, row_write_val, 9)
    worksheet.set_margins(.5,.5,.5,.5)
    worksheet.repeat_rows(0, 5)
    num_pages_budget_1 = math.ceil(row_write_val/65)
    worksheet.fit_to_pages(1, num_pages_budget_1)
    worksheet.set_landscape()
    worksheet.set_page_view(2)
    return df
# ----------------------------------------------------------------- new stuff create_xl_balance_sheet_v2
def create_xl_balance_sheet_v2(workbook, bal_sheet_1, BS_Change_Sheet):
    prop_name_bs_change = bal_sheet_1.columns[0]
    def rename_prop(property_name):
        try:
            property_name = property_name('(', 1)[0]
        except:
            try:
                name = ''
                for i in property_name:
                    if i == '(':
                        break
                    else:
                        name = name + i
                property_name = name
            except:
                pass
        return property_name
    try:
        prop_name_bs_change = rename_prop(prop_name_bs_change) 
        # prop_name_bs_change('(', 1)[0]
    except:
        pass
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    payment_register_gray = '#F2F2F2'
    def clean_text_col_1(text_val):
        return_text_val = ''
        index_val = 0
        for i in text_val:
            if index_val > 2:
                return_text_val = return_text_val + str(i)
            index_val += 1
        return return_text_val
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
    def flag_total_rows(val_1):
        flag_total_val = 0
        try:
            if 'TOTAL' in val_1:
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    def flag_header_rows(val_1):
        flag_total_val = 0
        try:
            if val_1[0] != ' ':
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    header_format_2 = workbook.add_format({'font_color': black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                    , 'num_format':'[$-en-US]mmmm d, yyyy;@'
                                     })
    row_val_format_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    row_val_format_sub_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                     })
    row_val_format_total_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    row_val_format_header_l_and_a = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'top':1
                                    , 'bottom':1 
                                     })
    row_val_format_sub_item_num_l_and_a = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                     })
    row_val_format_header_equity = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border':6
                                    , 'left':0
                                    , 'right':0
                                    , 'top':0
                                     })
    row_val_format_sub_item_num_equity = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border':6
                                    , 'left':0
                                    , 'right':0
                                    , 'top':0
                                     })
    row_val_format_sub_item_equity_top = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_item_num_equity_top = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_sub_item_equity_top_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_item_num_equity_top_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
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
    BS_Change_Sheet.merge_range("A1:D1", prop_name_bs_change, header_format_1)
    BS_Change_Sheet.merge_range("A2:D2", header_2_bal_sheet_1, header_format_1)
    BS_Change_Sheet.merge_range("A3:D3", header_3_bal_sheet_1, header_format_2)
    BS_Change_Sheet.merge_range("A4:D4", header_4_bal_sheet_1, header_format_2)
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
    for i in range(3, bal_sheet_1.shape[0]):
        if bal_sheet_1['Col1'][i] in ['EQUITY', 'ASSETS', 'LIABILITIES  EQUITY', 'LIABILITIES']:
            BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
            row_write_val = row_write_val + 1
            if bal_sheet_1['Col1'][i] != 'EQUITY':
                BS_Change_Sheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
        elif bal_sheet_1['Header_Check'][i] == 1:
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
            elif bal_sheet_1['Col1'][i] in ['TOTAL ASSETS', 'TOTAL LIABILITIES']:
                BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header_l_and_a)
                BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i], row_val_format_sub_item_num_l_and_a)
                BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i], row_val_format_sub_item_num_l_and_a)
                BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i], row_val_format_sub_item_num_l_and_a)
                row_write_val = row_write_val + 1
                BS_Change_Sheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            elif bal_sheet_1['Col1'][i] == 'TOTAL EQUITY':
                BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_sub_item_equity_top)
                BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i],row_val_format_total_item_num_equity_top)
                BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i],row_val_format_total_item_num_equity_top)
                BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i],row_val_format_total_item_num_equity_top)
                row_write_val = row_write_val + 1
            elif bal_sheet_1['Col1'][i] == 'TOTAL LIABILITIES  EQUITY':
                BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header_equity)
                BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i],row_val_format_sub_item_num_equity)
                BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i],row_val_format_sub_item_num_equity)
                BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i],row_val_format_sub_item_num_equity)
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
                        if bal_sheet_1['Nan_Var_Check'][i] >= 0 :
                            BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
                            BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i],row_val_format_total_item_num)
                            BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i],row_val_format_total_item_num)
                            BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i],row_val_format_total_item_num)
                            row_write_val = row_write_val + 1
                            if next_header_val == 1:
                                new_row_needed = 1
                            else:
                                pass
                elif next_header_val == 0:
                    row_val_check = i
                    value_needed = 0
                    next_header_found = 0
                    while row_val_check < bal_sheet_1.shape[0] - 1 and next_header_found == 0:
                        if bal_sheet_1['Header_Check'][row_val_check] == 1:
                            next_header_found = 1
                        elif bal_sheet_1['Nan_Var_Check'][row_val_check] == 1:
                            value_needed = 1
                        else:
                            row_val_check = row_val_check + 1
                    if value_needed == 0:
                        BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
                        row_write_val = row_write_val + 1
                else:
                    if next_nan_val == 0:
                        if next_total_val == 1:
                            pass
                        else:
                            BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
                            row_write_val = row_write_val + 1
                    elif next_header_val == 0:
                        row_val_check = i
                        value_needed = 0
                        next_header_found = 0
                        while row_val_check <= bal_sheet_1.shape[0] - 1 and next_header_found == 0:
                            if bal_sheet_1['Header_Check'][row_val_check] == 1:
                                next_header_found = 1
                            elif bal_sheet_1['Nan_Var_Check'][row_val_check] == 1:
                                value_needed = 1
                            else:
                                row_val_check = row_val_check + 1
                        if value_needed == 0:
                            BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i], row_val_format_header)
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
            #print(bal_sheet_1['Col1'][i])
            try:
                next_val = bal_sheet_1['Col1'][i + 1]
            except:
                next_val = ''
            if bal_sheet_1['Nan_Var_Check'][i] == 2:
                #print(bal_sheet_1['Col1'][i])
                pass
            elif next_val == 'TOTAL EQUITY':
                BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i],row_val_format_sub_item_equity_top_2)
                BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i],row_val_format_total_item_num_equity_top_2)
                BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i],row_val_format_total_item_num_equity_top_2)
                BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i],row_val_format_total_item_num_equity_top_2)
                row_write_val = row_write_val + 1
            else:
                BS_Change_Sheet.write_string(row_write_val, 0, bal_sheet_1['Col1'][i],row_val_format_sub_item)
                BS_Change_Sheet.write_number(row_write_val, 1, bal_sheet_1['Col2'][i],row_val_format_sub_item_num)
                BS_Change_Sheet.write_number(row_write_val, 2, bal_sheet_1['Col3'][i],row_val_format_sub_item_num)
                BS_Change_Sheet.write_number(row_write_val, 3, bal_sheet_1['Col4'][i],row_val_format_sub_item_num)
                row_write_val = row_write_val + 1
    column_width_list = [
                [50, 0, 0, BS_Change_Sheet] ## Trial Balance
                ,[18, 1, 4, BS_Change_Sheet] ## Trial Balance
    ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass
    BS_Change_Sheet.set_portrait()
    BS_Change_Sheet.set_margins(.5,.5,.5,.5)
    BS_Change_Sheet.repeat_rows(0, 6)
    BS_Change_Sheet.print_area(0,0, row_write_val, 3)
    BS_Change_Sheet.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    BS_Change_Sheet.fit_to_pages(1, total_pages)
    return bal_sheet_1
# ----------------------------------------------------------------- new stuff create_xl_tb_v2
def create_xl_tb_v2(workbook, trail_balance_df, Trial_Balance):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    payment_register_gray = '#F2F2F2'
    def clean_text_col_1(text_val):
        return_text_val = ''
        index_val = 0
        for i in text_val:
            if index_val > 2:
                return_text_val = return_text_val + str(i)
            index_val += 1
        return return_text_val
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
    def flag_total_rows(val_1):
        flag_total_val = 0
        try:
            if 'TOTAL' in val_1:
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    def flag_header_rows(val_1):
        flag_total_val = 0
        try:
            if val_1[0] != ' ':
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    header_format_2 = workbook.add_format({'font_color': black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                    , 'num_format':'[$-en-US]mmmm d, yyyy;@'
                                     })
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
    row_val_format_sub_item_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    row_val_format_sub_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                     })
    row_val_format_total_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    row_val_format_sub_item_prior_total = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_num_prior_total = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_sub_item_total = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_num_total = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                    })
    def rename_prop(property_name):
        try:
            property_name = property_name('(', 1)[0]
        except:
            try:
                name = ''
                for i in property_name:
                    if i == '(':
                        break
                    else:
                        name = name + i
                property_name = name
            except:
                pass
        return property_name
    prop_name_tb = trail_balance_df.columns[0]
    try:
        prop_name_tb = rename_prop(prop_name_tb)
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
    Trial_Balance.merge_range("A1:E1", prop_name_tb, header_format_1)
    Trial_Balance.merge_range("A2:E2", header_2_tb, header_format_1)
    Trial_Balance.merge_range("A3:E3", header_3_tb, header_format_2)
    Trial_Balance.merge_range("A4:E4", header_4_tb, header_format_2)
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
    for i in range(3, trail_balance_df.shape[0]):
        if trail_balance_df['Col1'][i] == 'Total':
            previous_value = trail_balance_df['Col1'][i - 1]
            if previous_value != 'Total':
                Trial_Balance.write_string(row_write_val, 0, trail_balance_df['Col1'][i], row_val_format_sub_item_total)
                Trial_Balance.write_number(row_write_val, 1, trail_balance_df['Col2'][i], row_val_format_total_item_num_total)
                Trial_Balance.write_number(row_write_val, 2, trail_balance_df['Col3'][i], row_val_format_total_item_num_total)
                Trial_Balance.write_number(row_write_val, 3, trail_balance_df['Col4'][i], row_val_format_total_item_num_total)
                Trial_Balance.write_number(row_write_val, 4, trail_balance_df['Col5'][i], row_val_format_total_item_num_total)
                row_write_val = row_write_val + 1
        elif trail_balance_df['Col1'][i + 1] == 'Total':
            Trial_Balance.write_string(row_write_val, 0, trail_balance_df['Col1'][i], row_val_format_sub_item_prior_total)
            Trial_Balance.write_number(row_write_val, 1, trail_balance_df['Col2'][i],row_val_format_total_item_num_prior_total)
            Trial_Balance.write_number(row_write_val, 2, trail_balance_df['Col3'][i],row_val_format_total_item_num_prior_total)
            Trial_Balance.write_number(row_write_val, 3, trail_balance_df['Col4'][i],row_val_format_total_item_num_prior_total)
            Trial_Balance.write_number(row_write_val, 4, trail_balance_df['Col5'][i],row_val_format_total_item_num_prior_total)
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
    column_width_list = [
                [38.7, 0, 0, Trial_Balance] ## Trial Balance
                ,[18, 1, 4, Trial_Balance] ## Trial Balance
    ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass
    Trial_Balance.set_portrait()
    Trial_Balance.set_margins(.5,.5,.5,.5)
    Trial_Balance.repeat_rows(0, 6)
    Trial_Balance.print_area(0,0, row_write_val, 4)
    Trial_Balance.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    Trial_Balance.fit_to_pages(1, total_pages)
    return trail_balance_df
# ----------------------------------------------------------------- new stuff create_xl_cf_v2
def create_xl_cf_v2(workbook, cash_flow_1_df, Cash_Flow_1):
    prop_name_cf_1 = cash_flow_1_df.columns[0]
    try:
        prop_name_cf_1 = rename_prop(prop_name_cf_1)
    except:
        pass
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    payment_register_gray = '#F2F2F2'
    def clean_text_col_1(text_val):
        return_text_val = ''
        index_val = 0
        for i in text_val:
            if index_val > 2:
                return_text_val = return_text_val + str(i)
            index_val += 1
        return return_text_val
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
    def flag_total_rows(val_1):
        flag_total_val = 0
        try:
            if 'TOTAL' in val_1:
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    def flag_header_rows(val_1):
        flag_total_val = 0
        try:
            if val_1[0] != ' ':
                flag_total_val = 1
        except:
            pass
        return flag_total_val
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    header_format_2 = workbook.add_format({'font_color': black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    row_val_format_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_header_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    row_val_format_sub_item = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    row_val_format_sub_item_percent = workbook.add_format({'font_color': black_color
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
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_num_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_num_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_num_2_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_num_total_net_income = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border':6
                                    , 'left':0
                                    , 'top':0
                                    , 'right':0
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_num_total_net_income_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'border':6
                                    , 'left':0
                                    , 'top':0
                                    , 'right':0
                                    , 'border_color':black_color
                                    })
    row_val_format_header_num_total_net_income = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border':6
                                    , 'left':0
                                    , 'top':0
                                    , 'right':0
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
    row_val_format_header_prior_net_income = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'bottom':1
                                     })
    row_val_format_total_item_num_prior_net_income = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_num_prior_net_income_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    
    cash_flow_1_df = cash_flow_1_df.rename(columns={cash_flow_1_df.columns[0]: 'Col1'
                           , cash_flow_1_df.columns[1]: 'Col2'
                           , cash_flow_1_df.columns[2]: 'Col3'
                           , cash_flow_1_df.columns[3]: 'Col4'
                           , cash_flow_1_df.columns[4]: 'Col5'
                        })
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
    # write
    Cash_Flow_1.merge_range("A1:C1", prop_name_cf_1, header_format_1)
    Cash_Flow_1.merge_range("A2:C2", header_2_cf_1, header_format_1)
    Cash_Flow_1.merge_range("A3:C3", header_3_cf_1, header_format_2)
    Cash_Flow_1.merge_range("A4:C4", header_4_cf_1, header_format_2)
    header_cols = [ 'Month to Date', 'Year to Date']
    for row in range(4):
        if row <= 1:
            Cash_Flow_1.write_blank(4, row, '', header_format_3)
            Cash_Flow_1.write_blank(5, row, '', header_format_3)
        else:
            Cash_Flow_1.write_string(4, row, header_cols[row - 2], header_format_3)
    Cash_Flow_1.write_formula(5, 2, '=+TEXT(RIGHT(A3,8),"mmmm yyyy")', header_format_3)
    Cash_Flow_1.write_formula(5, 3, '=+C6', header_format_3)
    Cash_Flow_1.merge_range(6, 0, 6, 4, '', header_format_2)
    Cash_Flow_1.set_row(6,7.5)
    row_write_val = 7
    for i in range(3, cash_flow_1_df.shape[0] - 1):
        if cash_flow_1_df['Col1'][i] == 'TOTAL OF ALL':
            break
        elif cash_flow_1_df['Col1'][i] in ['NON-OPERATING EXPENSES', 'OPERATING EXPENSES']:
            if cash_flow_1_df['Col1'][i] == 'OPERATING EXPENSES':
                Cash_Flow_1.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
                Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                Cash_Flow_1.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                Cash_Flow_1.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
                Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
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
                Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 3, '', header_format_2)
                Cash_Flow_1.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            elif cash_flow_1_df['Col1'][i] in ['TOTAL OPERATING INCOME', 'TOTAL OPERATING EXPENSES', 'TOTAL NET OPERATING INCOME', 'TOTAL NET INCOME']:
                if cash_flow_1_df['Col1'][i] == 'TOTAL NET INCOME':
                    Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header_num_total_net_income)
                    Cash_Flow_1.write_blank(row_write_val, 1, '', row_val_format_header_num_total_net_income)
                    Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col2'][i], row_val_format_total_item_num_total_net_income)
                    Cash_Flow_1.write_number(row_write_val, 3, cash_flow_1_df['Col4'][i], row_val_format_total_item_num_total_net_income)
                    row_write_val = row_write_val + 1
                else:
                    Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header_2)
                    Cash_Flow_1.write_blank(row_write_val, 1, '', row_val_format_header_2)
                    Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col2'][i], row_val_format_total_item_num_2)
                    Cash_Flow_1.write_number(row_write_val, 3, cash_flow_1_df['Col4'][i], row_val_format_total_item_num_2)
                    row_write_val = row_write_val + 1
            else:
                if cash_flow_1_df['Total_Check'][i] == 1:
                    if cash_flow_1_df['Col1'][i + 1] == 'TOTAL NET INCOME': # row_val_format_total_item_num_prior_net_income_percent
                        Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header_prior_net_income)
                        Cash_Flow_1.write_blank(row_write_val, 1, '', row_val_format_header_prior_net_income)
                        Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col2'][i], row_val_format_total_item_num_prior_net_income)
                        Cash_Flow_1.write_number(row_write_val, 3, cash_flow_1_df['Col4'][i], row_val_format_total_item_num_prior_net_income)
                        row_write_val = row_write_val + 1
                    elif next_total_val == 1:
                        Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                        Cash_Flow_1.write_blank(row_write_val, 1, '', row_val_format_header)
                        Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col2'][i], row_val_format_total_item_num)
                        Cash_Flow_1.write_number(row_write_val, 3, cash_flow_1_df['Col4'][i], row_val_format_total_item_num)
                        row_write_val = row_write_val + 1
                    else:
                        if cash_flow_1_df['Nan_Var_Check'][i] == 0:
                            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df['Col1'][i], row_val_format_header)
                            Cash_Flow_1.write_blank(row_write_val, 1, '', row_val_format_header)
                            Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col2'][i],row_val_format_total_item_num)
                            Cash_Flow_1.write_number(row_write_val, 3, cash_flow_1_df['Col4'][i],row_val_format_total_item_num)
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
                Cash_Flow_1.write_blank(row_write_val, 1, '', row_val_format_sub_item)
                Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df['Col2'][i],row_val_format_sub_item_num)
                Cash_Flow_1.write_number(row_write_val, 3, cash_flow_1_df['Col4'][i],row_val_format_sub_item_num)
                row_write_val = row_write_val + 1
    cf_bottom_format_1 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    cf_header_format_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    cf_header_format_3 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'align':'right'
                                    })
    cf_header_format_3_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'align':'right'
                                    })
    cf_header_format_4 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                    })
    cf_header_format_5 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
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
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
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
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    for i in range(cash_flow_1_df_end_of_page.shape[0]):
        if cash_flow_1_df_end_of_page['Col1'][i] in ['Year to Date', 'Period to Date']:
            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df_end_of_page['Col1'][i],cf_bottom_format_6)
            Cash_Flow_1.write_string(row_write_val, 1, 'Beginning Balance',cf_bottom_format_1)
            Cash_Flow_1.write_string(row_write_val, 2, 'Ending Balance',cf_bottom_format_1)
            Cash_Flow_1.write_string(row_write_val, 3, 'Difference',cf_bottom_format_1)
            row_write_val = row_write_val + 1
        elif cash_flow_1_df_end_of_page['Col1'][i] == 'Cash Flow':
            Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
            Cash_Flow_1.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df_end_of_page['Col1'][i],cf_header_format_7)
            Cash_Flow_1.write_blank(row_write_val, 0, '',cf_header_format_7)
            Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df_end_of_page['Col2'][i],cf_header_format_3) #cf_header_
            Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df_end_of_page['Col4'][i],cf_header_format_3)
            row_write_val = row_write_val + 1
            Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
            Cash_Flow_1.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
        elif cash_flow_1_df_end_of_page['Col1'][i] == 'Total Cash':
            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df_end_of_page['Col1'][i],cf_header_format_8)
            Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df_end_of_page['Col2'][i],cf_header_format_2)
            Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df_end_of_page['Col3'][i],cf_header_format_2)
            Cash_Flow_1.write_number(row_write_val, 3, cash_flow_1_df_end_of_page['Col4'][i],cf_header_format_2)
            row_write_val = row_write_val + 1
            Cash_Flow_1.merge_range(row_write_val, 0, row_write_val, 2, '', header_format_2)
            Cash_Flow_1.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
        else:
            Cash_Flow_1.write_string(row_write_val, 0, cash_flow_1_df_end_of_page['Col1'][i],cf_header_format_4)
            Cash_Flow_1.write_number(row_write_val, 1, cash_flow_1_df_end_of_page['Col2'][i],cf_header_format_5)
            Cash_Flow_1.write_number(row_write_val, 2, cash_flow_1_df_end_of_page['Col3'][i],cf_header_format_5)
            Cash_Flow_1.write_number(row_write_val, 3, cash_flow_1_df_end_of_page['Col4'][i],cf_header_format_5)
            row_write_val = row_write_val + 1
    column_width_list = [
        [71, 0, 0, Cash_Flow_1] ## cashflow
        ,[20, 1, 4, Cash_Flow_1] ## cashflow
    ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass
    Cash_Flow_1.set_portrait()
    Cash_Flow_1.set_margins(.5,.5,.5,.5)
    Cash_Flow_1.repeat_rows(0, 6)
    Cash_Flow_1.print_area(0,0, row_write_val,3)
    Cash_Flow_1.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    Cash_Flow_1.fit_to_pages(1, total_pages)
    return cash_flow_1_df_end_of_page
# ----------------------------------------------------------------- new stuff twelve_month_actual_budget_v2
def twelve_month_actual_budget_v2(workbook, df, worksheet):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    prop_name = df.columns[0]
    try:
        prop_name = rename_prop(prop_name)
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
                           , df.columns[9]: 'Col10'
                           , df.columns[10]: 'Col11'
                           , df.columns[11]: 'Col12'
                           , df.columns[12]: 'Col13'
                           , df.columns[13]: 'Col14'
                           , df.columns[14]: 'Col15'
                           , df.columns[15]: 'Col16'
                           , df.columns[16]: 'Col17' 
                           })
    header_2 = df['Col1'][0]
    header_3 = df['Col1'][1]
    header_4 = df['Col1'][2]
    header_list_1 = [None #A
                     , None #B
                     , None #C
                     , None #D
                     , None #E
                     , None #F
                     , None #G
                     , None #H
                     , None #I
                     , None #J
                     , None #K
                     , None #L
                     , None #M
                     , 'Total' #N
                     , None #O
                     , None #P
                     , None #Q
                    ]
    header_list_2 = [None #A
                     , df['Col2'][4] #B
                     , df['Col3'][4] #C
                     , df['Col4'][4] #D
                     , df['Col5'][4] #E
                     , df['Col6'][4] #F
                     , df['Col7'][4] #G
                     , df['Col8'][4] #H
                     , df['Col9'][4] #I
                     , df['Col10'][4] #J
                     , df['Col11'][4] #K
                     , df['Col12'][4] #L
                     , df['Col13'][4] #M
                     , df['Col14'][4] #N
                     , df['Col15'][4] #O
                     , df['Col16'][4] #P
                     , df['Col17'][4] #Q
                    ]
    header_list_3 = [None #A
                     , df['Col2'][5] #B
                     , df['Col3'][5] #C
                     , df['Col4'][5] #D
                     , df['Col5'][5] #E
                     , df['Col6'][5] #F
                     , df['Col7'][5] #G
                     , df['Col8'][5] #H
                     , df['Col9'][5] #I
                     , df['Col10'][5] #J
                     , df['Col11'][5] #K
                     , df['Col12'][5] #L
                     , df['Col13'][5] #M
                     , df['Col14'][5] #N
                     , df['Col15'][5] #O
                     , df['Col16'][5] #P
                     , df['Col17'][5] #Q
                    ]
    #print(header_list_2)
    df=df.dropna(subset=['Col1']).reset_index(drop=True)
    df['Col1'] = df.apply(lambda x: clean_text_col_1(x['Col1']), axis=1)
    df['Nan_Var_Check'] = df.apply(lambda x: flag_zero_vals(x['Col16'], x['Col17']), axis=1)
    df['Total_Check'] = df.apply(lambda x: flag_total_rows(x['Col1']), axis=1)
    df['Header_Check'] = df.apply(lambda x: flag_header_rows(x['Col1']), axis=1)
    # wirte excel
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A1:Q1", prop_name, header_format_1)
    worksheet.merge_range("A2:Q2", header_2, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A3:Q3", header_3, header_format_2)
    worksheet.merge_range("A4:Q4", header_4, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    for col in range(17):
        try:
            worksheet.write(4, col, header_list_1[col], header_format_3)
        except:
            worksheet.write_blank(4, col, None, header_format_3)
        try:
            worksheet.write(5, col, header_list_2[col], header_format_3)
        except:
            worksheet.write_blank(5, col, None, header_format_3)
        try:
            worksheet.write(6, col, header_list_3[col], header_format_3)
        except:
            worksheet.write_blank(6, col, None, header_format_3)
    worksheet.merge_range(7, 0, 7, 17, '', header_format_2)
    worksheet.set_row(7,7.5)
    row_write_val = 8
    row_val_format_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    row_val_format_sub_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    row_val_format_sub_item_num_red = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    row_val_format_sub_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    row_val_format_total_item_num = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                     })
    row_val_format_total_item_num_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_header_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_item_num_2_grand = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent_2_grand = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;(#,##0.00%)'
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                     })
    row_val_format_total_header_2_grand = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                     })
    row_val_format_sub_item_num_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    row_val_format_sub_item_percent_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                     })
    row_val_format_total_item_num_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_header_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border_color':black_color
                                     })
    row_val_format_total_item_num_2_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent_2_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_header_2_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    row_val_format_total_item_num_2_grand_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                    })
    row_val_format_total_item_percent_2_grand_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'#,##0.00%;[Red](#,##0.00%)'
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                     })
    row_val_format_total_header_2_grand_var = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);[Red]_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                     })
    for i in range(3, df.shape[0]):
        new_row_needed = 0
        try:
            val = str(df['Col1'][i])
        except:
            val = ''
        if val == 'AL OF ALL':
            pass
        elif val == 'TOTAL OF ALL':
            pass
        elif df['Header_Check'][i] == 1:
            try:
                next_total_val = df['Total_Check'][i-1]
            except:
                next_total_val = 0
            try:
                next_header_val = df['Header_Check'][i+1]
                next_nan_val = df['Nan_Var_Check'][i+1]
                next_next_total_val = df['Total_Check'][i+1]
            except:
                next_header_val = 0
                next_next_total_val = 0
                next_nan_val = 0
            if df['Col1'][i] in ['OPERATING INCOME', 'OPERATING EXPENSES', 'RECOVERABLE', 'NON-RECOVERABLE']:
                worksheet.write_string(row_write_val, 0, df['Col1'][i], row_val_format_header)
                row_write_val = row_write_val + 1
                worksheet.merge_range(row_write_val, 0, row_write_val, 17, '', header_format_2)
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            else:
                if df['Total_Check'][i] == 1:
                    try:
                        next_row_col_1 = df['Col1'][i + 1]
                    except:
                        next_row_col_1 = ''
                    if df['Col1'][i] == 'TOTAL NET INCOME':
                        worksheet.write_string(row_write_val, 0, df['Col1'][i],row_val_format_total_header_2_grand)
                        worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 4, df['Col5'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 8, df['Col9'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 9, df['Col10'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 10, df['Col11'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 11, df['Col12'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 12, df['Col13'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 13, df['Col14'][i],row_val_format_total_item_num_2_grand)
                        worksheet.write_number(row_write_val, 14, df['Col15'][i],row_val_format_total_item_num_2_grand)
                        try:
                            worksheet.write_number(row_write_val, 15, df['Col16'][i],row_val_format_total_item_num_2_grand_var)
                        except:
                            worksheet.write_number(row_write_val, 15, 0,row_val_format_total_item_num_2_grand_var)
                        try:
                            worksheet.write_number(row_write_val, 16, df['Col17'][i]/100,row_val_format_total_item_percent_2_grand_var)
                        except:
                            if df['Col15'][i] == 0 and df['Col16'][i] != 0:
                                worksheet.write_number(row_write_val, 16, -1,row_val_format_total_item_percent_2_grand_var)
                            else:
                                worksheet.write_number(row_write_val, 16, 0,row_val_format_total_item_percent_2_grand_var)
                        row_write_val = row_write_val + 1
                        if next_header_val == 1:
                            if next_next_total_val == 0:
                                new_row_needed = 1
                    elif next_row_col_1 == 'TOTAL NET INCOME':
                        worksheet.write_string(row_write_val, 0, df['Col1'][i],row_val_format_total_header_2)
                        worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 4, df['Col5'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 8, df['Col9'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 9, df['Col10'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 10, df['Col11'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 11, df['Col12'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 12, df['Col13'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 13, df['Col14'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 14, df['Col15'][i],row_val_format_total_item_num_2)
                        try:
                            worksheet.write_number(row_write_val, 15, df['Col16'][i],row_val_format_total_item_num_2_var)
                        except:
                            worksheet.write_number(row_write_val, 15, 0,row_val_format_total_item_num_2)
                        try:
                            worksheet.write_number(row_write_val, 16, df['Col17'][i]/100,row_val_format_total_item_percent_2_var)
                        except:
                            if df['Col15'][i] == 0 and df['Col16'][i] != 0:
                                worksheet.write_number(row_write_val, 16, -1,row_val_format_total_item_percent_2_var)
                            else:
                                worksheet.write_number(row_write_val, 16, 0,row_val_format_total_item_percent_2_var)
                        row_write_val = row_write_val + 1
                        if next_header_val == 1:
                            if next_next_total_val == 0:
                                new_row_needed = 1
                    elif next_total_val == 1:
                        worksheet.write_string(row_write_val, 0, df['Col1'][i],row_val_format_total_header_2)
                        worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 4, df['Col5'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 8, df['Col9'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 9, df['Col10'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 10, df['Col11'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 11, df['Col12'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 12, df['Col13'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 13, df['Col14'][i],row_val_format_total_item_num_2)
                        worksheet.write_number(row_write_val, 14, df['Col15'][i],row_val_format_total_item_num_2)
                        try:
                            worksheet.write_number(row_write_val, 15, df['Col16'][i],row_val_format_total_item_num_2_var)
                        except:
                            worksheet.write_number(row_write_val, 15, 0,row_val_format_total_item_num_2_var)
                        try:
                            worksheet.write_number(row_write_val, 16, df['Col17'][i]/100,row_val_format_total_item_percent_2_var)
                        except:
                            if df['Col15'][i] == 0 and df['Col16'][i] != 0:
                                worksheet.write_number(row_write_val, 16, -1,row_val_format_total_item_percent_2_var)
                            else:
                                worksheet.write_number(row_write_val, 16, 0,row_val_format_total_item_percent_2_var)
                        row_write_val = row_write_val + 1
                        if next_header_val == 1:
                            if next_next_total_val == 0:
                                new_row_needed = 1
                    else:
                        if df['Nan_Var_Check'][i] == 0:
                            worksheet.write_string(row_write_val, 0, df['Col1'][i],row_val_format_total_header)
                            worksheet.write_number(row_write_val, 1, df['Col2'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 2, df['Col3'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 3, df['Col4'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 4, df['Col5'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 8, df['Col9'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 9, df['Col10'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 10, df['Col11'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 11, df['Col12'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 12, df['Col13'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 13, df['Col14'][i],row_val_format_total_item_num)
                            worksheet.write_number(row_write_val, 14, df['Col15'][i],row_val_format_total_item_num)
                            try:
                                worksheet.write_number(row_write_val, 15, df['Col16'][i],row_val_format_total_item_num_var)
                            except:
                                worksheet.write_number(row_write_val, 15, 0,row_val_format_total_item_num_var)
                            try:
                                worksheet.write_number(row_write_val, 16, df['Col17'][i]/100,row_val_format_total_item_percent_var)
                            except:
                                if df['Col15'][i] == 0 and df['Col16'][i] != 0:
                                    worksheet.write_number(row_write_val, 16, -1,row_val_format_total_item_percent_var)
                                else:
                                    worksheet.write_number(row_write_val, 16, 0,row_val_format_total_item_percent_var)
                            row_write_val = row_write_val + 1
                            if next_header_val == 1 and next_next_total_val == 0:
                                new_row_needed = 1
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
                worksheet.merge_range(row_write_val, 0, row_write_val, 17, '', header_format_2)
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
                new_row_needed = 0
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
                worksheet.write_number(row_write_val, 4, df['Col5'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 5, df['Col6'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 6, df['Col7'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 8, df['Col9'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 9, df['Col10'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 10, df['Col11'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 11, df['Col12'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 12, df['Col13'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 13, df['Col14'][i],row_val_format_sub_item_num)
                worksheet.write_number(row_write_val, 14, df['Col15'][i],row_val_format_sub_item_num)
                try:
                    worksheet.write_number(row_write_val, 15, df['Col16'][i],row_val_format_sub_item_num_red)
                except:
                    worksheet.write_number(row_write_val, 15, 0,row_val_format_sub_item_num)
                try:
                    worksheet.write_number(row_write_val, 16, df['Col17'][i]/100,row_val_format_sub_item_percent_var)
                except:
                    if df['Col15'][i] == 0 and df['Col16'][i] != 0:
                        worksheet.write_number(row_write_val, 16, -1,row_val_format_sub_item_percent_var)
                    else:
                        worksheet.write_number(row_write_val, 16, 0,row_val_format_sub_item_percent_var)
                row_write_val = row_write_val + 1
        # add a row or not
        if new_row_needed == 1:
            worksheet.merge_range(row_write_val, 0, row_write_val, 17, '', header_format_2)
            worksheet.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
        else:
            pass
    column_width_list = [
        [42.3,0,0, worksheet]
        ,[13.91,1,16, worksheet]
    ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass
    worksheet.set_landscape()
    worksheet.set_margins(.5,.5,.5,.5)
    worksheet.repeat_rows(0, 7)
    worksheet.print_area(0,0, row_write_val, 16)
    worksheet.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    worksheet.fit_to_pages(1, total_pages)
    return df
# ----------------------------------------------------------------- new stuff mnth_gl_sheet
def mnth_gl_sheet(workbook, df, worksheet): ##mnth_gl_sheet
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    header_1 = df.columns[0]
    try:
        header_1 = rename_prop(header_1)
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
                           , df.columns[9]: 'Col10'
                           , df.columns[10]: 'Col11'
                           })
    
    try:
        df = df.rename(columns={df.columns[11]: 'Col12'
                           , df.columns[12]: 'Col13'
                           , df.columns[13]: 'Col14'
                           , df.columns[14]: 'Col15'
                           })
        df = df.drop(['Col12', 'Col13','Col14', 'Col15'], axis=1)
    except:
        pass
    header_2 = df['Col1'][0]
    header_3 = df['Col1'][1]
    header_4 = df['Col1'][2]
    header_5 = df['Col1'][3]
    header_list_1 = [df['Col1'][4] #A
                     , df['Col2'][4] #B
                     , df['Col3'][4] #C
                     , df['Col4'][4] #D
                     , df['Col5'][4] #E
                     , df['Col6'][4] #F
                     , df['Col7'][4] #G
                     , df['Col8'][4] #H
                     , df['Col9'][4] #I
                     , df['Col10'][4] #J
                     , df['Col11'][4] #K
                    ]
    df=df.dropna(how='all').reset_index(drop=True)
    df['color_col'] = df.apply(lambda x: flag_box(x['Col5'], x['Col9']), axis=1) # there will be some headers that check this, so headers is first in if elif logic
    #df['data'] = df.apply(lambda x: flag_box(x['Col3'], x['Col1']), axis=1)
    df['total'] = df.apply(lambda x: flag_total_rows_3(x['Col5']), axis=1)
    # wirte excel
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A1:K1", header_2, header_format_1)
    worksheet.merge_range("A2:K2", header_1, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A3:K3", header_3, header_format_2)
    worksheet.merge_range("A4:K4", header_4, header_format_2)
    worksheet.merge_range("A5:K5", header_5, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    for col in range(11):
        try:
            str(header_list_1[col])
            worksheet.write(5, col, header_list_1[col], header_format_3)
        except:
            worksheet.write_blank(5, col, None, header_format_3)
    worksheet.merge_range(6, 0, 6, 10, '', header_format_2)
    worksheet.set_row(6,7.5)
    row_write_val = 7
    header_format_body = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    header_format_body_wrap = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'text_wrap':True
                                     })
    header_format_body_num = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    header_format_date_1 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'mm-yyyy'
                                     })
    header_format_date_2 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':14
                                     })
    data_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    })
    data_format_1_wrap = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'text_wrap':True
                                    })
    data_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                        })
    data_format_3 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':14
                                        })
    data_format_4 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'mm-yyyy'
                                        })
    # 
    total_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                        })
    total_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                        })
    grand_total_format_bottom = workbook.add_format({'font_color': black_color
                                    , 'border_color': black_color
                                    , 'bottom':1
                                        }) 
    grand_total_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'bottom':1
                                    , 'border_color': black_color
                                        })
    grand_total_format_22 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'top':1
                                    , 'border_color': black_color
                                        })
    grand_total_format_1_wrap = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'bottom':1
                                    , 'border_color': black_color
                                    , 'text_wrap':True
                                        })
    grand_total_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'border':6
                                    , 'left':0
                                    , 'top':0
                                    , 'right':0
                                    #, 'bottom':3
                                    , 'border_color':black_color
                                        })
    grand_total_format_3 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'border':6
                                    , 'left':0
                                    , 'top':0
                                    , 'right':0
                                    #, 'bottom':3
                                    , 'border_color':black_color
                                        })
    for i in range(5, df.shape[0]): #flag_total_rows_2
        if i + 1 == df.shape[0]:
            try:
                row_write_val = row_write_val + 1
                for row_loop in range(11):
                    if row_loop == 7:
                        worksheet.write(row_write_val, 7, df['Col8'][i],grand_total_format_22)
                    elif row_loop == 8:
                        worksheet.write(row_write_val, 8, df['Col9'][i],grand_total_format_22)
                    else:
                        worksheet.write_blank(row_write_val, row_loop, '',grand_total_format_22)#grand_total_format_22
                row_write_val = row_write_val + 1
            except:
                pass
        elif df['color_col'][i] == 1:
            worksheet.write_string(row_write_val, 0, df['Col1'][i],header_format_body)
            worksheet.write_blank(row_write_val, 1, df['Col2'][i],header_format_body)
            worksheet.write_blank(row_write_val, 2, None,header_format_body)
            worksheet.write_blank(row_write_val, 3, None,header_format_body)
            worksheet.write_string(row_write_val, 4, df['Col5'][i],header_format_body_wrap)
            worksheet.write_blank(row_write_val, 5, None,header_format_body)
            worksheet.write_blank(row_write_val, 6, None,header_format_body)
            worksheet.write_blank(row_write_val, 7, None,header_format_body)
            worksheet.write_blank(row_write_val, 8, None,header_format_body)
            try:
                val_col_10_color = float(df['Col10'][i])
            except:
                val_col_10_color = ''
                for i in df['Col10'][i]:
                    if i == ' ':
                        pass
                    elif i == ',':
                        pass
                    else:
                        val_col_10_color = val_col_10_color+i
                val_col_10_color = float(val_col_10_color)
            worksheet.write_number(row_write_val, 9, val_col_10_color,header_format_body_num)
            worksheet.write_string(row_write_val, 10, 'Beginning Balance',header_format_body)
            row_write_val += 1
        elif df['total'][i] == 1:
            worksheet.write_blank(row_write_val, 0, None,header_format_body)
            worksheet.write_blank(row_write_val, 1, None,header_format_body)
            worksheet.write_blank(row_write_val, 2, None,header_format_body)
            worksheet.write_blank(row_write_val, 3, None,header_format_body)
            worksheet.write_string(row_write_val, 4, df['Col5'][i],header_format_body_wrap)
            worksheet.write_blank(row_write_val, 5, None,header_format_body)
            worksheet.write_blank(row_write_val, 6, None,header_format_body)
            try:
                val_col_8 = float(df['Col8'][i])
                val_col_9 = float(df['Col9'][i])
                worksheet.write_number(row_write_val, 7, val_col_8,header_format_body_num)
                worksheet.write_number(row_write_val, 8, val_col_9,header_format_body_num)
            except:
                worksheet.write(row_write_val, 7, df['Col8'][i],header_format_body_num)
                worksheet.write(row_write_val, 8, df['Col9'][i],header_format_body_num)
            try:
                val_col_10_deux = float(df['Col10'][i])
                worksheet.write_number(row_write_val, 9, val_col_10_deux,header_format_body_num)
            except:
                val_col_10_deux = ''
                for i in df['Col10'][i]:
                    if i == ' ':
                        pass
                    elif i == ',':
                        pass
                    else:
                        val_col_10_deux = val_col_10_deux+i
                val_col_10_deux = float(val_col_10_deux)
                worksheet.write(row_write_val, 9, val_col_10_deux,header_format_body_num)
            worksheet.write(row_write_val, 10, 'Ending Balance',header_format_body)
            row_write_val = row_write_val + 1
            worksheet.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
        else:
            try:
                worksheet.write(row_write_val, 0, df['Col1'][i],data_format_1)
                worksheet.write(row_write_val, 1, df['Col2'][i],data_format_1)
                worksheet.write(row_write_val, 2, df['Col3'][i],data_format_3)
                worksheet.write(row_write_val, 3, df['Col4'][i],data_format_4)
                worksheet.write(row_write_val, 4, df['Col5'][i],data_format_1_wrap)
                worksheet.write(row_write_val, 5, df['Col6'][i],data_format_1)
                worksheet.write(row_write_val, 6, df['Col7'][i],data_format_3)
                worksheet.write(row_write_val, 7, df['Col8'][i],data_format_2)
                worksheet.write(row_write_val, 8, df['Col9'][i],data_format_2)
                worksheet.write(row_write_val, 9, df['Col10'][i],data_format_2)
                worksheet.write(row_write_val, 10, df['Col11'][i],data_format_1_wrap)
                row_write_val += 1
            except:
                pass
    column_width_list = [
                [11.5, 0, 0, worksheet]
                ,[25, 1, 1, worksheet]
                ,[10.8, 2, 2, worksheet]
                ,[11.3, 3, 3, worksheet]
                ,[32, 4, 4, worksheet]
                ,[9.3, 5, 5, worksheet]
                ,[23.3, 6, 6, worksheet]
                ,[14, 7, 8, worksheet]
                ,[14.4, 9, 9, worksheet]
                ,[32, 10, 10, worksheet]
    ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass
    worksheet.set_landscape()
    worksheet.set_margins(.5,.5,.5,.5)
    worksheet.repeat_rows(0, 5)
    worksheet.print_area(0,0, row_write_val - 1, 10)
    worksheet.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    worksheet.fit_to_pages(1, total_pages)
    return df
# ----------------------------------------------------------------- new stuff ten_sched_1_v2
def ten_sched_1_v2(workbook, df, worksheet):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    header_1 = df.columns[0]
    df = df.rename(columns={df.columns[0]: 'Col1'
                           , df.columns[1]: 'Col2'
                           , df.columns[2]: 'Col3'
                           , df.columns[3]: 'Col4'
                           , df.columns[4]: 'Col5'
                           , df.columns[5]: 'Col6'
                           , df.columns[6]: 'Col7'
                           , df.columns[7]: 'Col8'
                           , df.columns[8]: 'Col9'
                           , df.columns[9]: 'Col10'
                           , df.columns[10]: 'Col11'
                           , df.columns[11]: 'Col12'
                           , df.columns[12]: 'Col13'
                           , df.columns[13]: 'Col14'
                           , df.columns[14]: 'Col15'
                           , df.columns[15]: 'Col16'
                           , df.columns[16]: 'Col17' 
                           })
    header_2 = df['Col1'][0]
    header_list_1 = [df['Col1'][1]
                     , df['Col2'][1] #B
                     , df['Col3'][1] #C
                     , df['Col4'][1] #D
                     , df['Col5'][1] #E
                     , df['Col6'][1] #F
                     , df['Col7'][1] #G
                     , df['Col8'][1] #H
                     , df['Col9'][1] #I
                     , df['Col10'][1] #J
                     , df['Col11'][1] #K
                     , df['Col12'][1] #L
                     , df['Col13'][1] #M
                     , df['Col14'][1] #N
                     , df['Col15'][1] #O
                     , df['Col16'][1] #P
                     , df['Col17'][1] #Q
                    ]
    header_list_2 = [None #A
                     , df['Col2'][2] #B
                     , df['Col3'][2] #C
                     , df['Col4'][2] #D
                     , df['Col5'][2] #E
                     , df['Col6'][2] #F
                     , df['Col7'][2] #G
                     , df['Col8'][2] #H
                     , df['Col9'][2] #I
                     , df['Col10'][2] #J
                     , df['Col11'][2] #K
                     , df['Col12'][2] #L
                     , df['Col13'][2] #M
                     , df['Col14'][2] #N
                     , df['Col15'][2] #O
                     , df['Col16'][2] #P
                     , df['Col17'][2] #Q
                    ]
    header_list_3 = [None #A
                     , df['Col2'][3] #B
                     , df['Col3'][3] #C
                     , df['Col4'][3] #D
                     , df['Col5'][3] #E
                     , df['Col6'][3] #F
                     , df['Col7'][3] #G
                     , df['Col8'][3] #H
                     , df['Col9'][3] #I
                     , df['Col10'][3] #J
                     , df['Col11'][3] #K
                     , df['Col12'][3] #L
                     , df['Col13'][3] #M
                     , df['Col14'][3] #N
                     , df['Col15'][3] #O
                     , df['Col16'][3] #P
                     , df['Col17'][3] #Q
                    ]
    df=df.dropna(subset=['Col3']).reset_index(drop=True)
    df['rent_step_base'] = df.apply(lambda x: flag_box(x['Col3'], x['Col2']), axis=1)
    df['rent_step_header'] = df.apply(lambda x: flag_box(x['Col2'], x['Col17']), axis=1)
    df['Header_Check'] = df.apply(lambda x: check_header_ten_sch(x['Col17']), axis=1)
    df['BOX'] = df.apply(lambda x: flag_box(x['Col1'], x['Col9']), axis=1)
    # wirte excel
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A1:Q1", header_1, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A2:Q2", header_2, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    for col in range(17):
        try:
            str(header_list_1[col])
            worksheet.write(2, col, header_list_1[col], header_format_3)
        except:
            worksheet.write_blank(2, col, None, header_format_3)
        try:
            str(header_list_2[col])
            worksheet.write(3, col, header_list_2[col], header_format_3)
        except:
            worksheet.write_blank(3, col, None, header_format_3)
        try:
            str(header_list_3[col])
            worksheet.write(4, col, header_list_3[col], header_format_3)
        except:
            worksheet.write_blank(4, col, None, header_format_3)
    worksheet.merge_range(5, 0, 5, 17, '', header_format_2)
    worksheet.set_row(5,7.5)
    row_write_val = 6
    row_val_format_header = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_header_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'valign':'vcenter'
                                     })
    row_val_format_header_date = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    , 'num_format':14
                                    , 'align':'center'
                                    , 'valign':'vcenter'
                                     })
    rent_step_header_1 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    , 'valign':'vcenter'
                                     })
    rent_step_header_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    , 'valign':'vcenter'
                                     })
    rent_step_header_3 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    , 'valign':'vcenter'
                                     })
    rent_step_base_1 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    , 'valign':'vcenter'
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    rent_step_base_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    , 'valign':'vcenter'
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    rent_step_base_date = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    , 'valign':'vcenter'
                                    ,'num_format':14
                                    , 'align':'center'
                                     })
    box_title_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':14
                                    , 'align':'center'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    box_title_2 = workbook.add_format({'font_color': black_color
                                    , 'bold':True #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':14
                                    , 'align':'center'
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    box_base_1 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    box_base_3 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    box_base_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':10
                                     })
    box_total_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    box_total_2 = workbook.add_format({'font_color': black_color
                                    , 'bold':True #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':10
                                    , 'top':1
                                    , 'border_color':black_color
                                     })
    format_vacant = workbook.add_format({'font_color': black_color
                                    , 'bold':True #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    format_vacant_2 = workbook.add_format({'font_color': black_color
                                    , 'bold':False #, 'bg_color':black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                    ,'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                     })
    row_write_val = 6
    vacant_counter = 0
    for i in range(1, df.shape[0]):
        if df['Header_Check'][i] == 1: # row_val_format_header
            worksheet.write_string(row_write_val, 0, df['Col1'][i],row_val_format_header)
            worksheet.write_string(row_write_val, 1, df['Col2'][i],row_val_format_header)
            worksheet.write_string(row_write_val, 2, df['Col3'][i],row_val_format_header)
            worksheet.write_string(row_write_val, 3, df['Col4'][i],row_val_format_header)
            worksheet.write_number(row_write_val, 4, df['Col5'][i],row_val_format_header)
            try:
                worksheet.write_datetime(row_write_val, 5, df['Col6'][i],row_val_format_header_date)
            except:
                #print(df['Col6'][i])
                worksheet.write(row_write_val, 5, df['Col6'][i],row_val_format_header_date)
            try:
                worksheet.write_datetime(row_write_val, 6, df['Col7'][i],row_val_format_header_date)
            except:
                #print(df['Col7'][i])
                try:
                    worksheet.write(row_write_val, 6, df['Col7'][i],row_val_format_header_date)
                except:
                    worksheet.write_blank(row_write_val, 6, None, row_val_format_header_date)
            try:
                worksheet.write_number(row_write_val, 7, df['Col8'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 7, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 8, df['Col9'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 8, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 9, df['Col10'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 8, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 10, df['Col11'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 10, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 11, df['Col12'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 11, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 12, df['Col13'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 12, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 13, df['Col14'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 13, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 14, df['Col15'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 14, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 15, df['Col16'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 15, None, row_val_format_header_2)
            try:
                worksheet.write_number(row_write_val, 16, df['Col17'][i],row_val_format_header_2)
            except:
                worksheet.write_blank(row_write_val, 16, None, row_val_format_header_2)
            row_write_val = row_write_val + 1
            worksheet.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
        elif df['Col3'][i] == 'VACANT':
            if vacant_counter == 0:
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            for j in range(17):
                if j == 0:
                    worksheet.write(row_write_val, 0, df['Col1'][i],format_vacant)
                elif j == 1:
                    worksheet.write(row_write_val, 1, df['Col2'][i],format_vacant)
                elif j == 2:
                    worksheet.write(row_write_val, 2, df['Col3'][i],format_vacant_2)
                elif j == 4:
                    try:
                        val_44 = float(df['Col5'][i])
                    except:
                        val_44 = df['Col5'][i]
                    worksheet.write(row_write_val, 4, val_44,format_vacant_2)
                else:
                    worksheet.write_blank(row_write_val, j, None, format_vacant)
            row_write_val = row_write_val + 1
            worksheet.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
            vacant_counter += 1
        elif df['BOX'][i] == 1:
            #print('hey')
            worksheet.set_row(row_write_val-1,14.5)
            if df['Col1'][i] == 'Occupancy Summary':
                worksheet.write_string(row_write_val, 0, 'Occupancy Summary',rent_step_header_3)
                worksheet.write_blank(row_write_val, 1, None, rent_step_header_1)
                worksheet.write_string(row_write_val, 2, 'Area',box_title_2)
                worksheet.write_string(row_write_val, 3, 'Percentage',box_title_2)
                row_write_val = row_write_val + 1
            elif df['Col1'][i] == 'Occupied Area':
                worksheet.write_string(row_write_val, 0, 'Occupied Area',box_base_3)
                worksheet.write_blank(row_write_val, 1, None, box_base_1)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],box_base_1)
                try:
                    worksheet.write_number(row_write_val, 3, df['Col4'][i]/100,box_base_2)
                except:
                    worksheet.write_number(row_write_val, 3, 0,box_base_2)
                row_write_val = row_write_val + 1
            elif df['Col1'][i] == 'Vacant Area':
                worksheet.write_string(row_write_val, 0, 'Vacant Area',box_base_3)
                worksheet.write_blank(row_write_val, 1, None, box_base_1)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],box_base_1)
                try:
                    worksheet.write_number(row_write_val, 3, df['Col4'][i]/100,box_base_2)
                except:
                    worksheet.write_number(row_write_val, 3, 0,box_base_2)
                row_write_val = row_write_val + 1
            elif df['Col1'][i] == 'Total': #box_total_1
                worksheet.write_string(row_write_val, 0, 'Total',box_total_1)
                worksheet.write_blank(row_write_val, 1, None, box_total_1)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],box_total_1)
                try:
                    worksheet.write_number(row_write_val, 3, df['Col4'][i]/100,box_total_2)
                except:
                    worksheet.write_number(row_write_val, 3, 0,box_total_2)
                row_write_val = row_write_val + 1
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            elif df['Col1'][i] == 'Total Occupied Area':
                worksheet.write_string(row_write_val, 0, 'Occupied Area',box_base_3)
                worksheet.write_blank(row_write_val, 1, None, box_base_1)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],box_base_1)
                try:
                    worksheet.write_number(row_write_val, 3, df['Col4'][i]/100,box_base_2)
                except:
                    worksheet.write_number(row_write_val, 3, 0,box_base_2)
                row_write_val = row_write_val + 1
            elif df['Col1'][i] == 'Grand Total':
                worksheet.write_string(row_write_val, 0, 'Total',box_total_1)
                worksheet.write_blank(row_write_val, 1, None, box_total_1)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],box_total_1)
                try:
                    worksheet.write_number(row_write_val, 3, df['Col4'][i]/100,box_total_2)
                except:
                    worksheet.write_number(row_write_val, 3, 0,box_total_2)
                row_write_val = row_write_val + 1
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
            elif df['Col1'][i] == 'Total Vacant Area':
                worksheet.write_string(row_write_val, 0, 'Vacant Area',box_base_3)
                worksheet.write_blank(row_write_val, 1, None, box_base_1)
                worksheet.write_number(row_write_val, 2, df['Col3'][i],box_base_1)
                try:
                    worksheet.write_number(row_write_val, 3, df['Col4'][i]/100,box_base_2)
                except:
                    worksheet.write_number(row_write_val, 3, 0,box_base_2)
                row_write_val = row_write_val + 1
        elif df['rent_step_header'][i] == 1:
            worksheet.write_string(row_write_val, 1, df['Col2'][i],rent_step_header_1)
            worksheet.write_string(row_write_val, 2, df['Col3'][i],rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 3, df['Col4'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 3, df['Col4'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 3, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 4, df['Col5'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 4, df['Col5'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 4, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 5, df['Col6'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 5, df['Col6'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 5, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 6, df['Col7'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 6, df['Col7'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 6, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 7, df['Col8'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 7, df['Col8'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 7, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 8, df['Col9'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 8, df['Col9'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 8, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 9, df['Col10'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 9, df['Col10'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 9, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 10, df['Col11'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 10, df['Col11'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 10, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 11, df['Col12'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 11, df['Col12'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 11, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 12, df['Col13'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 12, df['Col13'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 12, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 13, df['Col14'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 13, df['Col14'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 13, None,rent_step_header_2)
            try:
                worksheet.write_string(row_write_val, 14, df['Col15'][i],rent_step_header_2)
            except:
                try:
                    worksheet.write(row_write_val, 14, df['Col15'][i],rent_step_header_2)
                except:
                    worksheet.write_blank(row_write_val, 14, None,rent_step_header_2)
            row_write_val = row_write_val + 1
        elif df['rent_step_base'][i] == 1:
            next_rent_base = 0
            try:
                if df['rent_step_base'][i+1] != 1:
                    next_rent_base = 1
            except:
                pass
            try:
                if df['BOX'][i+1] == 1:
                    next_rent_base = 1
            except:
                pass
            worksheet.write_string(row_write_val, 2, df['Col3'][i],rent_step_base_1)
            worksheet.write_string(row_write_val, 3, df['Col4'][i],rent_step_base_1)
            worksheet.write_string(row_write_val, 4, df['Col5'][i],rent_step_base_1)
            worksheet.write_string(row_write_val, 5, df['Col6'][i],rent_step_base_1)
            worksheet.write_number(row_write_val, 6, df['Col7'][i],rent_step_base_2)
            worksheet.write_datetime(row_write_val, 7, df['Col8'][i],rent_step_base_date)
            try:
                worksheet.write_datetime(row_write_val, 8, df['Col9'][i],rent_step_base_date)
            except:
                worksheet.write_blank(row_write_val, 8, None, rent_step_base_date)
            try:
                worksheet.write_number(row_write_val, 9, df['Col10'][i],rent_step_base_2)
            except:
                worksheet.write_blank(row_write_val, 9, None, rent_step_base_2)
            try:
                worksheet.write_number(row_write_val, 10, df['Col11'][i],rent_step_base_2)
            except:
                worksheet.write_blank(row_write_val, 10, None, rent_step_base_2)
            try:
                worksheet.write_number(row_write_val, 11, df['Col12'][i],rent_step_base_2)
            except:
                worksheet.write_blank(row_write_val, 11, None, rent_step_base_2)
            try:
                worksheet.write_number(row_write_val, 12, df['Col13'][i],rent_step_base_2)
            except:
                worksheet.write_blank(row_write_val, 12, None, rent_step_base_2)
            try:
                worksheet.write_number(row_write_val, 13, df['Col14'][i],rent_step_base_2)
            except:
                worksheet.write_blank(row_write_val, 13, None, rent_step_base_2)
            try:
                worksheet.write_number(row_write_val, 14, df['Col15'][i],rent_step_base_2)
            except:
                worksheet.write_blank(row_write_val, 14, None, rent_step_base_2)
            row_write_val = row_write_val + 1
            if next_rent_base == 1:
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
        else:
            pass
    column_width_list = [
        [21, 0, 0, worksheet]        
        ,[12.7, 1, 1, worksheet]
        ,[18, 2, 2, worksheet]
        ,[18, 3, 3, worksheet]
        ,[12.7, 4, 15, worksheet]
        ,[17.55, 16, 16, worksheet]
    ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass ## Tenancy Schedule
    worksheet.set_landscape()
    worksheet.set_margins(.5,.5,.5,.5)
    worksheet.repeat_rows(0, 4)
    worksheet.print_area(0,0, row_write_val - 1, 16)
    worksheet.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    worksheet.fit_to_pages(1, total_pages)
    return df
# -----------------------------------------------------------------
def payment_register_sheet_2(workbook, df, worksheet):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    header_1 = df.columns[0]
    df = df.rename(columns={df.columns[0]: 'Col1'
                           , df.columns[1]: 'Col2'
                           , df.columns[2]: 'Col3'
                           , df.columns[3]: 'Col4'
                           , df.columns[4]: 'Col5'
                           , df.columns[5]: 'Col6'
                           , df.columns[6]: 'Col7'
                           , df.columns[7]: 'Col8'
                           , df.columns[8]: 'Col9'
                           , df.columns[9]: 'Col10'
                           , df.columns[10]: 'Col11'
                           , df.columns[11]: 'Col12'
                           , df.columns[12]: 'Col13'
                           , df.columns[13]: 'Col14'
                           })
    header_2 = df['Col1'][0]
    header_3 = df['Col1'][1]
    header_list_1 = [df['Col1'][1]
                     , df['Col2'][1] #B
                     , df['Col3'][1] #C
                     , df['Col4'][1] #D
                     , df['Col5'][1] #E
                     , df['Col6'][1] #F
                     , df['Col7'][1] #G
                     , df['Col8'][1] #H
                     , df['Col9'][1] #I
                     , df['Col10'][1] #J
                     , df['Col11'][1] #K
                     , df['Col12'][1] #L
                     , df['Col13'][1] #M
                     , df['Col14'][1] #N
                    ]
    header_list_2 = [None #A
                     , df['Col2'][2] #B
                     , df['Col3'][2] #C
                     , df['Col4'][2] #D
                     , df['Col5'][2] #E
                     , df['Col6'][2] #F
                     , df['Col7'][2] #G
                     , df['Col8'][2] #H
                     , df['Col9'][2] #I
                     , df['Col10'][2] #J
                     , df['Col11'][2] #K
                     , df['Col12'][2] #L
                     , df['Col13'][2] #M
                     , df['Col14'][2] #N
                    ]
    df=df.dropna(how='all').reset_index(drop=True)
    df['color_col'] = df.apply(lambda x: flag_box(x['Col1'], x['Col11']), axis=1) # there will be some headers that check this, so headers is first in if elif logic
    df['data'] = df.apply(lambda x: flag_box(x['Col9'], x['Col8']), axis=1)
    df['total'] = df.apply(lambda x: flag_total_rows_2(x['Col1']), axis=1)
    # wirte excel
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A1:N1", header_1, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A2:N2", header_2, header_format_2)
    worksheet.merge_range("A3:N3", header_3, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    for col in range(14):
        try:
            str(header_list_2[col])
            worksheet.write(3, col, header_list_2[col], header_format_3)
        except:
            worksheet.write_blank(3, col, None, header_format_3)
    worksheet.merge_range(4, 0, 4, 14, '', header_format_2)
    worksheet.set_row(4,7.5)
    row_write_val = 5
    header_format_body = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'text_wrap':True
                                     })
    header_format_date_1 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'mm-yyyy'
                                    , 'text_wrap':True
                                     })
    header_format_date_2 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':14
                                    , 'align':'center'
                                    , 'text_wrap':True
                                     })
    data_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'text_wrap':True
                                    })
    data_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'text_wrap':True
                                        })
    total_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'border':6
                                    , 'left':0
                                    , 'right':0
                                    ,'top':0
                                    , 'text_wrap':True
                                        })
    total_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'border':6
                                    , 'left':0
                                    , 'right':0
                                    ,'top':0
                                    , 'text_wrap':True
                                        })
    total_format_1_gray = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':'#EEEEEE'
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'text_wrap':True
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color': black_color
                                        })
    total_format_2_gray = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':'#EEEEEE'
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'text_wrap':True
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color': black_color
                                        })
    grand_total_base = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':False
                                    , 'bottom': 1
                                    , 'text_wrap':True
                                        })
    
    for i in range(4, df.shape[0]): #flag_total_rows_2
        if df['color_col'][i] == 1:
            worksheet.write_string(row_write_val, 0, df['Col1'][i],header_format_body)
            worksheet.write_string(row_write_val, 1, df['Col2'][i],header_format_body)
            worksheet.write_string(row_write_val, 2, df['Col3'][i],header_format_body)
            worksheet.write_string(row_write_val, 3, df['Col4'][i],header_format_body)
            worksheet.write_string(row_write_val, 4, df['Col5'][i],header_format_body)
            worksheet.write_datetime(row_write_val, 5, df['Col6'][i],header_format_date_2)
            worksheet.write_datetime(row_write_val, 6, df['Col7'][i],header_format_date_1)
            worksheet.write_string(row_write_val, 7, df['Col8'][i],header_format_body)
            worksheet.write_blank(row_write_val, 8, None,header_format_body)
            worksheet.write_blank(row_write_val, 9, None,header_format_body)
            worksheet.write_blank(row_write_val, 10, None,header_format_body)
            worksheet.write_blank(row_write_val, 11, None,header_format_body)
            worksheet.write_blank(row_write_val, 12, None,header_format_body)
            worksheet.write_blank(row_write_val, 13, None,header_format_body)
            row_write_val += 1
        elif df['total'][i] == 1:
            if df['Col1'][i] == 'Grand Total ':
                worksheet.merge_range(row_write_val-1, 0, row_write_val-1, 13, '', grand_total_base)
                worksheet.write_string(row_write_val, 0, df['Col1'][i],total_format_1)
                worksheet.write_blank(row_write_val, 1, None,total_format_1)
                worksheet.write_blank(row_write_val, 2, None,total_format_1)
                worksheet.write_blank(row_write_val, 3, None,total_format_1)
                worksheet.write_blank(row_write_val, 4, None,total_format_1)
                worksheet.write_blank(row_write_val, 5, None,total_format_1)
                worksheet.write_blank(row_write_val, 6, None,total_format_1)
                worksheet.write_blank(row_write_val, 7, None,total_format_1)
                worksheet.write_blank(row_write_val, 8, None,total_format_1)
                worksheet.write_blank(row_write_val, 9, None,total_format_1)
                worksheet.write_number(row_write_val, 10, df['Col11'][i],total_format_2)
                worksheet.write_blank(row_write_val, 11, None,total_format_1)
                worksheet.write_blank(row_write_val, 12, None,total_format_1)
                worksheet.write_blank(row_write_val, 13, None,total_format_1)
                row_write_val += 1
            else:
                worksheet.write_string(row_write_val, 0, df['Col1'][i],total_format_1_gray)
                worksheet.write_blank(row_write_val, 1, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 2, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 3, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 4, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 5, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 6, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 7, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 8, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 9, None,total_format_1_gray)
                worksheet.write_number(row_write_val, 10, df['Col11'][i],total_format_2_gray)
                worksheet.write_blank(row_write_val, 11, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 12, None,total_format_1_gray)
                worksheet.write_blank(row_write_val, 13, None,total_format_1_gray)
                row_write_val += 1
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
        elif df['data'][i] == 1:
            worksheet.write_blank(row_write_val, 0, None,data_format_1)
            worksheet.write_blank(row_write_val, 1, None,data_format_1)
            worksheet.write_blank(row_write_val, 2, None,data_format_1)
            worksheet.write_blank(row_write_val, 3, None,data_format_1)
            worksheet.write_blank(row_write_val, 4, None,data_format_1)
            worksheet.write_blank(row_write_val, 5, None,data_format_1)
            worksheet.write_blank(row_write_val, 6, None,data_format_1)
            worksheet.write_blank(row_write_val, 7, None,data_format_1)
            try:
                worksheet.write_string(row_write_val, 8, df['Col9'][i],data_format_1)
            except:
                try:
                    worksheet.write(row_write_val, 8, df['Col9'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 8, None,data_format_1)
            #----------------
            try:
                worksheet.write_string(row_write_val, 9, df['Col10'][i],data_format_1)
            except:
                try:
                    worksheet.write(row_write_val, 9, df['Col10'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 8, None,data_format_1)
            try:
                loop_val_col_11 = float(df['Col11'][i])
            except:
                loop_val_col_11 = str(df['Col11'][i])
            try:
                worksheet.write_number(row_write_val, 10, loop_val_col_11,data_format_2)
            except:
                worksheet.write(row_write_val, 10, loop_val_col_11,data_format_2)
            try:
                worksheet.write(row_write_val, 11, df['Col12'][i],data_format_1)
            except:
                worksheet.write_blank(row_write_val, 7, None,data_format_1)
            try:
                worksheet.write(row_write_val, 12, df['Col13'][i],data_format_1)
            except:
                worksheet.write_blank(row_write_val, 7, None,data_format_1)
            worksheet.write_string(row_write_val, 13, df['Col14'][i] ,data_format_1)
            row_write_val += 1
        else:
            pass            
    column_width_list = [[14.4, 0, 0, worksheet] ## Payment Register 0
            ,[13, 1, 2, worksheet] ## Payment Register 1
            ,[15.27, 3, 3, worksheet] ## Payment Register 2
            ,[20, 4, 4, worksheet] ## Payment Register 3
            ,[13, 5, 6, worksheet] ## Payment Register 4
            ,[19.36, 7, 7, worksheet] ## Payment Register 5
            ,[10.73, 8, 10, worksheet] ## Payment Register 6
            ,[18.45, 11, 11, worksheet] ## Payment Register 7
            ,[15.27, 12, 12, worksheet] ## Payment Register 8
            ,[49.35, 13, 13, worksheet] ## Payment Register 9
    ]
    val_loop = 0
    for i in column_width_list:
        try:
            if val_loop == 7:
                i[3].set_column(first_col=i[1],last_col=i[2], width=i[0], options={'level':1, 'hidden':True})
            elif val_loop == 8:
                i[3].set_column(first_col=i[1],last_col=i[2], width=i[0], options={'level':1, 'hidden':True})
            elif val_loop == 9:
                i[3].set_column(first_col=i[1],last_col=i[2], width=i[0], options={'collapsed':1})
            else:
                i[3].set_column(i[1],i[2], i[0])
        except:
            #print(val_loop)
            pass
        val_loop += 1
    worksheet.set_landscape()
    worksheet.set_margins(.5,.5,.5,.5)
    worksheet.repeat_rows(0, 3)
    worksheet.print_area(0,0, row_write_val, 13)
    worksheet.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    worksheet.fit_to_pages(1, total_pages)
    return df
# ----------------------------------------------------------------- new stuff aging_detail_2
def aging_detail_2(workbook, df, worksheet):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    header_1 = df.columns[0]
    df = df.rename(columns={df.columns[0]: 'Col1'
                           , df.columns[1]: 'Col2'
                           , df.columns[2]: 'Col3'
                           , df.columns[3]: 'Col4'
                           , df.columns[4]: 'Col5'
                           , df.columns[5]: 'Col6'
                           , df.columns[6]: 'Col7'
                           , df.columns[7]: 'Col8'
                           , df.columns[8]: 'Col9'
                           , df.columns[9]: 'Col10'
                           , df.columns[10]: 'Col11'
                           , df.columns[11]: 'Col12'
                           , df.columns[12]: 'Col13'
                           , df.columns[13]: 'Col14'
                           , df.columns[14]: 'Col15'
                           })
    header_2 = df['Col1'][0]
    header_list_1 = [df['Col1'][1]
                     , df['Col2'][1] #B
                     , df['Col3'][1] #C
                     , df['Col4'][1] #D
                     , df['Col5'][1] #E
                     , df['Col6'][1] #F
                     , df['Col7'][1] #G
                     , df['Col8'][1] #H
                     , df['Col9'][1] #I
                     , df['Col10'][1] #J
                     , df['Col11'][1] #K
                     , df['Col12'][1] #L
                     , df['Col13'][1] #M
                     , df['Col14'][1] #N
                     , df['Col15'][1] #O
                    ]
    header_list_2 = [None #A
                     , df['Col2'][2] #B
                     , df['Col3'][2] #C
                     , df['Col4'][2] #D
                     , df['Col5'][2] #E
                     , df['Col6'][2] #F
                     , df['Col7'][2] #G
                     , df['Col8'][2] #H
                     , df['Col9'][2] #I
                     , df['Col10'][2] #J
                     , df['Col11'][2] #K
                     , df['Col12'][2] #L
                     , df['Col13'][2] #M
                     , df['Col14'][2] #N
                     , df['Col15'][2] #O
                    ]
    df=df.dropna(how='all').reset_index(drop=True)
    df['total'] = df.apply(lambda x: flag_box(x['Col1'], x['Col3']), axis=1) # there will be some headers that check this, so headers is first in if elif logic
    df['subtotal'] = df.apply(lambda x: flag_box(x['Col3'], x['Col4']), axis=1)
    df['header'] = df.apply(lambda x: flag_box(x['Col1'], x['Col9']), axis=1)
    df['data'] = df.apply(lambda x: flag_box(x['Col4'], x['Col2']), axis=1)
    # wirte excel
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A1:N1", header_1, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A2:N2", header_2, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    for col in range(15):
        try:
            str(header_list_1[col])
            worksheet.write(2, col, header_list_1[col], header_format_3)
        except:
            worksheet.write_blank(2, col, None, header_format_3)
        try:
            str(header_list_2[col])
            worksheet.write(3, col, header_list_2[col], header_format_3)
        except:
            worksheet.write_blank(3, col, None, header_format_3)
    worksheet.merge_range(4, 0, 4, 13, '', header_format_2)
    worksheet.set_row(4,7.5)
    row_write_val = 5
    header_format_body = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    data_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    data_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                     })
    data_format_3 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':14
                                    , 'align':'center'
                                     })
    subtotal_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bold':True
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    subtotal_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    })
    total_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bold':True
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'})
    total_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':14
                                    , 'align':'center'
                                    , 'bold':True
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    })
    grand_total_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'bold':True
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    })
    grand_total_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':14
                                    , 'align':'center'
                                    , 'bold':True
                                    , 'border':6
                                    , 'top':0
                                    , 'left':0
                                    , 'right':0
                                    , 'border_color':black_color
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    })
    bottom_line_format = workbook.add_format({'font_color': black_color
                                    , 'bottom':1
                                    , 'border_color':black_color
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    })
    for i in range(3, df.shape[0]):
        if df['header'][i] == 1:
            worksheet.merge_range(row_write_val, 0, row_write_val, 14, df['Col1'][i], header_format_body)
            row_write_val = row_write_val + 1
        elif df['data'][i] == 1:
            worksheet.write_string(row_write_val, 0, df['Col1'][i],data_format_1)
            worksheet.write_string(row_write_val, 2, df['Col3'][i],data_format_1)
            worksheet.write_string(row_write_val, 3, df['Col4'][i],data_format_1)
            worksheet.write_string(row_write_val, 4, df['Col5'][i],data_format_1)
            worksheet.write_string(row_write_val, 5, df['Col6'][i],data_format_1)
            worksheet.write_datetime(row_write_val, 6, df['Col7'][i],data_format_3)
            worksheet.write_string(row_write_val, 7, df['Col8'][i],data_format_3)
            worksheet.write_number(row_write_val, 8, df['Col9'][i],data_format_2)
            worksheet.write_number(row_write_val, 9, df['Col10'][i],data_format_2)
            worksheet.write_number(row_write_val, 10, df['Col11'][i],data_format_2)
            worksheet.write_number(row_write_val, 11, df['Col12'][i],data_format_2)
            worksheet.write_number(row_write_val, 12, df['Col13'][i],data_format_2)
            worksheet.write_number(row_write_val, 13, df['Col14'][i],data_format_2)
            worksheet.write_number(row_write_val, 14, df['Col15'][i],data_format_2)
            row_write_val = row_write_val + 1
        elif df['subtotal'][i] == 1:
            worksheet.write_blank(row_write_val, 0, None, subtotal_format_1)
            worksheet.write_blank(row_write_val, 1, None, subtotal_format_1)
            worksheet.write_string(row_write_val, 2, df['Col3'][i],subtotal_format_1)
            worksheet.write_blank(row_write_val, 3, None, subtotal_format_1)
            worksheet.write_blank(row_write_val, 4, None, subtotal_format_1)
            worksheet.write_blank(row_write_val, 5, None, subtotal_format_1)
            worksheet.write_blank(row_write_val, 6, None, subtotal_format_1)
            worksheet.write_blank(row_write_val, 7, None, subtotal_format_1)
            worksheet.write_blank(row_write_val, 8, None, subtotal_format_1)
            worksheet.write_number(row_write_val, 8, df['Col9'][i],subtotal_format_2)
            worksheet.write_number(row_write_val, 9, df['Col10'][i],subtotal_format_2)
            worksheet.write_number(row_write_val, 10, df['Col11'][i],subtotal_format_2)
            worksheet.write_number(row_write_val, 11, df['Col12'][i],subtotal_format_2)
            worksheet.write_number(row_write_val, 12, df['Col13'][i],subtotal_format_2)
            worksheet.write_number(row_write_val, 13, df['Col14'][i],subtotal_format_2)
            worksheet.write_number(row_write_val, 14, df['Col15'][i],subtotal_format_2)
            row_write_val = row_write_val + 1
            worksheet.set_row(row_write_val,7.5)
            row_write_val = row_write_val + 1
        elif df['total'][i] == 1:
            if df['Col1'][i] == 'Grand Total':
                worksheet.set_row(row_write_val-1,7.5)
                worksheet.merge_range(row_write_val - 1, 0, row_write_val - 1, 14, '', bottom_line_format)
                worksheet.write_string(row_write_val, 0, df['Col1'][i],grand_total_format_1)
                worksheet.write_blank(row_write_val, 1, None, grand_total_format_1)
                worksheet.write_blank(row_write_val, 2, None, grand_total_format_1)
                worksheet.write_blank(row_write_val, 3, None, grand_total_format_1)
                worksheet.write_blank(row_write_val, 4, None, grand_total_format_1)
                worksheet.write_blank(row_write_val, 5, None, grand_total_format_1)
                worksheet.write_blank(row_write_val, 6, None, grand_total_format_1)
                worksheet.write_blank(row_write_val, 7, None, grand_total_format_1)
                worksheet.write_number(row_write_val, 8, float(df['Col9'][i]),grand_total_format_2)
                worksheet.write_number(row_write_val, 9, df['Col10'][i],grand_total_format_2)
                worksheet.write_number(row_write_val, 10, df['Col11'][i],grand_total_format_2)
                worksheet.write_number(row_write_val, 11, df['Col12'][i],grand_total_format_2)
                worksheet.write_number(row_write_val, 12, df['Col13'][i],grand_total_format_2)
                worksheet.write_number(row_write_val, 13, df['Col14'][i],grand_total_format_2)
                worksheet.write_number(row_write_val, 14, df['Col15'][i],grand_total_format_2)
                row_write_val = row_write_val + 1
            else:
                worksheet.write_string(row_write_val, 0, df['Col1'][i],total_format_1)
                worksheet.write_blank(row_write_val, 1, None, total_format_1)
                worksheet.write_blank(row_write_val, 2, None, total_format_1)
                worksheet.write_blank(row_write_val, 3, None, total_format_1)
                worksheet.write_blank(row_write_val, 4, None, total_format_1)
                worksheet.write_blank(row_write_val, 5, None, total_format_1)
                worksheet.write_blank(row_write_val, 6, None, total_format_1)
                worksheet.write_blank(row_write_val, 7, None, total_format_1)
                worksheet.write_number(row_write_val, 8, df['Col9'][i],total_format_2)
                worksheet.write_number(row_write_val, 9, df['Col10'][i],total_format_2)
                worksheet.write_number(row_write_val, 10, df['Col11'][i],total_format_2)
                worksheet.write_number(row_write_val, 11, df['Col12'][i],total_format_2)
                worksheet.write_number(row_write_val, 12, df['Col13'][i],total_format_2)
                worksheet.write_number(row_write_val, 13, df['Col14'][i],total_format_2)
                worksheet.write_number(row_write_val, 14, df['Col15'][i],total_format_2)
                row_write_val = row_write_val + 1
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
        else:
            pass
    if df['Col1'][3] == 'No outstanding AR as of period end':
        #worksheet.merge_range(6, 0, 6, 13, df['Col1'][3], data_format_1)
        worksheet.set_row(6,7.5)
        row_write_val = 6
    column_width_list = [
        [10, 0, 0, worksheet] ## AR Detail
        ,[9.8, 1, 1, worksheet] ## AR Detail
        ,[42.2, 2, 2, worksheet] ## AR Detail
        ,[9.8, 3, 3, worksheet] ## AR Detail
        ,[7.8, 4, 4, worksheet] ## AR Detail
        ,[8.8, 5, 5, worksheet] ## AR Detail
        ,[10.3, 6, 6, worksheet] ## AR Detail
        ,[10.3, 7, 7, worksheet] ## AR Detail
        ,[14.36, 8, 8, worksheet] ## AR Detail
        ,[13, 9, 11, worksheet] ## AR Detail
        ,[15.27, 12, 12, worksheet] ## AR Detail
        ,[13, 13, 13, worksheet] ## AR Detail
        ,[15.27, 14, 14, worksheet] ## AR Detail
    ]
    loop_val = 0
    for i in column_width_list:
        try:
            if loop_val == 1:
                i[3].set_column(i[1],i[2], i[0], options={'hidden':True})
            else:
                i[3].set_column(i[1],i[2], i[0])
        except:
            pass
        loop_val += 1
    worksheet.set_landscape()
    worksheet.set_margins(.5,.5,.5,.5)
    worksheet.repeat_rows(0, 3)
    worksheet.print_area(0,0, row_write_val, 14)
    worksheet.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    worksheet.fit_to_pages(1, total_pages)
    return df
# ----------------------------------------------------------------- new stuff JE_REGISTER_SHEET
def JE_REGISTER_SHEET(workbook, df, worksheet):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    header_1 = df.columns[0]
    try:
        header_1 = header_1.split('(', 1)[0]
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
                           , df.columns[9]: 'Col10'
                           , df.columns[10]: 'Col11'
                           , df.columns[11]: 'Col12'
                           , df.columns[12]: 'Col13'
                           , df.columns[13]: 'Col14'
                           })
    header_2 = df['Col1'][0]
    try:
        header_2 = header_2.split('(', 1)[0]
    except:
        pass
    header_3 = df['Col1'][1]
    header_list_1 = [df['Col1'][2] #A
                     , df['Col2'][2] #B
                     , df['Col3'][2] #C
                     , df['Col4'][2] #D
                     , df['Col5'][2] #E
                     , df['Col6'][2] #F
                     , df['Col7'][2] #G
                     , df['Col8'][2] #H
                     , df['Col9'][2] #I
                     , df['Col10'][2] #J
                     , df['Col11'][2] #K
                     , df['Col12'][2] #J
                     , df['Col13'][2] #J
                     , df['Col14'][2] #K
                    ]
    df=df.dropna(how='all').reset_index(drop=True)
    df['base_1'] = df.apply(lambda x: flag_box(x['Col1'], df['Col3'][0]), axis=1) # there will be some headers that check this, so headers is first in if elif logic
    #df['data'] = df.apply(lambda x: flag_box(x['Col3'], x['Col1']), axis=1)
    df['total'] = df.apply(lambda x: flag_total_rows_3(x['Col5']), axis=1)
    # wirte excel
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A2:N2", header_2, header_format_1)
    worksheet.merge_range("A1:N1", header_1, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A3:N3", header_3, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'left'
                                     })
    for col in range(14):
        try:
            worksheet.write(3, col, header_list_1[col], header_format_3)
        except:
            worksheet.write(3, col, '', header_format_3)
    worksheet.merge_range(4, 0, 4, 13, '', header_format_2)
    worksheet.set_row(4,7.5)
    row_write_val = 5
    data_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':4
                                    })
    data_format_1_center = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'center'
                                    , 'num_format':4
                                    })
    data_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':4
                                        })
    data_format_3 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':14
                                    , 'align':'center'
                                        })
    data_format_4 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'mm-yyyy'
                                        })
    data_format_total = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':4
                                    , 'bold': True
                                    , 'border_color': black_color
                                    , 'border':6
                                    , 'top': 0
                                    , 'left':0
                                    , 'right':0
                                    #, 'bottom':3
                                        })
    data_format_1_notes = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':4
                                    , 'text_wrap':True
                                    })
    bottom_line_two_format = workbook.add_format({'font_color': black_color
                                    , 'border_color': black_color
                                    , 'bottom': 1
                                    })
    for i in range(3, df.shape[0]):
        #print(i)
        try:
            next_base_1 = df['base_1'][i-1]
        except:
            next_base_1 = 0
        if i == df.shape[0] - 1:
            for j in range(14):
                if j == 9:
                    try:
                        worksheet.write_number(row_write_val, j, df['Col10'][i],data_format_total)
                    except:
                        worksheet.write(row_write_val, j, df['Col10'][i],data_format_total)
                elif j == 10:
                    try:
                        worksheet.write_number(row_write_val, j, df['Col11'][i],data_format_total)
                    except:
                        worksheet.write_blank(row_write_val, j, None,data_format_total)
                else:
                    worksheet.write_blank(row_write_val, j, None,data_format_total)
            row_write_val += 1
        elif df['base_1'][i] == 1:
            if 1 == 1:
                worksheet.write(row_write_val, 0, df['Col1'][i],data_format_1)
                worksheet.write(row_write_val, 1, df['Col2'][i],data_format_1)
                try:
                    worksheet.write(row_write_val, 2, df['Col3'][i],data_format_4)
                except:
                    worksheet.write_blank(row_write_val, 2, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 3, df['Col4'][i],data_format_3)
                except:
                    worksheet.write_blank(row_write_val, 3, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 4, df['Col5'][i],data_format_1_center)
                except:
                    worksheet.write_blank(row_write_val, 4, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 5, df['Col6'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 5, None ,data_format_1)
                try:
                    worksheet.write(row_write_val, 6, df['Col7'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 6, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 7, df['Col8'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 7, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 8, df['Col9'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 8, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 9, df['Col10'][i],data_format_2)
                except:
                    worksheet.write_blank(row_write_val, 9, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 10, df['Col11'][i],data_format_2)
                except:
                    worksheet.write_blank(row_write_val, 10, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 11, df['Col12'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 11, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 12, df['Col13'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 12, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 13, df['Col14'][i],data_format_1_notes)
                except:
                    worksheet.write_blank(row_write_val, 13, None,data_format_1_notes)
                row_write_val += 1
        else:
            if 1 == 1:
                try:
                    worksheet.write(row_write_val, 5, df['Col6'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 5, None ,data_format_1)
                try:
                    worksheet.write(row_write_val, 6, df['Col7'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 6, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 7, df['Col8'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 7, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 8, df['Col9'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 8, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 9, df['Col10'][i],data_format_2)
                except:
                    worksheet.write_blank(row_write_val, 9, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 10, df['Col11'][i],data_format_2)
                except:
                    worksheet.write_blank(row_write_val, 10, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 11, df['Col12'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 11, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 12, df['Col13'][i],data_format_1)
                except:
                    worksheet.write_blank(row_write_val, 12, None,data_format_1)
                try:
                    worksheet.write(row_write_val, 13, df['Col14'][i],data_format_1_notes)
                except:
                    worksheet.write_blank(row_write_val, 13, None,data_format_1_notes)
                row_write_val += 1
                if next_base_1 == 1:
                    if i == df.shape[0]-2:
                        worksheet.merge_range(row_write_val, 0, row_write_val, 13, '', bottom_line_two_format)#bottom_line_two_format
                        worksheet.set_row(row_write_val,7.5)
                        row_write_val = row_write_val + 1

                    else:
                        worksheet.set_row(row_write_val,7.5)
                        row_write_val = row_write_val + 1
    column_width_list = [
                [9.8, 0, 0, worksheet] ## JE Register 0
                ,[10.2, 1, 3, worksheet] ## JE Register 1 
                ,[9.5, 4, 4, worksheet] ## JE Register 2
                ,[12.8, 5, 5, worksheet] ## JE Register 3
                ,[42.2, 6, 6, worksheet] ## JE Register 4
                ,[16, 7, 7, worksheet] ## JE Register 5
                ,[19.82, 8, 8, worksheet] ## JE Register 6
                ,[12.5, 9, 10, worksheet] ## JE Register 7
                ,[20, 11, 11, worksheet] ## JE Register 8
                ,[31, 12, 12, worksheet] ## JE Register 9
                ,[37.5, 13, 13, worksheet] ## JE Register
    ]
    loop_val = 0
    for i in column_width_list:
        try:
            if loop_val in [5, 6, 9]:
                i[3].set_column(i[1],i[2], i[0], options={'hidden':True})
            else:
                i[3].set_column(i[1],i[2], i[0])
        except:
            pass
        loop_val += 1
    worksheet.set_landscape()
    worksheet.set_margins(.5,.5,.5,.5)
    worksheet.repeat_rows(0, 3)
    worksheet.print_area(0,0, row_write_val, 13)
    worksheet.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    worksheet.fit_to_pages(1, total_pages)
    return df
# ----------------------------------------------------------------- new stuff ap_detail_sheet_def_2
def ap_detail_sheet_def_2(workbook, df, worksheet):
    yellow_color = '#b4992d'
    dark_gray_color = '#505050'
    white_color = '#FFFFFF'
    black_color = '#000000'
    grey_color = '#211f20'
    header_1 = df.columns[0]
    df = df.rename(columns={df.columns[0]: 'Col1'
                           , df.columns[1]: 'Col2'
                           , df.columns[2]: 'Col3'
                           , df.columns[3]: 'Col4'
                           , df.columns[4]: 'Col5'
                           , df.columns[5]: 'Col6'
                           , df.columns[6]: 'Col7'
                           , df.columns[7]: 'Col8'
                           , df.columns[8]: 'Col9'
                           , df.columns[9]: 'Col10'
                           , df.columns[10]: 'Col11'
                           , df.columns[11]: 'Col12'
                           , df.columns[12]: 'Col13'
                           , df.columns[13]: 'Col14'
                           , df.columns[14]: 'Col15'
                           , df.columns[15]: 'Col16'
                           })
    header_2 = df['Col1'][0]
    header_3 = df['Col1'][1]
    header_4 = df['Col1'][2]
    header_list_1 = [df['Col1'][3] #A
                     , df['Col2'][3] #B
                     , df['Col3'][3] #C
                     , df['Col4'][3] #D
                     , df['Col5'][3] #E
                     , df['Col6'][3] #F
                     , df['Col7'][3] #G
                     , df['Col8'][3] #H
                     , df['Col9'][3] #I
                     , df['Col10'][3] #J
                     , df['Col11'][3] #K
                     , df['Col12'][3] #M
                     , df['Col13'][3] #N
                     , df['Col14'][3] #O
                     , df['Col15'][3] #P
                     , df['Col16'][3] #Q
                    ]
    header_list_2 = [df['Col1'][4] #A
                     , df['Col2'][4] #B
                     , df['Col3'][4] #C
                     , df['Col4'][4] #D
                     , df['Col5'][4] #E
                     , df['Col6'][4] #F
                     , df['Col7'][4] #G
                     , df['Col8'][4] #H
                     , df['Col9'][4] #I
                     , df['Col10'][4] #J
                     , df['Col11'][4] #K
                     , df['Col12'][4] #M
                     , df['Col13'][4] #N
                     , df['Col14'][4] #O
                     , df['Col15'][4] #P
                     , df['Col16'][4] #Q
                    ]
    header_list_3 = [df['Col1'][5] #A
                     , df['Col2'][5] #B
                     , df['Col3'][5] #C
                     , df['Col4'][5] #D
                     , df['Col5'][5] #E
                     , df['Col6'][5] #F
                     , df['Col7'][5] #G
                     , df['Col8'][5] #H
                     , df['Col9'][5] #I
                     , df['Col10'][5] #J
                     , df['Col11'][5] #K
                     , df['Col12'][5] #M
                     , df['Col13'][5] #N
                     , df['Col14'][5] #O
                     , df['Col15'][5] #P
                     , df['Col16'][5] #Q
                    ]
    df=df.dropna(how='all').reset_index(drop=True)
    df['color_col'] = df.apply(lambda x: flag_box(x['Col1'], x['Col11']), axis=1) # there will be some headers that check this, so headers is first in if elif logic
    df['data'] = df.apply(lambda x: flag_box(x['Col3'], x['Col2']), axis=1)
    df['total'] = df.apply(lambda x: flag_total_rows_2(x['Col1']), axis=1)
    # wirte excel
    header_format_1 = workbook.add_format({'font_color': black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':14
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A1:P1", header_1, header_format_1)
    header_format_2 = workbook.add_format({'font_color': dark_gray_color
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    worksheet.merge_range("A2:P2", header_2, header_format_2)
    worksheet.merge_range("A3:P3", header_3, header_format_2)
    worksheet.merge_range("A4:P4", header_4, header_format_2)
    header_format_3 = workbook.add_format({'font_color': white_color
                                    , 'bg_color':black_color
                                    , 'bold':True
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':11
                                    , 'align':'center'
                                     })
    for col in range(16):
        try:
            str(header_list_1[col])
            worksheet.write(4, col, header_list_1[col], header_format_3)
        except:
            worksheet.write_blank(4, col, None, header_format_3)
        try:
            str(header_list_2[col])
            worksheet.write(5, col, header_list_2[col], header_format_3)
        except:
            worksheet.write_blank(5, col, None, header_format_3)
        try:
            str(header_list_3[col])
            worksheet.write(6, col, header_list_3[col], header_format_3)
        except:
            worksheet.write_blank(6, col, None, header_format_3)
    worksheet.merge_range(7, 0, 7, 14, '', header_format_2)
    worksheet.set_row(7,7.5)
    row_write_val = 8
    header_format_body = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                     })
    header_format_date_1 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'mm-yyyy'
                                     })
    header_format_date_2 = workbook.add_format({'font_color': black_color
                                    , 'bg_color':'#EEEEEE'
                                    , 'bold':False
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':14
                                    , 'align':'center'
                                     })
    data_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'text_wrap':True
                                    })
    data_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                        })
    data_format_3 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':14
                                    , 'align':'center'
                                        })
    total_format_1 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':'#EEEEEE'
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color': black_color
                                        })
    total_format_2 = workbook.add_format({'font_color': black_color
                                    #, 'bg_color':'#EEEEEE'
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'top':1
                                    , 'bottom':1
                                    , 'border_color': black_color
                                        })
    grand_total_format_bottom = workbook.add_format({'font_color': black_color
                                    , 'border_color': black_color
                                    , 'bottom':1
                                        })
    grand_total_format_1 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'bottom':1
                                    , 'border_color': black_color
                                        })
    grand_total_format_2 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'left'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'border':6
                                    , 'left':0
                                    , 'top':0
                                    , 'right':0
                                    #, 'bottom':3
                                    , 'border_color':black_color
                                        })
    grand_total_format_3 = workbook.add_format({'font_color': black_color
                                    , 'font_name': 'Century Gothic'
                                    , 'font_size':10
                                    , 'align':'right'
                                    , 'num_format':'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                                    , 'bold':True
                                    , 'border':6
                                    , 'left':0
                                    , 'top':0
                                    , 'right':0
                                    #, 'bottom':3
                                    , 'border_color':black_color
                                    , 'text_wrap':True
                                        })
    for i in range(5, df.shape[0]): #flag_total_rows_2
        if df['color_col'][i] == 1:
            worksheet.write_string(row_write_val, 0, df['Col1'][i],header_format_body)
            try:
                worksheet.write(row_write_val, 1, df['Col2'][i],header_format_body)
            except:
                worksheet.write_blank(row_write_val, 1, None,header_format_body)
            worksheet.write_blank(row_write_val, 2, None,header_format_body)
            worksheet.write_blank(row_write_val, 3, None,header_format_body)
            worksheet.write_blank(row_write_val, 4, None,header_format_body)
            worksheet.write_blank(row_write_val, 5, None,header_format_body)
            worksheet.write_blank(row_write_val, 6, None,header_format_body)
            worksheet.write_blank(row_write_val, 7, None,header_format_body)
            worksheet.write_blank(row_write_val, 8, None,header_format_body)
            worksheet.write_blank(row_write_val, 9, None,header_format_body)
            worksheet.write_blank(row_write_val, 10, None,header_format_body)
            worksheet.write_blank(row_write_val, 11, None,header_format_body)
            worksheet.write_blank(row_write_val, 12, None,header_format_body)
            worksheet.write_blank(row_write_val, 13, None,header_format_body)
            worksheet.write_blank(row_write_val, 14, None,header_format_body)
            worksheet.write_blank(row_write_val, 15, None,header_format_body)
            row_write_val += 1
        elif df['total'][i] == 1:
            if df['Col1'][i] == 'Grand Total':
                # grand_total_format_bottom
                worksheet.write_string(row_write_val, 0, df['Col1'][i],grand_total_format_2)
                worksheet.write_blank(row_write_val, 1, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 2, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 3, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 4, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 5, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 6, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 7, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 8, None ,grand_total_format_2)
                worksheet.write_number(row_write_val, 9, df['Col10'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 10, df['Col11'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 11, df['Col12'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 12, df['Col13'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 13, df['Col14'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 14, df['Col15'][i],grand_total_format_3)
                worksheet.write_blank(row_write_val, 15, None ,grand_total_format_3)
                worksheet.merge_range(row_write_val - 1, 0, row_write_val - 1, 15, '', grand_total_format_bottom)
                row_write_val += 1
            elif df['Col1'][i] == 'Grand Total ':
                # grand_total_format_bottom
                worksheet.write_string(row_write_val, 0, df['Col1'][i],grand_total_format_2)
                worksheet.write_blank(row_write_val, 1, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 2, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 3, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 4, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 5, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 6, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 7, None ,grand_total_format_2)
                worksheet.write_blank(row_write_val, 8, None ,grand_total_format_2)
                worksheet.write_number(row_write_val, 9, df['Col10'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 10, df['Col11'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 11, df['Col12'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 12, df['Col13'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 13, df['Col14'][i],grand_total_format_3)
                worksheet.write_number(row_write_val, 14, df['Col15'][i],grand_total_format_3)
                worksheet.write_blank(row_write_val, 15, None ,grand_total_format_3)
                worksheet.merge_range(row_write_val - 1, 0, row_write_val - 1, 15, '', grand_total_format_bottom)
                row_write_val += 1
            else:
                worksheet.write_string(row_write_val, 0, df['Col1'][i],total_format_1)
                worksheet.write_blank(row_write_val, 1, None ,total_format_1)
                worksheet.write_blank(row_write_val, 2, None ,total_format_1)
                worksheet.write_blank(row_write_val, 3, None ,total_format_1)
                worksheet.write_blank(row_write_val, 4, None ,total_format_1)
                worksheet.write_blank(row_write_val, 5, None ,total_format_1)
                worksheet.write_blank(row_write_val, 6, None ,total_format_1)
                worksheet.write_blank(row_write_val, 7, None ,total_format_1)
                worksheet.write_blank(row_write_val, 8, None ,total_format_1)
                worksheet.write_number(row_write_val, 9, df['Col10'][i],total_format_2)
                worksheet.write_number(row_write_val, 10, df['Col11'][i],total_format_2)
                worksheet.write_number(row_write_val, 11, df['Col12'][i],total_format_2)
                worksheet.write_number(row_write_val, 12, df['Col13'][i],total_format_2)
                worksheet.write_number(row_write_val, 13, df['Col14'][i],total_format_2)
                worksheet.write_number(row_write_val, 14, df['Col15'][i],total_format_2)
                worksheet.write_blank(row_write_val, 15, None ,total_format_1)
                row_write_val += 1
                worksheet.set_row(row_write_val,7.5)
                row_write_val = row_write_val + 1
        elif df['data'][i] == 1:
            worksheet.write(row_write_val, 2, df['Col3'][i],data_format_1)
            worksheet.write(row_write_val, 3, df['Col4'][i],data_format_1)
            worksheet.write(row_write_val, 4, df['Col5'][i],data_format_1)
            worksheet.write(row_write_val, 5, df['Col6'][i],data_format_1)
            worksheet.write(row_write_val, 6, df['Col7'][i],data_format_3)
            worksheet.write(row_write_val, 7, df['Col8'][i],data_format_1)
            worksheet.write(row_write_val, 8, df['Col9'][i],data_format_1)
            worksheet.write(row_write_val, 9, df['Col10'][i],data_format_2)
            worksheet.write(row_write_val, 10, df['Col11'][i],data_format_2)
            worksheet.write(row_write_val, 11, df['Col12'][i],data_format_2)
            worksheet.write(row_write_val, 12, df['Col13'][i],data_format_2)
            worksheet.write(row_write_val, 13, df['Col14'][i],data_format_2)
            worksheet.write(row_write_val, 14, df['Col15'][i],data_format_2)
            worksheet.write(row_write_val, 15, df['Col16'][i],data_format_1)
            row_write_val += 1
        else:
            pass
    column_width_list = [[13.8, 0, 0, worksheet] ## AP Detail
                ,[12.2, 1, 1, worksheet] ## AP Detail
                ,[25.2, 2, 2, worksheet] ## AP Detail
                ,[10, 3, 3, worksheet] ## AP Detail
                ,[10, 4, 4, worksheet] ## AP Detail
                ,[14, 5, 5, worksheet] ## AP Detail
                ,[14, 6, 6, worksheet] ## AP Detail
                ,[20, 7, 7, worksheet] ## AP Detail
                ,[20, 8, 8, worksheet] ## AP Detail
                ,[10.73, 9, 14, worksheet] ## AP Detail
                ,[35.73, 15, 15, worksheet] ## AP Detail
    ]
    for i in column_width_list:
        try:
            i[3].set_column(i[1],i[2], i[0])
        except:
            pass
    worksheet.set_landscape()
    worksheet.set_margins(.5,.5,.5,.5)
    worksheet.repeat_rows(0, 7)
    worksheet.print_area(0,0, row_write_val, 15)
    worksheet.set_page_view(2)
    total_pages = max(math.ceil(row_write_val/50), 1)
    worksheet.fit_to_pages(1, total_pages)
    return df
# -----------------------------------------------------------------
