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
