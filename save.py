import pandas as pd
import numpy as np
import os
import openpyxl as op
from openpyxl.styles import fills,colors,NamedStyle,Font,Side,Border,PatternFill,Alignment,Protection


def beauty(wb_name):
    wb = op.load_workbook(wb_name)
    ft = op.styles.Font(name='宋体', size=12, bold=False)  
    ft1 =  op.styles.Font(name='黑体', size=13, bold=False)  
    align = op.styles.Alignment(horizontal='center',vertical='center' )
    border = op.styles.Border(left=Side(border_style="thin"),
                                right=Side(border_style="thin"),         
                                top=Side(border_style="thin"),
                                bottom=Side(border_style="thin"))
    
    for sheet in wb:
        autofit(sheet)
        for row in sheet.rows:
            for cell in row:
                cell.font= ft
                cell.alignment =align
                cell.border = border
                if cell.row<3 or cell.column<3:
                    cell.font = ft1
    
    wb.save(wb_name)


def autofit(ws):

    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                len_cell = max([(len(line.encode('utf-8'))-len(line))/2+len(line) for line in str(cell.value).split('\n')])
                #dims[chr(64+cell.column)] = max((dims.get(chr(64+cell.column), 0), len(str(cell.value))))
                dims[cell.column_letter] = max(dims.get(cell.column_letter, 0), len_cell)
    for col, value in dims.items():

        ws.column_dimensions[col].width = value+2 if value+2<=50 else 50


if __name__ == "__main__":
    dir_desktop = os.path.join(os.path.expanduser("~"), 'Desktop').replace("\\", "/")+"/"
    dir =  dir_desktop + '存酒汇总/'

    
    df = pd.read_excel(dir_desktop+'存酒明细.xlsx')
    
    stuff = pd.read_excel(dir_desktop+"基础数据.xlsx", sheet_name='员工名单')
    result = pd.merge(df, stuff, left_on='业务经理', right_on='姓名', how='left')
    result['部门'] = result['部门'].fillna('外卖台')
    result['手机号'] = result['手机号'].apply(lambda x: str(x)[0:3] + "****" + str(x)[7:11])
    print(result)
    departs = result['部门'].unique()
    
    df2=pd.DataFrame()
    for depart in departs:
        if depart != '外卖台':
            df2 = result.query('部门==@depart & 状态=="有效"')[['存酒房台','业务经理', '到期日期', '手机号', '酒水名称','数量']]
            df2.sort_values(by='到期日期', inplace=True)
            df2.sort_values(by='业务经理', inplace=True)
        
        else:
            df2 = result.query('部门==@depart & 状态=="有效"')[['到期日期', '手机号', '酒水名称','数量']]
            df2.sort_values(by='到期日期', inplace=True)
        print(df2)
        if not df2.empty:
            df2.to_excel(dir+depart+'.xlsx', sheet_name=depart, index=False)
            beauty(dir+depart+'.xlsx')


