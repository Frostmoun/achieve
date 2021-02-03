import pretty_errors
import pandas as pd
import os
import data_sourse as god
import numpy as np
import openpyxl as op


from consts import * \



# 核对业绩表:房台名称
if __name__ == "__main__":
    writer = pd.ExcelWriter('day_report.xlsx')    
    df = god.get_table_clean()
    depart = df['部门'].sort_values().unique()

    day = df['日期'].unique()[-1]

    check = pd.DataFrame()
    total = pd.DataFrame()

    for d in depart:
        temp = pd.pivot_table(df.query('日期 == @day and 部门 ==@d'), index=['部门', '订台人', '房台'], values={'实际业绩', '消费合计', '订房现抽'}, 
                        aggfunc={'实际业绩':np.sum, '消费合计':np.sum, '订房现抽':np.sum}, fill_value=0, margins=True, margins_name='汇总')
        if temp.loc['汇总']['实际业绩'].values != 0:
            check = check.append(temp)

    for d in depart:       
        temp = pd.pivot_table(df.query('部门 == @d'), index=['部门', '订台人'], values=['房台', '实际业绩', '消费合计'], columns='日期', 
                        aggfunc={'房台':'count', '实际业绩':np.sum, '消费合计':np.sum}, fill_value=0, margins=True, margins_name='汇总')

        total = total.append(temp)
       
    total = total[[('房台', day), ('实际业绩', day), ('消费合计', day), ('房台', '汇总'), ('实际业绩', '汇总'), ('消费合计', '汇总')]]
    check.to_excel(writer, sheet_name='核对表')

    total.to_excel(writer, sheet_name='汇总表')

    writer.save()
    writer.close()