import pretty_errors
import pandas as pd
import os
import numpy as np
import openpyxl as op
import datetime

# 业绩外支付项目
EXTRA_ACHIEVE = ['支付围台酒水', '支付职能招待', '支付宴请酒水', '支付会员权益'] 
# 部门列表
DEPARTS = ["销-2部", "销-3部", "销-5部", "销-6部", "销-8部", "销-9部", "市场部", "国际部"]

# 桌面路径
DIR_DESKTOP = os.path.join(os.path.expanduser("~"), 'Desktop').replace("\\", "/")
DIR_ROOT = "稽核/隆回"

# 读取文件
plan = pd.read_excel(DIR_ROOT + "/仓库/基础数据.xlsx", sheet_name="现抽方案")
staff = pd.read_excel(DIR_ROOT + "/仓库/基础数据.xlsx", sheet_name="艺人名单")
main_depart = pd.read_excel(DIR_ROOT + "/仓库/基础数据.xlsx", sheet_name="部门")
detail = pd.read_excel(DIR_DESKTOP + "/下载/落单明细_lh.xlsx", skipfooter=1)
table = pd.read_excel(DIR_DESKTOP + "/下载/营业日报_lh.xlsx", header=None, skipfooter=1)
total_award = pd.read_excel(DIR_ROOT + "/仓库/现抽汇总表.xlsx", sheet_name='汇总')
total_basket = pd.read_excel(DIR_ROOT + "/仓库/花单汇总表.xlsx", sheet_name='汇总')
total_air = pd.read_excel(DIR_ROOT + "/仓库/礼炮汇总表.xlsx", sheet_name='汇总')
total_achieve = pd.read_excel(DIR_ROOT + "/仓库/业绩汇总表.xlsx", sheet_name='汇总')
task_week = pd.read_excel(DIR_ROOT + "/周报/每周任务.xlsx", sheet_name='周任务')
task_month = pd.read_excel(DIR_ROOT + "/周报/每周任务.xlsx", sheet_name='月任务')

day = detail['日期'].max()
month = datetime.datetime.strptime(day,'%Y-%m-%d').strftime('%Y-%m')
# 保存路径
writer = pd.ExcelWriter(DIR_ROOT + "/每日现抽.xlsx")
writer_day = pd.ExcelWriter(DIR_ROOT + "/每日业绩/" + day + "每日个人业绩.xlsx")
writer_week = pd.ExcelWriter(DIR_ROOT + "/每周业绩/" + day + "每周个人业绩.xlsx")
writer_total_achieve = pd.ExcelWriter(DIR_ROOT + "/业绩汇总.xlsx")




table.columns = table.loc[0].ffill() + table.loc[1].fillna("")
table.fillna("", inplace=True)
table = table.drop(labels=[0,1],axis=0)
table = table.query('状态 != "取消预订"')

table["日期"] = table["日期"].apply(lambda x: str(datetime.datetime.now().year) +"-"+ str(x) if int(x[0:2]) <= datetime.datetime.now().month else str(datetime.datetime.now().year - 1) +"-"+ str(x))
table["日期主单"] = table["日期"].apply(lambda x: x.replace("-", "")) + table['主单']

table[['日期主单', '无业绩开台费','无业绩小费类', '无业绩赔偿类', '花单点舞小计', '计提成小计', '无业绩小计']] = table[['日期主单', '无业绩开台费','无业绩小费类', '无业绩赔偿类', '花单点舞小计', '计提成小计', '无业绩小计']].apply(pd.to_numeric)
for each in EXTRA_ACHIEVE:
    if each in table.columns:
        table[each] = table[each].apply(pd.to_numeric)

table.loc[table['支付职能招待']>0, ['订台人', '部门']] = ["散客", "散客"]
invalid = table.query('支付职能招待 > 0')['日期主单']
for each in invalid:
    detail.loc[detail['日期主单'] == each, ['订房人', "订房部门"]] = ["散客", "散客"]
re = pd.merge(detail, plan, on="酒水项目", how='left')
re = pd.merge(re, staff, on="艺人", how="left")
re = pd.merge(re, main_depart, left_on="订房部门", right_on='部门', how="left")
table = pd.merge(table, main_depart, on="部门", how="left")
re["订房现抽"] = 0
re.loc[(re['类型'] != "经理赠送") & (re['订房人'] != "自来客") & (re['订房人'] != "散客"),'订房现抽'] = re['落单金额']/re['单价']*re['现抽单价']
re.loc[(re['酒水项目']=="电子礼炮") & (re['类型'] == "消费"), '气氛道具扣除'] = re['落单金额'] * 0.25
re.to_excel(DIR_ROOT+'/落单明细(检查用).xlsx', index=False)
# 花单汇总表
girl_award = pd.pivot_table(re.query('艺人 != ""' ), index=['日期主单', '落单_时间','日期', '艺人部门', '艺人','房台', '酒水类别', '支付方式'], values=['数量', '落单金额'], aggfunc={'数量':np.sum, '落单金额':np.sum}).reset_index()
total_basket = total_basket.append(girl_award).drop_duplicates(keep ='first', inplace = False).sort_values(by='日期')
total_basket.to_excel(DIR_ROOT + "/仓库/花单汇总表.xlsx", sheet_name='汇总', index=False)

# 花单现抽
girl_award = pd.pivot_table(re.query('艺人部门=="资-B组" & 日期 == @day & 艺人 != ""' ), index=['日期', '艺人部门', '艺人','房台', '酒水类别', '支付方式'], values=['数量', '落单金额'], aggfunc={'数量':np.sum, '落单金额':np.sum}, margins=True).reset_index()
girl_award['提成金额'] = girl_award['落单金额']*0.5
girl_award.loc[girl_award['支付方式'] == "会员本金", '提成金额'] = girl_award['落单金额']*0.4
# 销售现抽
seller_award = pd.pivot_table(re.query('支付方式 != "围台酒水" & 房台 !="外卖台"'), index=['日期主单','落单_时间','日期', '房台', '酒水项目', '单价', '订房部门', '订房人', '支付方式'], values=['数量', '订房现抽'], aggfunc={'数量':np.sum, '订房现抽':np.sum}).query('订房现抽 != 0').reset_index()
# 更新现抽汇总表, 去掉重复项, 按时间排序
total_award = total_award.append(seller_award).drop_duplicates(keep ='first', inplace = False).sort_values(by='日期')
total_award.to_excel(DIR_ROOT + "/仓库/现抽汇总表.xlsx", sheet_name='汇总', index=False)


seller_award = seller_award[['日期', '房台', '酒水项目', '单价', '订房部门', '订房人', '支付方式', '数量', '订房现抽']].query('日期 == @day & 支付方式!="挂账"').sort_values(by='单价').reset_index()
seller_award.loc['合计']= seller_award[['数量', '订房现抽']].apply(lambda x:x.sum())
del seller_award['index']
# 副卡点舞提成
second_card = pd.pivot_table(re.query('酒水项目 == "点舞(副卡专用)" & 日期 == @day'), index=['日期', '房台', '酒水项目'], values=['数量', '落单金额'], aggfunc={'数量':np.sum, '落单金额':np.sum}, margins=True).reset_index()
second_card['提成金额'] = second_card['落单金额']*0.15



# 保存每日现抽  
girl_award.to_excel(writer, sheet_name="资源现抽")
seller_award.to_excel(writer, sheet_name="销售现抽")
second_card.to_excel(writer, sheet_name="副卡点舞")


# 保存现抽汇总表

# 电子礼炮
air = pd.pivot_table(re.query('酒水项目=="电子礼炮" & 类型=="消费" & 房台 !="外卖台"' ), index=['日期主单', '落单_时间','日期', '房台', '订房部门', '订房人', '支付方式'], values=['数量', '落单金额', "气氛道具扣除"], aggfunc={'数量':np.sum, '落单金额':np.sum, "气氛道具扣除":np.sum}).reset_index()
total_air = total_air.append(air).drop_duplicates(keep ='first', inplace = False).sort_values(by='日期')
total_air.to_excel(DIR_ROOT + "/仓库/礼炮汇总表.xlsx", sheet_name='汇总', index=False)
air = pd.pivot_table(total_air, index='日期主单', values='气氛道具扣除', aggfunc={"气氛道具扣除":np.sum}).reset_index().query('气氛道具扣除 != 0')
award = pd.pivot_table(total_award, index=['日期主单'], values='订房现抽', aggfunc={"订房现抽":np.sum}).reset_index().query('订房现抽 != 0')


# 不同门店有所不同
main = pd.merge(table, air, on='日期主单', how='left')
main = pd.merge(main, award, on='日期主单', how='left').fillna(0)

main['实际业绩'] = main['计提成小计'] -  main['气氛道具扣除'] - main['订房现抽_y'] 
for each in EXTRA_ACHIEVE:
    if each in table.columns:
        main['实际业绩'] = main['实际业绩'] - main[each]

main['主营业务收入'] = main['实际业绩'] + main['气氛道具扣除'] + main['订房现抽_y']
main['营业外收入'] = main['无业绩开台费'] + main['无业绩小费类'] + main['花单点舞小计'] + main['无业绩赔偿类']
main['营业总收入'] = main['主营业务收入'] + main['营业外收入'] 
main['检验值'] =main['营业总收入'] - main['主营业务收入'] - main['营业外收入'] + main['实际业绩'] + main['气氛道具扣除'] + main['订房现抽_y'] - main['计提成小计'] 


for each in EXTRA_ACHIEVE:
    if each in table.columns:
        main['检验值'] = main['检验值'] + main[each]
        
print('检验值:%d' % main['检验值'].sum())

main['周数'] = main['日期'].apply(lambda x: (int(datetime.datetime.strptime(x, '%Y-%m-%d').strftime('%W'))))
main['月份'] = main['日期'].apply(lambda x: (datetime.datetime.strptime(x, '%Y-%m-%d').strftime('%Y-%m')))

week = main['周数'].max()
month = main['月份'].max()
for index, row in task_week.iterrows():
    main.loc[(main['主部门']==row['部门']) & (main['周数']==row['周数']), '周业绩任务'] = row['周业绩任务']
    # main['周业绩任务'] = main.apply(lambda x: row['周业绩任务'] if (x.主部门 == row['部门'] and x.周数 == row['周数']) else 0, axis=1)
for index, row in task_month.iterrows():
    # print(main['月份']==row['月份'])
    # print(main['主部门']==row['部门'])
    main.loc[(main['主部门']==row['部门']) & (main['月份']==row['月份']), '月业绩任务'] = row['月业绩任务']
    # print(main['月业绩任务'])
    # main['月业绩任务'] = main.apply(lambda x: row['月业绩任务'] if (x.主部门 == row['部门'] and x.月份 == row['月份']) else 0, axis=1)

main['月完成率'] = main['实际业绩'] / main['月业绩任务']
main['周完成率'] = main['实际业绩'] / main['周业绩任务']
# main.loc[main.query('周业绩任务 != 0'), '周完成率'] = main['实际业绩'] / main['月业绩任务']

# 业绩汇总表
main.to_excel(writer_total_achieve,sheet_name="汇总", index=False)
temp = main[['日期','房台', '区域','订台人', '部门', '开台', '花单点舞小计', '气氛道具扣除', '订房现抽_y' ,'实际业绩','主营业务收入', '营业外收入','营业总收入', '周数', '月份', '主部门', '周业绩任务', '周完成率', '月业绩任务', '月完成率']]
total_achieve = total_achieve.append(temp).drop_duplicates(keep ='first', inplace = False).sort_values(by='日期')
total_achieve.to_excel(DIR_ROOT + "/仓库/业绩汇总表.xlsx", sheet_name='汇总', index=False)

for depart in DEPARTS:
    if depart in main['部门'].unique():
        # 每天业绩
        day_report = pd.pivot_table(main.query('日期==@day & 部门==@depart'), index=['订台人', '房台'], values=['实际业绩', '营业总收入'], aggfunc={'实际业绩':np.sum, '营业总收入':np.sum}, margins=True, margins_name="合计").reset_index()
        day_report.to_excel(writer_day, sheet_name=depart, index=False)
        # print(day_report)
        # 每周业绩
        week_report = pd.pivot_table(main.query('周数 == @week & 部门==@depart'), index=['订台人' ], values=['房台', '实际业绩'], aggfunc={'房台':'count', '实际业绩':np.sum})
        week_report = week_report.rename(columns={"实际业绩":"本周业绩", "房台":"本周台数"})
        # print(week)
        # 每月业绩
        month_report = pd.pivot_table(main.query('月份 == @month & 部门==@depart'), index=['订台人' ], values=['房台', '实际业绩'], aggfunc={'房台':'count', '实际业绩':np.sum})
        month_report = month_report.rename(columns={"实际业绩":"本月业绩", "房台":"本月台数"})
        week_report = pd.merge(week_report, month_report, on="订台人", how="right")
        
        week_report.loc['合计']= week_report.apply(lambda x:x.sum())
        week_report.to_excel(writer_week, sheet_name=depart)





writer.save()
writer.close()

writer_total_achieve.save()
writer_total_achieve.close()

writer_day.save()
writer_day.close()

writer_week.save()
writer_week.close()