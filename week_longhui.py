import pretty_errors
import pandas as pd
import os
import numpy as np
import openpyxl as op
import datetime


# 部门列表
DEPARTS = ["销-销售1部", "销-销售2部", "销-销售5部", "销-销售6部", "销-销售7部", "销-销售8部", "销-销售9部", "国际部", "市场部", '出品部', '先锋团', '独立团',  '独-2部', '独-1部', '独-3部', '会员中心', '运营部','礼宾部', '工程部', '资-B组', '资-气氛1部']
MAIN_DEPARTS = ["销-销售1部", "销-销售2部", "销-销售5部", "销-销售6部", "销-销售7部", "销-销售8部", "销-销售9部", "国际部", "市场部", '先锋团', '独立团', '会员中心', '运营部', '资源部']
# 桌面路径
DIR_DESKTOP = os.path.join(os.path.expanduser("~"), 'Desktop').replace("\\", "/")+"/"

# 读取文件
# plan = pd.read_excel(DIR_DESKTOP + "稽核/仓库/基础数据.xlsx", sheet_name="现抽方案")
# staff = pd.read_excel(DIR_DESKTOP + "稽核/仓库/基础数据.xlsx", sheet_name="艺人名单")
# detail = pd.read_excel(DIR_DESKTOP + "落单明细.xlsx").fillna("")
# table = pd.read_excel(DIR_DESKTOP + "营业日报.xlsx", header=None)
# total_award = pd.read_excel(DIR_DESKTOP + "/稽核/仓库/现抽汇总表.xlsx", sheet_name='汇总')
# total_basket = pd.read_excel(DIR_DESKTOP + "/稽核/仓库/花单汇总表.xlsx", sheet_name='汇总')
# total_air = pd.read_excel(DIR_DESKTOP + "/稽核/仓库/礼炮汇总表.xlsx", sheet_name='汇总')
# total_achieve = pd.read_excel(DIR_DESKTOP + "/稽核/仓库/业绩汇总表.xlsx", sheet_name='汇总')


# ? 输出文件
week_report = pd.ExcelWriter(DIR_DESKTOP + "/稽核/周报/周报.xlsx")

main = pd.read_excel(DIR_DESKTOP + "/稽核/仓库/业绩汇总表.xlsx", sheet_name='汇总')
detail = pd.read_excel(DIR_DESKTOP + "/稽核/落单明细(检查用).xlsx", sheet_name='Sheet1')

# 周报
weeknum = main['周数'].max()
# 部门 周基本数据:台数 台类型  营业额  业绩  任务   完成率    赠送    部门个人数据
# * 周部门数据对比
week_depart = pd.pivot_table(main.query('周数 in [@weeknum, @weeknum-1] & 部门 in @DEPARTS'), index='部门', columns='周数', values=['房台','实际业绩','营业总收入'], aggfunc={'房台':'count', '实际业绩':np.sum,'营业总收入':np.sum})
# ! 周个人数据对比
week_person = pd.pivot_table(main.query('周数 in [@weeknum, @weeknum-1] & 主部门 in @MAIN_DEPARTS'), index=['主部门','订台人'], columns='周数', values=['房台','实际业绩','营业总收入'], aggfunc={'房台':'count', '实际业绩':np.sum,'营业总收入':np.sum}).reset_index()
# // 赠送数据
donate = pd.pivot_table(detail.query('主部门 in @MAIN_DEPARTS & 类型 =="经理赠送" & (落单人部门 in @MAIN_DEPARTS | 落单人 in ["王秀军2","卢涛","李文"])'), index='主部门', values='金额', aggfunc={'金额':np.sum})

# ? 周完成率
week_rate = pd.pivot_table(main.query('周数==@weeknum & 主部门 in @MAIN_DEPARTS'), index='主部门', values=['周业绩任务', '实际业绩', '周完成率'], aggfunc={'周业绩任务':np.mean, '实际业绩':np.sum, '周完成率':np.sum})

# ? 月完成率
month_rate = pd.pivot_table(main.query('主部门 in @MAIN_DEPARTS'), index='主部门', values=['月业绩任务', '实际业绩', '月完成率'], aggfunc={'月业绩任务':np.mean, '实际业绩':np.sum, '月完成率':np.sum})


# ? 每日营业额
day_data = pd.pivot_table(main, index= ['日期'],values=['实际业绩','主营业务收入', '营业外收入','营业总收入'], aggfunc={'实际业绩':np.sum, '主营业务收入':np.sum, '营业外收入':np.sum,'营业总收入':np.sum}).reset_index()[['日期','实际业绩','主营业务收入', '营业外收入','营业总收入']]


week_depart.to_excel(week_report, sheet_name='周部门数据对比')
week_person.to_excel(week_report, sheet_name='周个人数据对比')
donate.to_excel(week_report, sheet_name='部门赠送数据')
week_rate.to_excel(week_report, sheet_name='周完成率')
month_rate.to_excel(week_report, sheet_name='月完成率')
day_data.to_excel(week_report, sheet_name='每日营业额', index=False)


week_report.save()
week_report.close()
# 部门 月基本数据:台数 台类型  营业额  业绩   

# 门店  :  售出酒数量   消耗酒数量   营收   业绩  开台    收支差(资源, 礼宾, 楼面)