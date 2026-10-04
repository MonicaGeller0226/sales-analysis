# -*- coding: utf-8 -*-
"""
数据准备脚本：读取原始数据，清洗，计算衍生指标，生成透视表到一个Excel工作簿
输出：
  - 销售数据分析.xlsx  (包含原始数据、汇总、多个透视表sheet)
  - cleaned_data.pkl   (清洗后的数据，供绘图脚本使用)
"""
import pandas as pd
import os

SRC = r'C:\Users\HQ\OneDrive\Desktop\7月销售数据.xlsx'
OUT_DIR = r'D:\Python\Python_item\PythonProject\销售数据分析'
OUT_XLSX = os.path.join(OUT_DIR, '销售数据分析.xlsx')

# ---------- 1. 读取 ----------
df = pd.read_excel(SRC)

# ---------- 2. 清洗 ----------
# 日期列是Excel序列号，转为标准日期
df['日期'] = pd.to_datetime(df['日期'], unit='D', origin='1899-12-30')
df['星期'] = df['日期'].dt.day_name(locale='zh_CN')
# 计算ISO周年份+周数（周一为一周开始）
iso = df['日期'].dt.isocalendar()
df['周数'] = iso['week'].astype(int)
df['周标签'] = 'W' + (df['周数'] - df['周数'].min() + 1).astype(str)
df['年'] = iso['year']

# 剔除成交台数为负/异常的脏数据（如有)
df = df[df['成交台数'] >= 0].copy()

# ---------- 3. 衍生指标 ----------
# 各转化率(百分比)
df['线索转化率%'] = (df['成交台数'] / df['线索量'] * 100).round(2)
df['到店转化率%'] = (df['成交台数'] / df['到店人数'] * 100).round(2)
df['到店线索率%'] = (df['到店人数'] / df['线索量'] * 100).round(2)

print('清洗后数据 shape:', df.shape)
print('缺失值统计:\n', df.isna().sum())

# ---------- 4. 汇总总览 ----------
overall = pd.DataFrame([{
    '指标': '总成交台数', '数值': int(df['成交台数'].sum()),
    '口径': 'sum(成交台数)',
}, {
    '指标': '总线索量', '数值': int(df['线索量'].sum()),
    '口径': 'sum(线索量)',
}, {
    '指标': '总到店人数', '数值': int(df['到店人数'].sum()),
    '口径': 'sum(到店人数)',
}, {
    '指标': '整体线索转化率%', '数值': round(df['成交台数'].sum()/df['线索量'].sum()*100, 2),
    '口径': '总成交/总线索',
}, {
    '指标': '整体到店转化率%', '数值': round(df['成交台数'].sum()/df['到店人数'].sum()*100, 2),
    '口径': '总成交/总到店',
}, {
    '指标': '整体到店线索率%', '数值': round(df['到店人数'].sum()/df['线索量'].sum()*100, 2),
    '口径': '总到店/总线索',
}, {
    '指标': '涉及城市数', '数值': df['城市'].nunique(), '口径': 'nunique',
}, {
    '指标': '涉及门店数', '数值': df['门店名称'].nunique(), '口径': 'nunique',
}, {
    '指标': '涉及车型数', '数值': df['车型'].nunique(), '口径': 'nunique',
}, {
    '指标': '涉及销售顾问数', '数值': df['销售顾问'].nunique(), '口径': 'nunique',
}, {
    '指标': '数据天数', '数值': df['日期'].nunique(), '口径': 'nunique',
}])

# ---------- 5. 生成透视表 ----------
def pt(df, index, columns=None, values='成交台数', aggfunc='sum'):
    return pd.pivot_table(df, index=index, columns=columns,
                          values=values, aggfunc=aggfunc,
                          fill_value=0, margins=True, margins_name='合计')

pivots = {}
# A. 城市 x 车型  成交台数
pivots['城市×车型_成交台数'] = pt(df, '城市', '车型')
# B. 城市 x 车型  线索量
pivots['城市×车型_线索量'] = pt(df, '城市', '车型', values='线索量')
# C. 城市 x 车型  到店人数
pivots['城市×车型_到店人数'] = pt(df, '城市', '车型', values='到店人数')
# D. 车型 x 周  成交台数
pivots['车型×周_成交台数'] = pt(df, '车型', '周标签')
# E. 城市 x 周  成交台数
pivots['城市×周_成交台数'] = pt(df, '城市', '周标签')
# F. 门店 x 销售顾问  成交台数(占比)
piv0 = pd.pivot_table(df, index='门店名称', columns='销售顾问',
                      values='成交台数', aggfunc='sum', fill_value=0, margins=True, margins_name='合计')
pivots['门店×顾问_成交台数'] = piv0
# G. 门店 x 车型  成交台数
pivots['门店×车型_成交台数'] = pt(df, '门店名称', '车型')
# H. 城市 各指标汇总
piv_city = pd.pivot_table(df, index='城市', values=['成交台数','线索量','到店人数'],
                          aggfunc='sum', fill_value=0).round(2)
piv_city['线索转化率%'] = (piv_city['成交台数']/piv_city['线索量']*100).round(2)
piv_city['到店转化率%'] = (piv_city['成交台数']/piv_city['到店人数']*100).round(2)
piv_city['到店线索率%'] = (piv_city['到店人数']/piv_city['线索量']*100).round(2)
pivots['城市_汇总指标'] = piv_city.sort_values('成交台数', ascending=False)
# I. 车型 各指标汇总
piv_car = pd.pivot_table(df, index='车型', values=['成交台数','线索量','到店人数'],
                         aggfunc='sum', fill_value=0).round(2)
piv_car['线索转化率%'] = (piv_car['成交台数']/piv_car['线索量']*100).round(2)
piv_car['到店转化率%'] = (piv_car['成交台数']/piv_car['到店人数']*100).round(2)
piv_car['到店线索率%'] = (piv_car['到店人数']/piv_car['线索量']*100).round(2)
pivots['车型_汇总指标'] = piv_car.sort_values('成交台数', ascending=False)
# J. 销售顾问 汇总
piv_sale = pd.pivot_table(df, index='销售顾问', values=['成交台数','线索量','到店人数'],
                          aggfunc='sum', fill_value=0).round(2)
piv_sale['线索转化率%'] = (piv_sale['成交台数']/piv_sale['线索量']*100).round(2)
piv_sale['到店转化率%'] = (piv_sale['成交台数']/piv_sale['到店人数']*100).round(2)
piv_sale['到店线索率%'] = (piv_sale['到店人数']/piv_sale['线索量']*100).round(2)
piv_sale['平均每日成交'] = (piv_sale['成交台数']/df['日期'].nunique()).round(2)
pivots['销售顾问_汇总指标'] = piv_sale.sort_values('成交台数', ascending=False)
# K. 周汇总
piv_week = pd.pivot_table(df, index='周标签', values=['成交台数','线索量','到店人数'],
                          aggfunc='sum', fill_value=0).round(2)
piv_week['线索转化率%'] = (piv_week['成交台数']/piv_week['线索量']*100).round(2)
piv_week['到店转化率%'] = (piv_week['成交台数']/piv_week['到店人数']*100).round(2)
piv_week['到店线索率%'] = (piv_week['到店人数']/piv_week['线索量']*100).round(2)
pivots['周_汇总指标'] = piv_week.sort_index()

# ---------- 6. 写入Excel(所有透视表在一个工作簿) ----------
with pd.ExcelWriter(OUT_XLSX, engine='openpyxl') as writer:
    df.round(2).to_excel(writer, sheet_name='清洗后明细', index=False)
    overall.to_excel(writer, sheet_name='总体概览', index=False)
    for name, piv in pivots.items():
        piv.round(2).to_excel(writer, sheet_name=name, index=True)

print(f'\n已生成透视表工作簿: {OUT_XLSX}')
print('工作簿sheet列表: 清洗后明细, 总体概览, ' + ', '.join(pivots.keys()))

# 保存清洗后数据供绘图脚本使用
df.to_pickle(os.path.join(OUT_DIR, 'cleaned_data.pkl'))
print('清洗后数据已保存 cleaned_data.pkl')
