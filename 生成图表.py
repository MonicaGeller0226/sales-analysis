# -*- coding: utf-8 -*-
r"""
绘图脚本：基于清洗后的数据生成各类统计图表
输出到 D:\Python\Python_item\PythonProject\销售数据分析\图表\
"""
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
import numpy as np

# 中文字体
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'KaiTi']
plt.rcParams['axes.unicode_minus'] = False

BASE = r'D:\Python\Python_item\PythonProject\销售数据分析'
OUT_DIR = os.path.join(BASE, '图表')
os.makedirs(OUT_DIR, exist_ok=True)

df = pd.read_pickle(os.path.join(BASE, 'cleaned_data.pkl'))

# 城市顺序(按销量)
city_order = df.groupby('城市')['成交台数'].sum().sort_values(ascending=False).index.tolist()
car_order = df.groupby('车型')['成交台数'].sum().sort_values(ascending=False).index.tolist()

COLORS = plt.cm.Set2(np.linspace(0, 1, 8))

# ============ 图表1: 每日销售趋势(成交+到店+线索) ============
daily = df.groupby('日期').agg(成交台数=('成交台数','sum'),
                               线索量=('线索量','sum'),
                               到店人数=('到店人数','sum')).reset_index()
fig, ax1 = plt.subplots(figsize=(12, 6))
ax1.bar(daily['日期'], daily['成交台数'], color='#4C72B0', width=0.7, label='成交台数')
ax1.set_ylabel('成交台数', color='#4C72B0')
ax1.tick_params(axis='y', labelcolor='#4C72B0')
ax2 = ax1.twinx()
ax2.plot(daily['日期'], daily['线索量'], color='#DD8452', marker='o', label='线索量')
ax2.plot(daily['日期'], daily['到店人数'], color='#55A868', marker='s', label='到店人数')
ax2.set_ylabel('线索量 / 到店人数')
ax2.legend(loc='upper left')
ax1.set_title('每日销售趋势 (成交 / 线索 / 到店)', fontsize=14)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '1_每日销售趋势.png'), dpi=150)
plt.close()

# ============ 图表2: 各周成交对比 ============
week = df.groupby('周标签').agg(成交台数=('成交台数','sum'),
                                线索量=('线索量','sum'),
                                到店人数=('到店人数','sum')).reset_index().sort_values('周标签')
x = np.arange(len(week))
w = 0.25
fig, ax = plt.subplots(figsize=(10, 6))
ax.bar(x-w, week['成交台数'], w, label='成交台数', color='#4C72B0')
ax.bar(x, week['到店人数'], w, label='到店人数', color='#55A868')
ax.bar(x+w, week['线索量'], w, label='线索量', color='#DD8452')
ax.set_xticks(x); ax.set_xticklabels(week['周标签'])
ax.set_title('分周汇总 (成交 / 到店 / 线索)', fontsize=14)
ax.legend()
for i in range(len(week)):
    ax.text(i-w, week['成交台数'][i]+3, str(week['成交台数'][i]), ha='center', fontsize=9)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '2_分周汇总.png'), dpi=150)
plt.close()

# ============ 图表3: 各城市销量对比(成交台数) ============
city_sum = df.groupby('城市').agg(成交台数=('成交台数','sum'),
                                  线索量=('线索量','sum'),
                                  到店人数=('到店人数','sum')).reindex(city_order)
fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.bar(city_sum.index, city_sum['成交台数'], color=COLORS[:len(city_sum)])
ax.bar_label(bars, padding=2)
ax.set_title('各城市成交台数对比', fontsize=14)
ax.set_ylabel('成交台数')
for i, city in enumerate(city_sum.index):
    cvt = city_sum.loc[city, '线索量']
    ax.text(i, city_sum['成交台数'].iloc[i]+4, f'线索{cvt:.0f}', ha='center', fontsize=9, color='#DD8452')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '3_各城市成交对比.png'), dpi=150)
plt.close()

# ============ 图表4: 各城市转化率 ============
city_cvt = pd.DataFrame({
    '线索转化率%': city_sum['成交台数']/city_sum['线索量']*100,
    '到店转化率%': city_sum['成交台数']/city_sum['到店人数']*100,
}).round(2)
fig, ax = plt.subplots(figsize=(10, 6))
x = np.arange(len(city_cvt))
w = 0.3
ax.bar(x-w/2, city_cvt['线索转化率%'], w, label='线索转化率%', color='#DD8452')
ax.bar(x+w/2, city_cvt['到店转化率%'], w, label='到店转化率%', color='#4C72B0')
ax.set_xticks(x); ax.set_xticklabels(city_cvt.index)
ax.set_title('各城市转化率对比', fontsize=14)
ax.legend()
for i in range(len(city_cvt)):
    ax.text(i-w/2, city_cvt['线索转化率%'].iloc[i]+0.3, f"{city_cvt['线索转化率%'].iloc[i]:.1f}%", ha='center', fontsize=8, color='#DD8452')
    ax.text(i+w/2, city_cvt['到店转化率%'].iloc[i]+0.3, f"{city_cvt['到店转化率%'].iloc[i]:.1f}%", ha='center', fontsize=8, color='#4C72B0')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '4_各城市转化率.png'), dpi=150)
plt.close()

# ============ 图表5: 车型销量占比(饼图) ============
car_sum = df.groupby('车型')['成交台数'].sum().reindex(car_order)
fig, ax = plt.subplots(figsize=(8, 8))
wedges, texts, autotexts = ax.pie(car_sum.values, labels=car_sum.index, autopct='%1.1f%%',
                                  colors=COLORS[:len(car_sum)], startangle=90, counterclock=False)
ax.set_title('各车型成交台数占比', fontsize=14)
for at in autotexts:
    at.set_fontsize(11); at.set_color('white')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '5_各车型销量占比.png'), dpi=150)
plt.close()

# ============ 图表6: 车型-转化率 ============
car_cvt = pd.DataFrame({
    '线索转化率%': df.groupby('车型')['成交台数'].sum()/df.groupby('车型')['线索量'].sum()*100,
    '到店转化率%': df.groupby('车型')['成交台数'].sum()/df.groupby('车型')['到店人数'].sum()*100,
}).reindex(car_order).round(2)
fig, ax = plt.subplots(figsize=(10, 6))
x = np.arange(len(car_cvt))
w = 0.3
ax.bar(x-w/2, car_cvt['线索转化率%'], w, label='线索转化率%', color='#DD8452')
ax.bar(x+w/2, car_cvt['到店转化率%'], w, label='到店转化率%', color='#4C72B0')
ax.set_xticks(x); ax.set_xticklabels(car_cvt.index)
ax.set_title('各车型转化率对比', fontsize=14)
ax.legend()
for i in range(len(car_cvt)):
    ax.text(i-w/2, car_cvt['线索转化率%'].iloc[i]+0.3, f"{car_cvt['线索转化率%'].iloc[i]:.1f}%", ha='center', fontsize=8)
    ax.text(i+w/2, car_cvt['到店转化率%'].iloc[i]+0.3, f"{car_cvt['到店转化率%'].iloc[i]:.1f}%", ha='center', fontsize=8)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '6_各车型转化率.png'), dpi=150)
plt.close()

# ============ 图表7: 销售顾问业绩排行 ============
sale = df.groupby('销售顾问').agg(成交台数=('成交台数','sum'),
                                  线索量=('线索量','sum'),
                                  到店人数=('到店人数','sum')).sort_values('成交台数', ascending=False)
fig, ax = plt.subplots(figsize=(10, 7))
bars = ax.barh(sale.index, sale['成交台数'], color='#4C72B0')
ax.bar_label(bars, padding=2)
ax.invert_yaxis()
ax.set_title('销售顾问成交台数排行', fontsize=14)
ax.set_xlabel('成交台数')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '7_销售顾问业绩排行.png'), dpi=150)
plt.close()

# ============ 图表8: 销售顾问转化率(散点: 产能 vs 转化率) ============
sale_cvt = pd.DataFrame({
    '成交台数': sale['成交台数'],
    '线索转化率%': (sale['成交台数']/sale['线索量']*100).round(2),
    '到店转化率%': (sale['成交台数']/sale['到店人数']*100).round(2),
})
fig, ax = plt.subplots(figsize=(10, 7))
sc = ax.scatter(sale_cvt['成交台数'], sale_cvt['线索转化率%'], s=120,
                c=sale_cvt['到店转化率%'], cmap='viridis')
for idx in sale_cvt.index:
    ax.annotate(idx, (sale_cvt.loc[idx,'成交台数'], sale_cvt.loc[idx,'线索转化率%']),
                xytext=(4,4), textcoords='offset points', fontsize=9)
cb = plt.colorbar(sc); cb.set_label('到店转化率%')
ax.axhline(sale_cvt['线索转化率%'].mean(), color='red', ls='--', lw=1, label='平均线索转化率')
ax.axvline(sale_cvt['成交台数'].mean(), color='gray', ls='--', lw=1, label='平均成交台数')
ax.set_title('销售顾问: 产能 vs 转化率', fontsize=14)
ax.set_xlabel('总成交台数'); ax.set_ylabel('线索转化率%')
ax.legend()
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '8_销售顾问产能转化散点.png'), dpi=150)
plt.close()

# ============ 图表9: 转化漏斗 ============
total_lx = df['线索量'].sum(); total_dd = df['到店人数'].sum(); total_cj = df['成交台数'].sum()
fig, ax = plt.subplots(figsize=(7, 6))
labels = ['线索量', '到店人数', '成交台数']
values = [total_lx, total_dd, total_cj]
colors = ['#DD8452', '#55A868', '#4C72B0']
bars = ax.barh(labels, values, color=colors)
ax.bar_label(bars, padding=2, fmt='%.0f')
for i, v in enumerate(values):
    conv = '' if i == 0 else f'  ({values[i]/values[i-1]*100:.1f}%)'
    ax.text(values[i]+5, i, conv, va='center', fontsize=10)
ax.set_title('整体转化漏斗 (线索→到店→成交)', fontsize=14)
ax.set_xlabel('数量')
ax.set_xlim(0, total_lx*1.12)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '9_整体转化漏斗.png'), dpi=150)
plt.close()

# ============ 图表10: 星期几销售规律 ============
weekday = df.groupby('星期').agg(成交台数=('成交台数','sum'),
                                  线索量=('线索量','sum'),
                                  到店人数=('到店人数','sum')).reindex(
    ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'])
wd_cn = { 'Monday':'周一','Tuesday':'周二','Wednesday':'周三','Thursday':'周四',
          'Friday':'周五','Saturday':'周六','Sunday':'周日'}
weekday.index = [wd_cn[i] for i in weekday.index]
fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.bar(weekday.index, weekday['成交台数'], color='#4C72B0')
ax.bar_label(bars, padding=2)
ax.set_title('星期几成交规律', fontsize=14)
ax.set_ylabel('总成交台数')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '10_星期几成交规律.png'), dpi=150)
plt.close()

# ============ 图表11: 热力图 - 城市x车型成交 ============
heat = pd.pivot_table(df, index='城市', columns='车型', values='成交台数', aggfunc='sum', fill_value=0)
heat = heat.reindex(index=city_order, columns=car_order)
fig, ax = plt.subplots(figsize=(9, 6))
im = ax.imshow(heat.values, cmap='YlOrRd')
ax.set_xticks(range(len(heat.columns))); ax.set_xticklabels(heat.columns)
ax.set_yticks(range(len(heat.index))); ax.set_yticklabels(heat.index)
for i in range(len(heat.index)):
    for j in range(len(heat.columns)):
        v = heat.values[i, j]
        ax.text(j, i, f'{v}', ha='center', va='center',
                color='white' if v > heat.values.max()*0.6 else 'black', fontsize=10)
ax.set_title('城市 × 车型 成交台数热力图', fontsize=14)
plt.colorbar(im, ax=ax, label='成交台数')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '11_城市车型热力图.png'), dpi=150)
plt.close()

# ============ 图表12: 门店销量占比(合计per顾问) ============
# 简化：显示各门店在所有城市中的销量及其线索质量
store = df.groupby('门店名称').agg(成交台数=('成交台数','sum'),
                                    线索量=('线索量','sum')).sort_values('成交台数', ascending=False)
store['单线索成交率%'] = (store['成交台数']/store['线索量']*100).round(2)
fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.bar(store.index, store['成交台数'], color='#55A868')
ax.bar_label(bars, padding=2)
ax.set_title('各门店成交台数与单线索成交率', fontsize=14)
ax.set_xticklabels(store.index, rotation=15, ha='right')
for i, idx in enumerate(store.index):
    ax.text(i, store.loc[idx, '成交台数']+2, f"{store.loc[idx, '单线索成交率%']:.1f}%", ha='center', fontsize=9, color='#DD8452')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, '12_门店成交与线索质量.png'), dpi=150)
plt.close()

print('所有图表已生成到:', OUT_DIR)
print('图表文件:', sorted(os.listdir(OUT_DIR)))
