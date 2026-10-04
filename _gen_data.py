# -*- coding: utf-8 -*-
"""Generate the embedded dataset (data.js) for the BI dashboard from 销售数据分析.xlsx"""
import pandas as pd, json

SRC = r'D:\Python\Python_item\PythonProject\销售数据分析\销售数据分析.xlsx'
df = pd.read_excel(SRC)
df['日期'] = pd.to_datetime(df['日期'])

# ---- raw rows trimmed to what the dashboard needs ----
rows = []
for _, r in df.iterrows():
    rows.append({
        'd': r['日期'].strftime('%Y-%m-%d'),
        'c': r['城市'],
        's': r['门店名称'].replace('理想汽车', '').replace('店', ''),
        'a': r['销售顾问'],
        'v': r['车型'].replace('理想 ', ''),
        'sold': int(r['成交台数']),
        'lead': int(r['线索量']),
        'visit': int(r['到店人数']),
        'wk': r['星期'],
        'w': r['周标签'],
    })

# ---- per date trend ----
trend = (df.groupby('日期').agg(sold=('成交台数', 'sum'), lead=('线索量', 'sum'), visit=('到店人数', 'sum'))
         .reset_index().sort_values('日期'))
trend = [{'d': t.日期.strftime('%m-%d'), 'sold': int(t.sold), 'lead': int(t.lead), 'visit': int(t.visit)}
         for t in trend.itertuples()]

# ---- per week ----
weeks = (df.groupby('周标签').agg(sold=('成交台数', 'sum'), lead=('线索量', 'sum'),
         visit=('到店人数', 'sum'), days=('日期', 'nunique')).reset_index())
weeks = [{'w': t.周标签, 'sold': int(t.sold), 'lead': int(t.lead), 'visit': int(t.visit), 'days': int(t.days)}
         for t in weeks.itertuples()]

# ---- week-of-day sold ----
wd = df.groupby('星期').agg(sold=('成交台数', 'sum')).reset_index()
order = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
wd['星期'] = pd.Categorical(wd['星期'], order, ordered=True)
wd = wd.sort_values('星期')
wd = [{'d': t.星期, 'sold': int(t.sold)} for t in wd.itertuples()]

# ---- city comparison ----
cities = df.groupby('城市').agg(sold=('成交台数', 'sum'), lead=('线索量', 'sum'),
          visit=('到店人数', 'sum'), adv=('销售顾问', 'nunique')).reset_index()
store_map = df.groupby('城市').agg(store=('门店名称', 'first')).reset_index().set_index('城市')['store'].to_dict()
cities = [{'c': t.城市, 'store': store_map[t.城市].replace('理想汽车', '').replace('店', ''),
           'sold': int(t.sold), 'lead': int(t.lead), 'visit': int(t.visit), 'adv': int(t.adv)}
          for t in cities.itertuples()]

# ---- city x vehicle sold (heatmap) ----
cv_piv = df.groupby(['城市', '车型'])['成交台数'].sum().unstack(fill_value=0)
cvCols = [v.replace('理想 ', '') for v in cv_piv.columns]
cvList = [[c] + [int(cv_piv.loc[c, v]) for v in cv_piv.columns] for c in cv_piv.index]

# ---- vehicle comparison ----
veh = df.groupby('车型').agg(sold=('成交台数', 'sum'), lead=('线索量', 'sum'),
       visit=('到店人数', 'sum')).reset_index()
veh = [{'v': t.车型.replace('理想 ', ''), 'sold': int(t.sold), 'lead': int(t.lead),
        'visit': int(t.visit), 'lr': round(100*t.sold/t.lead, 1), 'vr': round(100*t.sold/t.visit, 1)}
       for t in veh.itertuples()]

# ---- advisor comparison ----
adv = df.groupby('销售顾问').agg(sold=('成交台数', 'sum'), lead=('线索量', 'sum'),
       visit=('到店人数', 'sum'), cities=('城市', 'nunique')).reset_index()
adv['citynames'] = df.groupby('销售顾问')['城市'].apply(lambda s: '、'.join(sorted(set(s)))).reset_index(drop=True)
adv = [{'a': t.销售顾问, 'sold': int(t.sold), 'lead': int(t.lead), 'visit': int(t.visit),
        'lr': round(100*t.sold/t.lead, 1) if t.lead else 0.0,
        'vr': round(100*t.sold/t.visit, 1) if t.visit else 0.0,
        'cities': t.citynames, 'nc': int(t.cities), 'daily': round(t.sold/28, 2),
        'store': df.loc[df['销售顾问'] == t.销售顾问, '城市'].map(lambda x: store_map[x].replace('理想汽车', '').replace('店', '')).iloc[0]}
       for t in adv.itertuples()]
adv.sort(key=lambda x: x['sold'], reverse=True)

# ---- pivot: week x vehicle ----
wv = df.groupby(['周标签', '车型'])['成交台数'].sum().unstack(fill_value=0)
wvRows = [[w] + [int(wv.loc[w, v]) for v in wv.columns] for w in wv.index]
wvCols = [v.replace('理想 ', '') for v in wv.columns]

# TODO list of advisors' vehicles etc not necessary

out = {
    'totals': {'sold': int(df['成交台数'].sum()), 'lead': int(df['线索量'].sum()),
               'visit': int(df['到店人数'].sum()), 'days': int(df['日期'].nunique())},
    'rows': rows,
    'trend': trend, 'weeks': weeks, 'weekDay': wd, 'cities': cities,
    'cv': cvList, 'cvCols': cvCols, 'vehicles': veh, 'advisors': adv,
    'wv': wvRows, 'wvCols': wvCols,
    'advisorList': sorted(df['销售顾问'].unique().tolist()),
    'vehicleList': [v.replace('理想 ', '') for v in ['理想 L7', '理想 L8', '理想 L9', '理想 L6', '理想 MEGA']],
    'cityList': sorted(df['城市'].unique().tolist()),
}
with open(r'D:\Python\Python_item\PythonProject\销售数据分析\_BiData.js', 'w', encoding='utf-8') as f:
    f.write('window.BIDATA=' + json.dumps(out, ensure_ascii=False) + ';')
print('rows', len(rows))
print('OK totals', out['totals'])
print('cities', out['cities'])
