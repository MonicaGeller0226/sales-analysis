# 销售数据分析 - BI 看板

基于 `销售数据分析.xlsx` 生成的交互式销售看板。

## 文件结构

| 文件 | 角色 | 说明 |
|---|---|---|
| `销售数据分析.xlsx` | 源数据 | 原始 Excel 数据（唯一数据来源） |
| `_BiData.js` | 数据层 | 供页面读取，含 `window.BIDATA`（原始 rows + 预聚合结果） |
| `销售数据BI看板.html` | 展示层 | 静态页面，纯前端逐行计算实现真联动筛选；引用 Chart.js 4 CDN 与本地 `_BiData.js` |
| `_gen_data.py` | 生成脚本 | 从 xlsx 提取/汇总并导出 `_BiData.js` |
| `销售数据分析报告.md` | 分析文档 | 结论文档，看板的结论依据 |
| `README.md` | 本说明 | 使用方法 |

依赖链:

```
销售数据分析.xlsx → _gen_data.py → _BiData.js → 销售数据BI看板.html (浏览器)
```

## 数据改动后如何更新页面

唯一需要做的是**重跑生成脚本让 `_BiData.js` 重新导出**，页面本身不用改：

```powershell
D:\Python\miniconda3\python.exe "D:\Python\Python_item\PythonProject\销售数据分析\_gen_data.py"
```

完成后**刷新浏览器里的看板**（建议 Ctrl+F5 强刷，避免缓存旧 `_BiData.js`）。

前提:
- 不改 xlsx 的列结构（表头/列含义/日期、城市、车型、顾问、成交、线索、到店、周等字段）——脚本依赖固定映射。
- 每次以同文件名、同路径覆盖保存 xlsx，先关闭占用该文件的程序（如别在 Excel 里锁着文件）再跑脚本。
- 若日后要改 xlsx 结构（加列/换列/改口径），需同步改 `_gen_data.py` 的字段映射与上层口径——属扩展改动，先确认再动手。

> 注：本机 `python` 命令会被 Microsoft Store 占位符拦截，一律用完整路径 `D:\Python\miniconda3\python.exe`。

## 数据口径（总览，用于核对）

- 总体（7月）：成交 490 台 / 线索 3692 / 到店 11263
- 车型：L7=241 · L8=186 · L9=49 · L6=14 · MEGA=0
- 城市 TOP：上海133 · 北京104 · 深圳56 · 成都54 · 广州49 · 杭州47 · 重庆47
- 顾问 TOP：王浩91 · 张磊83 · 黄杰56 · 周强54 · 林宇49
- 时间范围：2026-07-01 ~ 07-28

## `_BiData.js` 对象结构

`window.BIDATA`:
- `rows`：553 条逐行记录，字段 `d,c,s,a,v,sold,lead,visit,wk,w`
  - d=日期 c=城市 s=销售顾问 a=负责人 v=车型 sold=成交 lead=线索 visit=到店 wk=第几周 w=月
- 预聚合：`totals / weeks / trend / cityList / vehicleList / advisorList / cities / vehicles / advisors`

## 相关修复记录

- **顾问排行条等长 Bug（2026-09-06）**：根因是 `advList()` 生成内联样式时漏了 `%` 单位（输出 `width:91` 无效，浏览器回退满格），仅在该行补上 `'%'` 修复，未动其它代码。详见 `memory/2026-09-06.md`。
