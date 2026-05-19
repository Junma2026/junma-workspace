"""
2024年报损报表 vAI
按名称模糊合并 + 金额列
输出：2024年药房报损报表_AI.xlsx
"""
import pandas as pd, warnings, os
warnings.filterwarnings('ignore')

base = r"C:\Users\Administrator\Desktop\药品进销存AI取数"
out_path = os.path.join(base, "2024年药房报损报表_AI.xlsx")

df23 = pd.read_excel(os.path.join(base, "2023年12月末盘点表.xlsx"))
df24 = pd.read_excel(os.path.join(base, "2024年12月末盘点表.xlsx"))
df_pur = pd.read_excel(os.path.join(base, "2024年药品采购单.xls"))
df_sale = pd.read_excel(os.path.join(base, "2024年药品销售.xlsx"))

def clean(s):
    s = str(s).strip()
    for old, new in [('\uff08','('),('\uff09',')'),('（','('),('）',')'),
                     ('×','*'),('X','*'),('x','*'),(' ',''),('\u3000',''),
                     ('\u2013','-'),('~','-'),('—','-'),('：',':')]:
        s = s.replace(old, new)
    return s

for df in [df23, df24, df_pur]:
    df['c_name'] = df['药品名称'].apply(clean)
df_sale['c_name'] = df_sale['药品名称'].apply(clean)

# ================================================================
# 名称模糊映射
# ================================================================
all_names = set()
for df in [df23, df24, df_pur, df_sale]:
    all_names.update(df['c_name'])
all_names = sorted(all_names)

name_map = {n: n for n in all_names}

# 手动名称合并规则
merges = {
    '富马酸喹硫平片(舒思)': '富马酸喹硫平片',
    '富马酸喹硫平片(太伦佐)': '富马酸喹硫平片',
    '布洛芬缓释胶囊(盖芬)': '布洛芬缓释胶囊',
    '丙戊酸镁缓释片(神泰)': '丙戊酸镁缓释片',
    '右佐匹克隆片(文飞)': '右佐匹克隆片',
    '辛伐他汀片(京必舒新)': '辛伐他汀片',
    '多潘立酮片(吗丁啉)': '多潘立酮片',
    '多潘立酮片(快克啉)': '多潘立酮片',
    '盐酸苯海索片(安坦)': '盐酸苯海索片',
    '盐酸吗啉胍片(病毒灵)': '盐酸吗啉胍片',
    '盐酸左氧氟沙星片(左沙)': '盐酸左氧氟沙星片',
    '氯普噻吨片(泰尔登)': '氯普噻吨片',
    '盐酸哌罗匹隆片(康尔汀)': '哌罗匹隆片',
    '利培酮片(可同)': '利培酮片',
    '长春西汀注射液(润坦)': '长春西汀注射液',
    '奥氮平片(欧兰宁)': '奥氮平片',
    '苯磺酸氨氯地平片(络活喜)': '苯磺酸氨氯地平片',
    '苯磺酸左氨氯地平片(左沙)': '苯磺酸左氨氯地平片',
    '盐酸齐拉西酮胶囊(思贝格)': '盐酸齐拉西酮胶囊',
    '酚磺乙胺注射液(止血敏)': '酚磺乙胺注射液',
    '消旋山莨菪碱片(654-2)': '消旋山莨菪碱片',
    '阿莫西林胶囊(先锋4)': '阿莫西林胶囊',  # 先锋4就是头孢氨苄，不是阿莫西林
}
for k, v in merges.items():
    name_map[k] = v

for df in [df23, df24, df_pur, df_sale]:
    df['g_name'] = df['c_name'].map(name_map)

# ================================================================
# 按组汇总
# ================================================================
print("=" * 60)
print("汇总计算")
print("=" * 60)

# 2023盘点
s23 = df23.groupby('g_name', as_index=False).agg(
    q23=('数量','sum'),
    amt23=('金额','sum'),
    names=('药品名称', lambda x: '|'.join(sorted(set(x)))),
    specs=('规格', lambda x: '|'.join(x.unique())),
    units=('单位', lambda x: '|'.join(x.unique())),
    codes=('药品编码', lambda x: ','.join(sorted(set(x.astype(str)))))
)
# 2024盘点
s24 = df24.groupby('g_name', as_index=False).agg(
    q24=('数量','sum'),
    names24=('药品名称', lambda x: '|'.join(sorted(set(x)))),
    specs24=('规格', lambda x: '|'.join(x.unique())),
    units24=('单位', lambda x: '|'.join(x.unique())),
    codes24=('药品编码', lambda x: ','.join(sorted(set(x.astype(str)))))
)
# 采购
pur = df_pur.groupby('g_name', as_index=False).agg(
    pur_qty=('数量','sum'),
    pur_amt=('金额','sum')
)
# 销售
sale = df_sale.groupby('g_name', as_index=False).agg(
    sale_qty=('数量','sum'),
    sale_amt=('金额','sum')
)

# 2024盘点单价
s24_price = df24.groupby('g_name', as_index=False)['单价'].first()
pur['pur_price'] = pur.apply(lambda r: round(r['pur_amt']/r['pur_qty'],4) if r['pur_qty']!=0 else 0, axis=1)
s23['price23'] = s23.apply(lambda r: round(r['amt23']/r['q23'],4) if r['q23']!=0 else 0, axis=1)

all_groups = set(s23['g_name']) | set(s24['g_name']) | set(pur['g_name']) | set(sale['g_name'])
print(f"药品总数: {len(all_groups)}")

rows = []
for g in sorted(all_groups):
    r23 = s23[s23['g_name']==g]
    r24 = s24[s24['g_name']==g]
    rp = pur[pur['g_name']==g]
    rs = sale[sale['g_name']==g]
    r24p = s24_price[s24_price['g_name']==g]
    
    q23v = int(r23['q23'].iloc[0]) if len(r23) > 0 else 0
    q24v = int(r24['q24'].iloc[0]) if len(r24) > 0 else 0
    pur_v = int(rp['pur_qty'].iloc[0]) if len(rp) > 0 else 0
    sale_v = int(rs['sale_qty'].iloc[0]) if len(rs) > 0 else 0
    
    # 单价选择
    pur_price_v = rp['pur_price'].iloc[0] if len(rp) > 0 else 0
    price23_v = r23['price23'].iloc[0] if len(r23) > 0 else 0
    price24_v = r24p['单价'].iloc[0] if len(r24p) > 0 else 0
    
    if pur_price_v != 0:
        final_price = pur_price_v
    elif price24_v != 0:
        final_price = price24_v
    elif price23_v != 0:
        final_price = price23_v
    else:
        final_price = 0
    
    loss_qty = q23v + pur_v - sale_v - q24v
    loss_amt = round(loss_qty * final_price, 2)
    
    codes = r23['codes'].iloc[0] if len(r23) > 0 else (r24['codes24'].iloc[0] if len(r24) > 0 else '')
    name_display = r23['names'].iloc[0] if len(r23) > 0 else (r24['names24'].iloc[0] if len(r24) > 0 else g)
    spec_display = r23['specs'].iloc[0] if len(r23) > 0 else (r24['specs24'].iloc[0] if len(r24) > 0 else '')
    unit_display = r23['units'].iloc[0] if len(r23) > 0 else (r24['units24'].iloc[0] if len(r24) > 0 else '')
    
    rows.append({
        '编码': codes,
        '药品名称': name_display,
        '规格': spec_display,
        '单位': unit_display,
        '报损数量': loss_qty,
        '报损金额': loss_amt
    })

df = pd.DataFrame(rows)
df = df.sort_values('报损数量')

print(f"盘亏: {len(df[df['报损数量']<0])} 种, 金额 {df[df['报损数量']<0]['报损金额'].sum():.2f}")
print(f"盘盈: {len(df[df['报损数量']>0])} 种, 金额 {df[df['报损数量']>0]['报损金额'].sum():.2f}")
print(f"持平: {len(df[df['报损数量']==0])} 种")
print(f"净报损: {df['报损金额'].sum():.2f}")

# ================================================================
# 导出
# ================================================================
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='2024年报损报表', index=False)
    for ws in writer.sheets.values():
        for col in ws.columns:
            max_len = 0
            letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[letter].width = min(max_len + 4, 50)

print(f"\n完成: {out_path}")
