import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import numpy as np

df    = pd.read_excel(r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx', sheet_name='dataset')
areas = pd.read_excel(r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx', sheet_name='areas')
df    = df.merge(areas, on='AreaId', how='left')

ts_cols = ['OrderDate','ArriveDate','DriverCandidateAssignedDate','RejectedDateByRestaurant']
for c in ts_cols:
    df[c] = pd.to_datetime(df[c], errors='coerce')

df['delivered'] = df['ArriveDate'].notna()
df['Month']     = df['OrderDate'].dt.to_period('M')

total_revenue = df['Price'].sum()

# ══════════════════════════════════
# Q5-A: הכנסות לפי אזור
# ══════════════════════════════════
print("=== Q5-A: הכנסות לפי אזור ===")
area_rev = df.groupby('AreaName').agg(
    revenue=('Price','sum'),
    orders=('Id','count'),
    delivered=('delivered','sum'),
).assign(
    pct_revenue=lambda x: (x['revenue']/total_revenue*100).round(1),
    avg_order=lambda x: (x['revenue']/x['orders']).round(2),
    delivery_rate=lambda x: (x['delivered']/x['orders']*100).round(1),
).sort_values('revenue', ascending=False)

print(f"סה\"כ הכנסות: ₪{total_revenue:,.0f}")
print()
for area, row in area_rev.iterrows():
    print(f"  {area}: ₪{row['revenue']:,.0f} ({row['pct_revenue']}%) | {int(row['orders'])} הזמנות | AOV ₪{row['avg_order']} | מסירה {row['delivery_rate']}%")

# ══════════════════════════════════
# Q5-B: האם אזורי הכנסה גבוהה = מסירה טובה?
# ══════════════════════════════════
print()
print("=== Q5-B: קורלציה הכנסה vs מסירה ===")
top5_rev  = area_rev.head(5).index.tolist()
top5_del  = area_rev.sort_values('delivery_rate', ascending=False).head(5).index.tolist()
overlap   = set(top5_rev) & set(top5_del)
print(f"TOP 5 הכנסה: {top5_rev}")
print(f"TOP 5 מסירה: {top5_del}")
print(f"חפיפה: {overlap}")
corr = area_rev[['revenue','delivery_rate']].corr().iloc[0,1]
print(f"מתאם פירסון הכנסה-מסירה: {corr:.3f}")

# ══════════════════════════════════
# Q5-C: הכנסות לפי עסק (BusinessId)
# ══════════════════════════════════
print()
print("=== Q5-C: הכנסות לפי עסק (TOP 20) ===")
biz_rev = df.groupby('BusinessId').agg(
    revenue=('Price','sum'),
    orders=('Id','count'),
).assign(pct=lambda x: (x['revenue']/total_revenue*100).round(2))\
 .sort_values('revenue', ascending=False)

print(f"סה\"כ עסקים: {len(biz_rev)}")
print()
for rid, row in biz_rev.head(20).iterrows():
    print(f"  עסק {int(rid)}: ₪{row['revenue']:,.0f} ({row['pct']}%) | {int(row['orders'])} הזמנות")

# תלות TOP 3
top3_rev  = biz_rev.head(3)['revenue'].sum()
top10_rev = biz_rev.head(10)['revenue'].sum()
print()
print(f"TOP 3 עסקים: ₪{top3_rev:,.0f} = {top3_rev/total_revenue*100:.1f}% מסה\"כ הכנסות")
print(f"TOP 10 עסקים: ₪{top10_rev:,.0f} = {top10_rev/total_revenue*100:.1f}% מסה\"כ הכנסות")

# ══════════════════════════════════
# Q6-A: לקוחות חוזרים לפי אזור
# ══════════════════════════════════
print()
print("=== Q6-A: לקוחות חוזרים לפי אזור ===")
df_sorted = df.sort_values('OrderDate')
df_sorted['is_first'] = ~df_sorted.duplicated(subset='UserId', keep='first')
df_sorted['is_returning'] = ~df_sorted['is_first']

total_returning = df_sorted['is_returning'].mean() * 100
print(f"שיעור הזמנות מלקוחות חוזרים (כלל הפלטפורמה): {total_returning:.1f}%")
print()

area_ret = df_sorted.groupby('AreaName').agg(
    orders=('Id','count'),
    returning=('is_returning','sum'),
).assign(ret_rate=lambda x: (x['returning']/x['orders']*100).round(1))\
 .sort_values('ret_rate', ascending=False)

for area, row in area_ret.iterrows():
    print(f"  {area}: {int(row['returning'])}/{int(row['orders'])} הזמנות חוזרות = {row['ret_rate']}%")

# ══════════════════════════════════
# Q6-B: קוהורט מרץ 2024
# ══════════════════════════════════
print()
print("=== Q6-B: קוהורט מרץ 2024 ===")
march_first = df_sorted[
    (df_sorted['is_first']) &
    (df_sorted['Month'] == pd.Period('2024-03'))
]['UserId'].unique()

print(f"לקוחות שהזמנתם הראשונה היתה במרץ 2024: {len(march_first)}")

# כל ההזמנות של לקוחות אלו
cohort_df = df[df['UserId'].isin(march_first)].copy()
first_orders = cohort_df.groupby('UserId')['OrderDate'].min().reset_index()
first_orders.columns = ['UserId', 'first_date']
cohort_df = cohort_df.merge(first_orders, on='UserId')
cohort_df['days_since_first'] = (cohort_df['OrderDate'] - cohort_df['first_date']).dt.days

# 30 ימים
repeat_30 = cohort_df[
    (cohort_df['days_since_first'] > 0) &
    (cohort_df['days_since_first'] <= 30)
]['UserId'].nunique()
# 60 ימים
repeat_60 = cohort_df[
    (cohort_df['days_since_first'] > 0) &
    (cohort_df['days_since_first'] <= 60)
]['UserId'].nunique()

print(f"הזמינו שוב תוך 30 יום: {repeat_30} ({repeat_30/len(march_first)*100:.1f}%)")
print(f"הזמינו שוב תוך 60 יום: {repeat_60} ({repeat_60/len(march_first)*100:.1f}%)")

# ══════════════════════════════════
# Q6-C: שיעור שימור 30 יום לפי אזור
# ══════════════════════════════════
print()
print("=== Q6-C: שימור 30 יום לפי אזור ===")
# הזמנה ראשונה של כל לקוח
customer_first = df_sorted[df_sorted['is_first']][['UserId','OrderDate','AreaName']].copy()
customer_first.columns = ['UserId','first_date','first_area']

all_orders = df[['UserId','OrderDate']].copy()
cohort_area = customer_first.merge(all_orders, on='UserId')
cohort_area['days'] = (cohort_area['OrderDate'] - cohort_area['first_date']).dt.days

retained_30 = cohort_area[
    (cohort_area['days'] > 0) & (cohort_area['days'] <= 30)
].groupby('first_area')['UserId'].nunique().reset_index()
retained_30.columns = ['AreaName','retained_30']

total_cust = customer_first.groupby('first_area')['UserId'].nunique().reset_index()
total_cust.columns = ['AreaName','total_customers']

ret_by_area = total_cust.merge(retained_30, on='AreaName', how='left').fillna(0)
ret_by_area['retention_30'] = (ret_by_area['retained_30']/ret_by_area['total_customers']*100).round(1)
ret_by_area = ret_by_area.sort_values('retention_30', ascending=False)

for _, row in ret_by_area.iterrows():
    print(f"  {row['AreaName']}: {row['retention_30']}% ({int(row['retained_30'])}/{int(row['total_customers'])} לקוחות)")

best  = ret_by_area.iloc[0]
worst = ret_by_area.iloc[-1]
print(f"\nהטוב ביותר: {best['AreaName']} ({best['retention_30']}%)")
print(f"הגרוע ביותר: {worst['AreaName']} ({worst['retention_30']}%)")
