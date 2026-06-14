import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import numpy as np

df    = pd.read_excel(r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx', sheet_name='dataset')
areas = pd.read_excel(r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx', sheet_name='areas')
df    = df.merge(areas, on='AreaId', how='left')

ts_cols = ['OrderDate','ArriveDate','DriverCandidateAssignedDate','PickedDate','StartPrepare','RejectedDateByRestaurant']
for c in ts_cols:
    df[c] = pd.to_datetime(df[c], errors='coerce')

df['Month']     = df['OrderDate'].dt.to_period('M')
df['Hour']      = df['OrderDate'].dt.hour
df['delivered'] = df['ArriveDate'].notna()
total = len(df)

# ══════════════════════════════════
# Q3-A: שיעור מסירה לפי אזור וחודש
# ══════════════════════════════════
print("=== Q3-A: שיעור מסירה לפי אזור ===")
area_stats = df.groupby('AreaName').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
).assign(rate=lambda x: (x['delivered']/x['orders']*100).round(1))
area_stats = area_stats.sort_values('rate', ascending=False)
for area, row in area_stats.iterrows():
    print(f"  {area}: {int(row['orders'])} הזמנות | {int(row['delivered'])} נמסרו | {row['rate']}%")

print()
print("=== Q3-A: שיעור מסירה לפי חודש ===")
monthly = df.groupby('Month').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
).assign(rate=lambda x: (x['delivered']/x['orders']*100).round(1))
for m, row in monthly.iterrows():
    print(f"  {m}: {int(row['orders'])} הזמנות | {int(row['delivered'])} נמסרו | {row['rate']}%")

# ══════════════════════════════════
# Q3-B: נפח ושיעור מסירה לפי שעה
# ══════════════════════════════════
print()
print("=== Q3-B: נפח ושיעור מסירה לפי שעה ===")
hourly = df.groupby('Hour').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
).assign(rate=lambda x: (x['delivered']/x['orders']*100).round(1))
for h, row in hourly.iterrows():
    bar = '█' * int(row['rate'] / 5)
    print(f"  {h:02d}:00 | {int(row['orders']):4d} הזמנות | {row['rate']:5.1f}% {bar}")

print()
print("=== Q3-B: סיכום שעות ===")
peak_orders = hourly['orders'].idxmax()
worst_rate  = hourly['rate'].idxmin()
best_rate   = hourly['rate'].idxmax()
print(f"שעת שיא נפח: {peak_orders}:00 ({int(hourly.loc[peak_orders,'orders'])} הזמנות)")
print(f"שעת שיא מסירה (הטובה): {best_rate}:00 ({hourly.loc[best_rate,'rate']}%)")
print(f"שעה הגרועה במסירה: {worst_rate}:00 ({hourly.loc[worst_rate,'rate']}%)")

# ══════════════════════════════════
# Q3-C: השוואת אזורים
# ══════════════════════════════════
print()
print("=== Q3-C: השוואת אזורים — מיון לפי שיעור מסירה ===")
overall_rate = df['delivered'].mean() * 100
print(f"ממוצע כלל-פלטפורמה: {overall_rate:.1f}%")
for area, row in area_stats.iterrows():
    diff = row['rate'] - overall_rate
    flag = '▲' if diff > 5 else ('▼' if diff < -5 else '~')
    print(f"  {flag} {area}: {row['rate']}% (פער מממוצע: {diff:+.1f}%)")

# ══════════════════════════════════
# Q4-A: דירוג נהגים לפי משלוחים
# ══════════════════════════════════
print()
print("=== Q4-A: דירוג נהגים לפי הזמנות שנמסרו ===")
driver_df = df[df['DriverId'].notna()].copy()
driver_stats = driver_df.groupby('DriverId').agg(
    total_assigned=('Id','count'),
    total_delivered=('delivered','sum'),
    first_order=('OrderDate','min'),
    last_order=('OrderDate','max'),
).reset_index()
driver_stats['active_days'] = (driver_stats['last_order'] - driver_stats['first_order']).dt.days + 1
driver_stats['delivery_rate'] = (driver_stats['total_delivered'] / driver_stats['total_assigned'] * 100).round(1)
driver_stats = driver_stats.sort_values('total_delivered', ascending=False)

print("TOP 10 נהגים (לפי הזמנות שנמסרו):")
for _, row in driver_stats.head(10).iterrows():
    print(f"  נהג {int(row['DriverId'])}: {int(row['total_delivered'])} נמסרו | {int(row['total_assigned'])} שובצו | {row['delivery_rate']}% | {int(row['active_days'])} ימים פעיל")

print()
print("BOTTOM 10 נהגים (לפי הזמנות שנמסרו):")
for _, row in driver_stats.tail(10).iterrows():
    print(f"  נהג {int(row['DriverId'])}: {int(row['total_delivered'])} נמסרו | {int(row['total_assigned'])} שובצו | {row['delivery_rate']}% | {int(row['active_days'])} ימים פעיל")

# ══════════════════════════════════
# Q4-B: מדד יעילות מנורמל
# ══════════════════════════════════
print()
print("=== Q4-B: מדד יעילות — משלוחים מנורמלים ליום ===")
driver_stats['efficiency'] = (driver_stats['total_delivered'] / driver_stats['active_days']).round(3)
driver_stats = driver_stats.sort_values('efficiency', ascending=False)

print("TOP 10 יעילות:")
for _, row in driver_stats.head(10).iterrows():
    print(f"  נהג {int(row['DriverId'])}: {row['efficiency']} משלוחים/יום | {int(row['total_delivered'])} סה\"כ | {int(row['active_days'])} ימים")

print()
print("BOTTOM 10 יעילות:")
for _, row in driver_stats.tail(10).iterrows():
    print(f"  נהג {int(row['DriverId'])}: {row['efficiency']} משלוחים/יום | {int(row['total_delivered'])} סה\"כ | {int(row['active_days'])} ימים")

# ══════════════════════════════════
# Q4-C: התפלגות ואאוטליירים
# ══════════════════════════════════
print()
print("=== Q4-C: התפלגות מדד היעילות ===")
eff = driver_stats['efficiency']
print(f"מספר נהגים: {len(eff)}")
print(f"ממוצע: {eff.mean():.3f} משלוחים/יום")
print(f"חציון: {eff.median():.3f} משלוחים/יום")
print(f"P75: {eff.quantile(0.75):.3f}")
print(f"P90: {eff.quantile(0.90):.3f}")
print(f"P95: {eff.quantile(0.95):.3f}")
print(f"מקסימום: {eff.max():.3f}")
print(f"מינימום: {eff.min():.3f}")

# אאוטליירים: מעל Q3 + 3*IQR
Q1  = eff.quantile(0.25)
Q3  = eff.quantile(0.75)
IQR = Q3 - Q1
upper_fence = Q3 + 3 * IQR
outliers = driver_stats[driver_stats['efficiency'] > upper_fence]
print(f"\nגבול אאוטליירים (Q3+3*IQR): {upper_fence:.3f}")
print(f"מספר אאוטליירים: {len(outliers)}")
if len(outliers) > 0:
    for _, row in outliers.iterrows():
        print(f"  נהג {int(row['DriverId'])}: {row['efficiency']:.3f} משלוחים/יום | {int(row['total_delivered'])} נמסרו | {int(row['active_days'])} ימים פעיל")

print()
print("=== הפצה לפי טווחים ===")
bins = [0, 1, 2, 3, 5, 10, 100]
labels = ['0-1','1-2','2-3','3-5','5-10','10+']
driver_stats['eff_bin'] = pd.cut(driver_stats['efficiency'], bins=bins, labels=labels)
print(driver_stats['eff_bin'].value_counts().sort_index().to_string())
