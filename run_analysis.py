import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import numpy as np

df    = pd.read_excel(r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx', sheet_name='dataset')
areas = pd.read_excel(r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx', sheet_name='areas')
df    = df.merge(areas, on='AreaId', how='left')

ts_cols = ['OrderDate','ArriveDate','DriverCandidateAssignedDate','PickedDate','StartPrepare','RejectedDateByRestaurant']
for c in ts_cols:
    if c in df.columns:
        df[c] = pd.to_datetime(df[c], errors='coerce')

df['Month'] = df['OrderDate'].dt.to_period('M')
df['Hour']  = df['OrderDate'].dt.hour
total = len(df)

print("=== Q1-A: כמות הזמנות ===")
print(f"סה\"כ הזמנות: {total}")
print(f"תקופה: {df['OrderDate'].min().date()} עד {df['OrderDate'].max().date()}")
print()
print("פיזור חודשי:")
for m, cnt in df.groupby('Month').size().items():
    print(f"  {m}: {cnt} ({cnt/total*100:.1f}%)")
print()
print("פיזור אזורי:")
area_counts = df.groupby('AreaName').size().sort_values(ascending=False)
for area, cnt in area_counts.items():
    print(f"  {area}: {cnt} ({cnt/total*100:.1f}%)")

print()
print("=== Q1-B: ערך הזמנה ותשלום ===")
desc = df['Price'].describe(percentiles=[.25,.5,.75])
print(f"ממוצע: {desc['mean']:.2f}")
print(f"חציון: {desc['50%']:.2f}")
print(f"סטיית תקן: {desc['std']:.2f}")
print(f"מינימום: {desc['min']:.2f}")
print(f"מקסימום: {desc['max']:.2f}")
print(f"P75: {desc['75%']:.2f}")
print(f"מעל 1000 ש\"ח: {(df['Price'] > 1000).sum()}")
print()
pm = df['PaymentMethod'].value_counts().sort_index()
print("PaymentMethod value_counts:")
for k, v in pm.items():
    label = "מזומן" if k == 0 else ("אשראי" if k == 1 else "אחר")
    print(f"  {k} ({label}): {v} ({v/total*100:.1f}%)")

print()
print("=== Q1-C: זמן לשיוך נהג ===")
assigned = df[df['DriverCandidateAssignedDate'].notna()].copy()
assigned['time_to_assign_min'] = (assigned['DriverCandidateAssignedDate'] - assigned['OrderDate']).dt.total_seconds() / 60
print(f"הזמנות עם נהג: {len(assigned)} מתוך {total} ({len(assigned)/total*100:.1f}%)")
print(f"ללא נהג: {total - len(assigned)} ({(total-len(assigned))/total*100:.1f}%)")
stats = assigned['time_to_assign_min'].describe(percentiles=[.25,.5,.75,.9])
print(f"ממוצע: {stats['mean']:.1f} דקות")
print(f"חציון (P50): {stats['50%']:.1f} דקות")
print(f"P25: {stats['25%']:.1f} דקות")
print(f"P75: {stats['75%']:.1f} דקות")
print(f"P90: {stats['90%']:.1f} דקות")
print(f"מקסימום: {stats['max']:.1f} דקות")

print()
print("=== Q1-D: ממצא חריג ===")
no_arrive           = df['ArriveDate'].isna().sum()
no_driver_no_arrive = df[(df['ArriveDate'].isna()) & (df['DriverCandidateAssignedDate'].isna())].shape[0]
driver_no_arrive    = df[(df['ArriveDate'].isna()) & (df['DriverCandidateAssignedDate'].notna())].shape[0]
rejected            = df['RejectedDateByRestaurant'].notna().sum()
print(f"ללא ArriveDate (לא נמסרו): {no_arrive} ({no_arrive/total*100:.1f}%)")
print(f"  מתוכן ללא נהג כלל: {no_driver_no_arrive} ({no_driver_no_arrive/no_arrive*100:.1f}%)")
print(f"  מתוכן קיבלו נהג אך לא נמסרו: {driver_no_arrive}")
print(f"נדחו ע\"י מסעדה: {rejected} ({rejected/total*100:.1f}%)")

print()
print("=== Q2: KPI מצב נוכחי ===")
delivered = df['ArriveDate'].notna()
print(f"מדד 1 - שיעור הצלחת משלוח: {delivered.mean()*100:.1f}%")
print(f"מדד 2 - זמן שיוך נהג (חציון): {assigned['time_to_assign_min'].median():.1f} דקות")
print(f"מדד 2 - זמן שיוך נהג (ממוצע): {assigned['time_to_assign_min'].mean():.1f} דקות")
print(f"מדד 3 - AOV: {df['Price'].mean():.2f} ש\"ח")
print(f"מדד 4 - שיעור דחייה מסעדה: {df['RejectedDateByRestaurant'].notna().mean()*100:.1f}%")
print(f"מדד 5 - שיעור כיסוי נהגים: {df['DriverCandidateAssignedDate'].notna().mean()*100:.1f}%")
