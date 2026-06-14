import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import numpy as np

PATH  = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
df    = pd.read_excel(PATH, sheet_name='dataset')
areas = pd.read_excel(PATH, sheet_name='areas')
pm    = pd.read_excel(PATH, sheet_name='payment_methods_user')
df    = df.merge(areas, on='AreaId', how='left')

for c in ['OrderDate','ArriveDate','DriverCandidateAssignedDate','RejectedDateByRestaurant']:
    df[c] = pd.to_datetime(df[c], errors='coerce')

df['delivered'] = df['ArriveDate'].notna()
df['Week']      = df['OrderDate'].dt.to_period('W').astype(str)
df['Month']     = df['OrderDate'].dt.to_period('M').astype(str)
df['DayOfWeek'] = df['OrderDate'].dt.dayofweek   # 0=Monday
df['Hour']      = df['OrderDate'].dt.hour
df['Date']      = df['OrderDate'].dt.date

# ══════════════════════════════════════════════════════════════════
# Q7 — ORDER PATTERNS OVER TIME
# ══════════════════════════════════════════════════════════════════
print("=" * 60)
print("Q7 — ORDER PATTERNS OVER TIME")
print("=" * 60)

# ── יומי ──────────────────────────────────────────────────────────
daily = df.groupby('Date').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
    revenue=('Price','sum'),
).assign(
    delivery_rate=lambda x: (x['delivered']/x['orders']*100).round(1),
    aov=lambda x: (x['revenue']/x['orders']).round(2),
)
print("\n--- נפח יומי (סטטיסטיקה) ---")
print(f"ממוצע הזמנות ליום: {daily['orders'].mean():.1f}")
print(f"חציון:              {daily['orders'].median():.1f}")
print(f"מקסימום:            {daily['orders'].max()} ({daily['orders'].idxmax()})")
print(f"מינימום:            {daily['orders'].min()} ({daily['orders'].idxmin()})")
print(f"סטיית תקן:          {daily['orders'].std():.1f}")

# ימים חריגים — מעל ממוצע + 2STD
threshold = daily['orders'].mean() + 2 * daily['orders'].std()
spikes = daily[daily['orders'] > threshold].sort_values('orders', ascending=False)
print(f"\nימים חריגים (מעל ממוצע+2STD = {threshold:.0f}):")
for date, row in spikes.iterrows():
    print(f"  {date}: {int(row['orders'])} הזמנות | AOV ₪{row['aov']:.2f} | מסירה {row['delivery_rate']}%")

drops = daily[daily['orders'] < daily['orders'].mean() - 2*daily['orders'].std()]
print(f"\nימים שפל (מתחת לממוצע-2STD):")
for date, row in drops.iterrows():
    print(f"  {date}: {int(row['orders'])} הזמנות | מסירה {row['delivery_rate']}%")

# ── שבועי ─────────────────────────────────────────────────────────
print("\n--- נפח שבועי ---")
weekly = df.groupby('Week').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
    revenue=('Price','sum'),
).assign(
    delivery_rate=lambda x: (x['delivered']/x['orders']*100).round(1),
    aov=lambda x: (x['revenue']/x['orders']).round(2),
)
for w, row in weekly.iterrows():
    print(f"  {w}: {int(row['orders'])} הזמנות | AOV ₪{row['aov']:.2f} | מסירה {row['delivery_rate']}%")

# שינוי שבוע-על-שבוע
weekly['wow_pct'] = weekly['orders'].pct_change() * 100
print("\nשינוי שבועי (WoW %):")
for w, row in weekly.iterrows():
    if not np.isnan(row['wow_pct']):
        arrow = '▲' if row['wow_pct'] > 0 else '▼'
        print(f"  {w}: {arrow} {row['wow_pct']:+.1f}%")

# ── יום בשבוע ─────────────────────────────────────────────────────
print("\n--- דפוס יום בשבוע ---")
days_heb = ['שני','שלישי','רביעי','חמישי','שישי','שבת','ראשון']
dow = df.groupby('DayOfWeek').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
    revenue=('Price','sum'),
).assign(
    delivery_rate=lambda x: (x['delivered']/x['orders']*100).round(1),
    aov=lambda x: (x['revenue']/x['orders']).round(2),
    daily_avg=lambda x: x['orders'] / (df['Date'].nunique() / 7),
)
for dow_i, row in dow.iterrows():
    print(f"  {days_heb[dow_i]}: {int(row['orders'])} הזמנות סה\"כ | ממוצע {row['daily_avg']:.1f}/יום | AOV ₪{row['aov']:.2f} | מסירה {row['delivery_rate']}%")

# ── AOV לאורך זמן ─────────────────────────────────────────────────
print("\n--- שינוי AOV לאורך זמן (שבועי) ---")
for w, row in weekly.iterrows():
    print(f"  {w}: AOV ₪{row['aov']:.2f}")

# ── שעת שיא לפי חודש ──────────────────────────────────────────────
print("\n--- שעת שיא לפי חודש ---")
for month in df['Month'].unique():
    mdf    = df[df['Month'] == month]
    peak_h = mdf.groupby('Hour').size().idxmax()
    peak_v = mdf.groupby('Hour').size().max()
    del_at_peak = mdf[mdf['Hour']==peak_h]['delivered'].mean()*100
    print(f"  {month}: שעת שיא {peak_h}:00 ({peak_v} הזמנות) | מסירה בשיא {del_at_peak:.1f}%")

# ══════════════════════════════════════════════════════════════════
# Q8 — DATA QUALITY CHECK
# ══════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("Q8 — DATA QUALITY CHECK: payment_methods_user")
print("=" * 60)

print(f"\n--- מבנה הטבלה ---")
print(f"עמודות: {pm.columns.tolist()}")
print(f"שורות:  {len(pm)}")
print(f"סוגי עמודות:\n{pm.dtypes}")
print(f"\nדוגמת שורות ראשונות:")
print(pm.head(10).to_string())

print(f"\n--- בדיקת ערכים חסרים ---")
print(pm.isnull().sum().to_string())

print(f"\n--- התפלגות שיטות תשלום ---")
if 'PaymentMethod' in pm.columns:
    print(pm['PaymentMethod'].value_counts().to_string())
elif 'payment_method' in pm.columns:
    print(pm['payment_method'].value_counts().to_string())
else:
    for col in pm.columns:
        if 'pay' in col.lower() or 'method' in col.lower():
            print(f"{col}: {pm[col].value_counts().to_string()}")

print(f"\n--- לקוחות ייחודיים בטבלת payment ---")
user_col = [c for c in pm.columns if 'user' in c.lower() or 'id' in c.lower()][0]
print(f"עמודת משתמש: {user_col}")
n_users_pm = pm[user_col].nunique()
print(f"לקוחות ייחודיים: {n_users_pm}")

# לקוח עם כמה שיטות תשלום?
methods_per_user = pm.groupby(user_col).size()
print(f"\nשיטות תשלום לכל לקוח:")
print(methods_per_user.value_counts().sort_index().to_string())
multi = methods_per_user[methods_per_user > 1]
print(f"\nלקוחות עם >1 שיטת תשלום: {len(multi)}")
if len(multi) > 0:
    print(multi.head(10).to_string())

# ── השוואה מול טבלת הזמנות ────────────────────────────────────────
print(f"\n--- השוואה מול טבלת הזמנות ---")
n_users_orders = df['UserId'].nunique()
print(f"לקוחות ייחודיים בטבלת הזמנות: {n_users_orders}")
print(f"לקוחות ייחודיים בטבלת payment: {n_users_pm}")

users_pm_set  = set(pm[user_col].dropna())
users_ord_set = set(df['UserId'].dropna())

only_in_pm     = users_pm_set - users_ord_set
only_in_orders = users_ord_set - users_pm_set
in_both        = users_pm_set & users_ord_set

print(f"\nלקוחות בשתי הטבלאות:  {len(in_both)}")
print(f"רק ב-payment:          {len(only_in_pm)}")
print(f"רק בהזמנות:            {len(only_in_orders)}")

# ── השוואת שיטת תשלום בין הטבלאות ────────────────────────────────
print(f"\n--- השוואת PaymentMethod: orders vs payment_methods_user ---")
print("בטבלת הזמנות:")
print(df['PaymentMethod'].value_counts().to_string())

# cross reference
pm_col = [c for c in pm.columns if 'pay' in c.lower() or 'method' in c.lower()][0]
print(f"\nבטבלת payment_methods_user (עמודה: {pm_col}):")
print(pm[pm_col].value_counts().to_string())

# לקוחות שבהזמנות משלמים במזומן אבל בpm רשומים כאשראי ולהפך
merged_check = df[['UserId','PaymentMethod']].drop_duplicates().merge(
    pm[[user_col, pm_col]].rename(columns={user_col:'UserId', pm_col:'pm_method'}),
    on='UserId', how='inner'
)
mismatch = merged_check[merged_check['PaymentMethod'] != merged_check['pm_method']]
print(f"\nאי-התאמות PaymentMethod בין הטבלאות: {len(mismatch)}")
if len(mismatch) > 0:
    print("דוגמאות:")
    print(mismatch.head(10).to_string())
    print(f"\nהתפלגות אי-ההתאמות:")
    print(mismatch.groupby(['PaymentMethod','pm_method']).size().to_string())
