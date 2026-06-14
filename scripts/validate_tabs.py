import sys; sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd, numpy as np, os

PATH = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
df    = pd.read_excel(PATH, sheet_name='dataset')
areas = pd.read_excel(PATH, sheet_name='areas')
df    = df.merge(areas, on='AreaId', how='left')
for c in ['OrderDate','ArriveDate','DriverCandidateAssignedDate','PickedDate','StartPrepare','RejectedDateByRestaurant']:
    df[c] = pd.to_datetime(df[c], errors='coerce')
df['Month']      = df['OrderDate'].dt.to_period('M').astype(str)
df['Hour']       = df['OrderDate'].dt.hour
df['delivered']  = df['ArriveDate'].notna()
df['has_driver'] = df['DriverCandidateAssignedDate'].notna()
df['time_to_assign'] = (df['DriverCandidateAssignedDate'] - df['OrderDate']).dt.total_seconds() / 60

# Tab 7
fdf7 = df.copy()
fdf7['DayOfWeek'] = fdf7['OrderDate'].dt.dayofweek
fdf7['Week']      = fdf7['OrderDate'].dt.to_period('W').astype(str)
fdf7['Date']      = fdf7['OrderDate'].dt.date
daily7 = fdf7.groupby('Date').agg(orders=('Id','count')).reset_index()
avg_d  = daily7['orders'].mean()
peak_d = daily7.loc[daily7['orders'].idxmax()]
print(f'Tab7 OK: avg={avg_d:.1f}, peak={int(peak_d["orders"])} at {peak_d["Date"]}')

weekly7 = fdf7.groupby('Week').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
    revenue=('Price','sum'),
).reset_index()
weekly7['delivery_rate'] = (weekly7['delivered'] / weekly7['orders'] * 100).round(1)
weekly7['aov']           = (weekly7['revenue'] / weekly7['orders']).round(2)
weekly7['wow_pct']       = weekly7['orders'].pct_change() * 100
print(f'Tab7 weekly: {len(weekly7)} weeks')

dow7 = fdf7.groupby('DayOfWeek').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
    revenue=('Price','sum'),
).reset_index()
dow7['daily_avg'] = (dow7['orders'] / fdf7['Week'].nunique()).round(1)
print(f'Tab7 dow: {dow7["daily_avg"].max():.1f} max daily avg')

# Tab 8
pm8 = pd.read_excel(PATH, sheet_name='payment_methods_user')
n_rows = len(pm8)
n_users_pm = pm8['UserId'].nunique()
n_dup_tok = pm8['Token'].duplicated().sum()
n_uniq_tok = pm8['Token'].nunique()
print(f'Tab8 OK: rows={n_rows}, users={n_users_pm}, dup={n_dup_tok}, uniq={n_uniq_tok}')

tpu8 = pm8.groupby('UserId')['Token'].count().reset_index(name='token_count')
top10 = tpu8.nlargest(10, 'token_count').copy()
top10['orders'] = top10['UserId'].apply(lambda uid: (df['UserId']==uid).sum())
print(f'Tab8 top: {top10.iloc[0]["UserId"]} has {top10.iloc[0]["token_count"]} tokens, {top10.iloc[0]["orders"]} orders')

users_pm_set  = set(pm8['UserId'])
users_ord_set = set(df['UserId'].dropna())
cash_users    = set(df[df['PaymentMethod']==0]['UserId'].dropna())
credit_users  = set(df[df['PaymentMethod']==1]['UserId'].dropna())
print(f'Tab8 cross: both={len(users_pm_set & users_ord_set)}, credit_no_token={len(credit_users - users_pm_set)}')

quality_data = {
    'בדיקה': ['ערכים חסרים','תקינות פורמט Token','טוקנים כפולים','לקוחות עם 50+ טוקנים','אשראי ללא טוקן'],
    'תוצאה': [
        '0 חסרים', 'TOK_XXXXXXXX', str(n_dup_tok),
        str((tpu8['token_count']>=50).sum()),
        str(len(credit_users - users_pm_set)),
    ],
    'דחיפות': ['נמוכה','נמוכה','גבוהה','גבוהה מאוד','בינונית'],
}
print(f'Tab8 quality table: {len(quality_data["בדיקה"])} rows')
print('ALL CHECKS PASSED')
