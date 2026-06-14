import sys; sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import numpy as np

PATH = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
df   = pd.read_excel(PATH, sheet_name='dataset')
pm   = pd.read_excel(PATH, sheet_name='payment_methods_user')
df['delivered'] = pd.to_datetime(df['ArriveDate'], errors='coerce').notna()

print('=== טבלת payment_methods_user ===')
n_uid = pm['UserId'].nunique()
print(f'עמודות: {pm.columns.tolist()}')
print(f'שורות: {len(pm)} | UserId ייחודיים: {n_uid}')

tpu = pm.groupby('UserId')['Token'].count().sort_values(ascending=False)
print('\nפיזור טוקנים לכל לקוח:')
print(tpu.value_counts().sort_index().head(15).to_string())
n10 = (tpu >= 10).sum()
n5  = (tpu >= 5).sum()
n1  = (tpu == 1).sum()
print(f'\nלקוחות עם 10+ טוקנים: {n10}')
print(f'לקוחות עם 5+ טוקנים:  {n5}')
print(f'לקוחות עם 1 טוקן:     {n1}')

print('\nTOP 10 לקוחות לפי מספר טוקנים:')
for uid, cnt in tpu.head(10).items():
    orders_count = len(df[df['UserId'] == uid])
    cash_orders  = (df[df['UserId'] == uid]['PaymentMethod'] == 0).sum()
    print(f'  UserId {uid}: {cnt} טוקנים | {orders_count} הזמנות | {cash_orders} מזומן')

print('\n=== השוואה: לקוחות בהזמנות vs payment_methods_user ===')
users_pm  = set(pm['UserId'])
users_ord = set(df['UserId'].dropna())
both = users_pm & users_ord
only_pm  = users_pm - users_ord
only_ord = users_ord - users_pm
print(f'בהזמנות: {len(users_ord):,} | ב-payment: {len(users_pm):,}')
print(f'בשתיהן: {len(both):,}')
only_pm_n = len(only_pm)
only_ord_n = len(only_ord)
print(f'רק ב-payment (לא הזמינו מעולם): {only_pm_n:,}')
print(f'רק בהזמנות (אין טוקן כלל):      {only_ord_n:,}')

print('\n=== לקוחות מזומן עם טוקן (חריג!) ===')
cash_users   = set(df[df['PaymentMethod'] == 0]['UserId'].dropna())
credit_users = set(df[df['PaymentMethod'] == 1]['UserId'].dropna())
cash_with_token   = cash_users & users_pm
credit_no_token   = credit_users - users_pm
cash_pct   = len(cash_with_token) / len(cash_users) * 100
credit_pct = len(credit_no_token) / len(credit_users) * 100
print(f'משלמי מזומן עם טוקן:    {len(cash_with_token):,} / {len(cash_users):,} ({cash_pct:.1f}%)')
print(f'משלמי אשראי ללא טוקן:   {len(credit_no_token):,} / {len(credit_users):,} ({credit_pct:.1f}%)')

# בדיקה: לקוח שמשלם גם מזומן וגם אשראי
mixed_users = cash_users & credit_users
print(f'\nלקוחות שמשלמים גם מזומן וגם אשראי: {len(mixed_users):,}')
if len(mixed_users) > 0:
    sample = list(mixed_users)[:5]
    for uid in sample:
        u_pm = df[df['UserId'] == uid]['PaymentMethod'].value_counts().to_dict()
        n_tok = tpu.get(uid, 0)
        print(f'  UserId {uid}: {u_pm} | {n_tok} טוקנים')

print('\n=== טוקנים כפולים ===')
dup_count = pm['Token'].duplicated().sum()
n_unique_tok = pm['Token'].nunique()
print(f'טוקנים כפולים: {dup_count}')
print(f'טוקנים ייחודיים: {n_unique_tok} מתוך {len(pm)} שורות')
if dup_count == 0:
    print('=> כל טוקן מופיע בדיוק פעם אחת (ייחודי לחלוטין) — תקין')

print('\n=== תקינות פורמט Token ===')
starts_tok = pm['Token'].str.startswith('TOK_').all()
avg_len = pm['Token'].str.len().mean()
print(f'כל הטוקנים מתחילים ב-TOK_: {starts_tok}')
print(f'אורך ממוצע: {avg_len:.1f} תווים')
print(f'ערכים חסרים: {pm.isnull().sum().sum()}')
