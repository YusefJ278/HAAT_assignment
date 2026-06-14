import sys; sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd

PATH = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
df = pd.read_excel(PATH, sheet_name='dataset')
df['OrderDate'] = pd.to_datetime(df['OrderDate'], errors='coerce')
df['ArriveDate'] = pd.to_datetime(df['ArriveDate'], errors='coerce')
df['delivered']  = df['ArriveDate'].notna()
df['DayOfWeek']  = df['OrderDate'].dt.dayofweek
df['IsWeekend']  = df['DayOfWeek'].isin([4, 5])

def calc_bonus(orders_df, week_start, week_end):
    ws = pd.Timestamp(week_start)
    we = pd.Timestamp(week_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    w = orders_df[
        (orders_df['OrderDate'] >= ws) & (orders_df['OrderDate'] <= we) &
        orders_df['delivered'] & orders_df['DriverId'].notna()
    ].copy()
    results = []
    a1  = w[(w['AreaId']==1)  & w['IsWeekend']]
    g1  = a1.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g1['AreaId'] = 1;  g1['BonusAmount'] = (g1['TotalDelivered'] // 50) * 100
    results.append(g1)
    a4  = w[w['AreaId']==4]
    g4  = a4.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g4['AreaId'] = 4;  g4['BonusAmount'] = g4['TotalDelivered'].apply(lambda n: max(0, n-300)*3)
    results.append(g4)
    a10 = w[(w['AreaId']==10) & w['IsWeekend']]
    g10 = a10.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g10['AreaId'] = 10; g10['BonusAmount'] = (g10['TotalDelivered'] // 50) * 150
    results.append(g10)
    result = pd.concat(results, ignore_index=True)
    bonus = result[result['BonusAmount'] > 0][
        ['DriverId','AreaId','TotalDelivered','BonusAmount']
    ].sort_values(['AreaId','BonusAmount'], ascending=[True,False]).reset_index(drop=True)
    if len(bonus): bonus['DriverId'] = bonus['DriverId'].astype(int)
    return bonus

# ── בדיקה: על כל התקופה כולה ─────────────────────────────────────
print("=== כל התקופה (1 פבר' – 31 מאי 2024) ===")
r_all = calc_bonus(df, '2024-02-01', '2024-05-31')
print(r_all.to_string(index=False) if len(r_all) else 'ריק')
print()

# ── בדיקה: חלונות של 2, 3, 4 שבועות ────────────────────────────────
df['Week'] = df['OrderDate'].dt.to_period('W').astype(str)
weeks = sorted(df['Week'].dropna().unique())

print("=== חלונות הזמן המינימליים למציאת בונוס ===")
for window_size in [2, 3, 4, 6, 8, len(weeks)]:
    found = []
    for i in range(len(weeks) - window_size + 1):
        ws_str = weeks[i].split('/')[0]
        we_str = weeks[i + window_size - 1].split('/')[1]
        r = calc_bonus(df, ws_str, we_str)
        if len(r):
            found.append({
                'חלון': f'{ws_str} → {we_str}',
                'שבועות': window_size,
                'נהגים': len(r),
                'בונוס_כולל': r['BonusAmount'].sum(),
            })
    label = 'כל התקופה' if window_size == len(weeks) else f'{window_size} שבועות'
    if found:
        best = min(found, key=lambda x: x['נהגים'])  # first found
        print(f"  [{label}] ✅ נמצאו בונוסים! | חלון: {found[0]['חלון']} | נהגים: {found[0]['נהגים']} | בונוס: ₪{found[0]['בונוס_כולל']}")
    else:
        print(f"  [{label}] ✗ אין בונוסים")

# ── נהג 538 AreaId10 — מעקב מצטבר ─────────────────────────────────
print("\n=== מעקב מצטבר: נהג 538 באזור 10 (הקרוב ביותר לסף) ===")
d538 = df[(df['DriverId']==538) & (df['AreaId']==10) & df['delivered'] & df['IsWeekend']]
d538_sorted = d538.sort_values('OrderDate')
d538_sorted['cumulative'] = range(1, len(d538_sorted)+1)
print(f"סה\"כ מסירות סוף שבוע כל התקופה: {len(d538_sorted)}")
cross50 = d538_sorted[d538_sorted['cumulative'] >= 50]
if len(cross50):
    first_cross = cross50.iloc[0]
    print(f"חצה 50 מסירות ב: {first_cross['OrderDate'].date()}")
    print(f"תקופה נדרשת: {d538_sorted.iloc[0]['OrderDate'].date()} → {first_cross['OrderDate'].date()}")
    print(d538_sorted[['OrderDate','cumulative']].tail(10).to_string(index=False))
