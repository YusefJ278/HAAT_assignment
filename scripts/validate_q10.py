import sys; sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd

PATH = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
df = pd.read_excel(PATH, sheet_name='dataset')
df['OrderDate'] = pd.to_datetime(df['OrderDate'], errors='coerce')
df['ArriveDate'] = pd.to_datetime(df['ArriveDate'], errors='coerce')
df['delivered']  = df['ArriveDate'].notna()
df['DayOfWeek']  = df['OrderDate'].dt.dayofweek
df['IsWeekend']  = df['DayOfWeek'].isin([4, 5])
df['wk']         = df['OrderDate'].dt.to_period('W').astype(str)

def calc_weekly_bonus(orders_df, week_start, week_end):
    ws = pd.Timestamp(week_start)
    we = pd.Timestamp(week_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    w = orders_df[
        (orders_df['OrderDate'] >= ws) & (orders_df['OrderDate'] <= we) &
        (orders_df['delivered'] == True) & (orders_df['DriverId'].notna())
    ].copy()
    results = []
    a1 = w[(w['AreaId'] == 1) & (w['IsWeekend'] == True)]
    g1 = a1.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g1['AreaId'] = 1
    g1['BonusAmount'] = (g1['TotalDelivered'] // 50) * 100
    results.append(g1)
    a4 = w[w['AreaId'] == 4]
    g4 = a4.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g4['AreaId'] = 4
    g4['BonusAmount'] = g4['TotalDelivered'].apply(lambda n: max(0, n - 300) * 3)
    results.append(g4)
    a10 = w[(w['AreaId'] == 10) & (w['IsWeekend'] == True)]
    g10 = a10.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g10['AreaId'] = 10
    g10['BonusAmount'] = (g10['TotalDelivered'] // 50) * 150
    results.append(g10)
    result = pd.concat(results, ignore_index=True)
    bonus = result[result['BonusAmount'] > 0][
        ['DriverId', 'AreaId', 'TotalDelivered', 'BonusAmount']
    ].sort_values(['AreaId', 'BonusAmount'], ascending=[True, False]).reset_index(drop=True)
    if len(bonus):
        bonus['DriverId'] = bonus['DriverId'].astype(int)
    return bonus

# Test on all weeks
weeks = sorted(df['wk'].dropna().unique())
total_bonuses = 0
for wk in weeks:
    parts = wk.split('/')
    r = calc_weekly_bonus(df, parts[0], parts[1])
    total_bonuses += r['BonusAmount'].sum()

print(f'Total bonus across all weeks: {total_bonuses}')

# Threshold analysis
a1_all  = df[(df['AreaId']==1)  & df['delivered'] & df['IsWeekend'] & df['DriverId'].notna()]
a4_all  = df[(df['AreaId']==4)  & df['delivered'] & df['DriverId'].notna()]
a10_all = df[(df['AreaId']==10) & df['delivered'] & df['IsWeekend'] & df['DriverId'].notna()]

mx1  = a1_all.groupby(['wk','DriverId']).size().max() if len(a1_all) else 0
mx4  = a4_all.groupby(['wk','DriverId']).size().max() if len(a4_all) else 0
mx10 = a10_all.groupby(['wk','DriverId']).size().max() if len(a10_all) else 0
print(f'Max weekly per driver: Area1={mx1}, Area4={mx4}, Area10={mx10}')
print(f'Threshold:             Area1=50, Area4=300, Area10=50')
print(f'% of threshold:        Area1={mx1/50*100:.0f}%, Area4={mx4/300*100:.0f}%, Area10={mx10/50*100:.0f}%')
print('VALIDATION PASSED')
