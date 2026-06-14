import sys; sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd

PATH = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
df = pd.read_excel(PATH, sheet_name='dataset')
df['OrderDate'] = pd.to_datetime(df['OrderDate'], errors='coerce')
df['ArriveDate'] = pd.to_datetime(df['ArriveDate'], errors='coerce')
df['delivered']  = df['ArriveDate'].notna()
df['DayOfWeek']  = df['OrderDate'].dt.dayofweek
df['IsWeekend']  = df['DayOfWeek'].isin([4, 5])

print('DriverId nulls:', df['DriverId'].isna().sum(), '/', len(df))
print('DriverId sample:', df['DriverId'].dropna().head(10).tolist())
print('delivered count:', df['delivered'].sum())

a1 = (df['AreaId'] == 1) & df['delivered']
print('AreaId 1 + delivered:', a1.sum())
a1w = a1 & df['IsWeekend']
print('AreaId 1 + delivered + weekend:', a1w.sum())
a1wd = a1w & df['DriverId'].notna()
print('AreaId 1 + delivered + weekend + driverid:', a1wd.sum())

# Per-driver in AreaId 1
grp1 = df[a1wd].groupby('DriverId').size().sort_values(ascending=False)
print('\nTop 10 drivers AreaId1 weekend deliveries:')
print(grp1.head(10))
print('\nMax deliveries by single driver AreaId1 weekend:', grp1.max() if len(grp1) > 0 else 0)

a4 = (df['AreaId'] == 4) & df['delivered'] & df['DriverId'].notna()
grp4 = df[a4].groupby('DriverId').size().sort_values(ascending=False)
print('\nTop 10 drivers AreaId4 all-day deliveries:')
print(grp4.head(10))
print('Max deliveries AreaId4:', grp4.max() if len(grp4) > 0 else 0)

a10 = (df['AreaId'] == 10) & df['delivered'] & df['IsWeekend'] & df['DriverId'].notna()
grp10 = df[a10].groupby('DriverId').size().sort_values(ascending=False)
print('\nTop 10 drivers AreaId10 weekend deliveries:')
print(grp10.head(10))
print('Max deliveries AreaId10:', grp10.max() if len(grp10) > 0 else 0)

# Check all-time totals (ignoring weekly constraint)
print('\n=== ללא הגבלת שבוע (כל התקופה) ===')
print(f'AreaId 1 — מקסימום מסירות סוף שבוע ע"י נהג אחד כל התקופה: {grp1.max() if len(grp1)>0 else 0}')
print(f'AreaId 4 — מקסימום מסירות ע"י נהג אחד כל התקופה: {grp4.max() if len(grp4)>0 else 0}')
print(f'AreaId 10 — מקסימום מסירות סוף שבוע ע"י נהג אחד כל התקופה: {grp10.max() if len(grp10)>0 else 0}')

# Check per-week
print('\n=== מסירות שבועיות מקסימליות לנהג ===')
df['Week'] = df['OrderDate'].dt.to_period('W').astype(str)

# AreaId 1 weekend per week per driver
a1_all = df[(df['AreaId'] == 1) & df['delivered'] & df['IsWeekend'] & df['DriverId'].notna()]
if len(a1_all):
    wpd1 = a1_all.groupby(['Week','DriverId']).size().reset_index(name='deliveries')
    print(f'AreaId1 max per driver per week: {wpd1["deliveries"].max()}')
    print(wpd1.nlargest(5, 'deliveries'))

a4_all = df[(df['AreaId'] == 4) & df['delivered'] & df['DriverId'].notna()]
if len(a4_all):
    wpd4 = a4_all.groupby(['Week','DriverId']).size().reset_index(name='deliveries')
    print(f'AreaId4 max per driver per week: {wpd4["deliveries"].max()}')
    print(wpd4.nlargest(5, 'deliveries'))
