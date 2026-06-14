import sys; sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import numpy as np
import math

PATH = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
df   = pd.read_excel(PATH, sheet_name='dataset')

df['OrderDate'] = pd.to_datetime(df['OrderDate'], errors='coerce')
df['ArriveDate'] = pd.to_datetime(df['ArriveDate'], errors='coerce')
df['delivered']  = df['ArriveDate'].notna()
df['DayOfWeek']  = df['OrderDate'].dt.dayofweek   # 0=Mon, 4=Fri, 5=Sat
df['IsWeekend']  = df['DayOfWeek'].isin([4, 5])   # Friday, Saturday


# ══════════════════════════════════════════════════════════════════
# פונקציית חישוב בונוס שבועי
# ══════════════════════════════════════════════════════════════════
def calc_weekly_bonus(orders_df, week_start, week_end):
    """
    מחשבת בונוסים שבועיים לשליחים לפי חוקי אזור.

    Parameters
    ----------
    orders_df  : DataFrame עם עמודות: OrderDate, ArriveDate, DriverId, AreaId
    week_start : str/date — תחילת השבוע (כולל)
    week_end   : str/date — סוף השבוע (כולל)

    Returns
    -------
    DataFrame: DriverId, AreaId, TotalDelivered, BonusAmount (רק שורות עם בונוס > 0)

    חוקי הבונוס:
    - AreaId  1 : סוף שבוע בלבד — כל 50 מסירות → 100 ₪
    - AreaId  4 : כל הימים — כל מסירה מעל 300 → 3 ₪/מסירה
    - AreaId 10 : סוף שבוע בלבד — כל 50 מסירות → 150 ₪
    """
    ws = pd.Timestamp(week_start)
    we = pd.Timestamp(week_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

    # סינון לשבוע ולמסירות בלבד
    w = orders_df[
        (orders_df['OrderDate'] >= ws) &
        (orders_df['OrderDate'] <= we) &
        (orders_df['delivered'] == True) &
        (orders_df['DriverId'].notna())
    ].copy()

    results = []

    # ── AreaId 1 : סוף שבוע, 100 ₪ לכל 50 מסירות ─────────────────
    a1 = w[(w['AreaId'] == 1) & (w['IsWeekend'] == True)]
    grp1 = a1.groupby('DriverId').size().reset_index(name='TotalDelivered')
    grp1['AreaId']      = 1
    grp1['BonusAmount'] = (grp1['TotalDelivered'] // 50) * 100
    results.append(grp1)

    # ── AreaId 4 : כל הימים, 3 ₪ לכל מסירה מעל 300 ────────────────
    a4 = w[w['AreaId'] == 4]
    grp4 = a4.groupby('DriverId').size().reset_index(name='TotalDelivered')
    grp4['AreaId']      = 4
    grp4['BonusAmount'] = grp4['TotalDelivered'].apply(lambda n: max(0, n - 300) * 3)
    results.append(grp4)

    # ── AreaId 10 : סוף שבוע, 150 ₪ לכל 50 מסירות ────────────────
    a10 = w[(w['AreaId'] == 10) & (w['IsWeekend'] == True)]
    grp10 = a10.groupby('DriverId').size().reset_index(name='TotalDelivered')
    grp10['AreaId']      = 10
    grp10['BonusAmount'] = (grp10['TotalDelivered'] // 50) * 150

    results.append(grp10)

    result = pd.concat(results, ignore_index=True)
    result = result[result['BonusAmount'] > 0][
        ['DriverId', 'AreaId', 'TotalDelivered', 'BonusAmount']
    ].sort_values(['AreaId', 'BonusAmount'], ascending=[True, False]).reset_index(drop=True)
    result['DriverId'] = result['DriverId'].astype(int)
    return result


# ══════════════════════════════════════════════════════════════════
# בדיקה על שבוע אחד מהדאטאסט
# ══════════════════════════════════════════════════════════════════
WEEK_START = '2024-04-15'
WEEK_END   = '2024-04-21'

print(f"=== בונוס שבועי: {WEEK_START} עד {WEEK_END} ===")
result = calc_weekly_bonus(df, WEEK_START, WEEK_END)
print(result.to_string(index=False))
print(f"\nסה\"כ שורות (נהגים עם בונוס): {len(result)}")
print(f"סה\"כ בונוסים: ₪{result['BonusAmount'].sum():,.0f}")
print(f"ממוצע בונוס לנהג: ₪{result['BonusAmount'].mean():.1f}")

print("\n--- לפי אזור ---")
by_area = result.groupby('AreaId').agg(
    נהגים=('DriverId','count'),
    סה_כ_מסירות=('TotalDelivered','sum'),
    סה_כ_בונוס=('BonusAmount','sum'),
    ממוצע_בונוס=('BonusAmount','mean'),
)
print(by_area.to_string())

# הצגת TOP 5 לפי אזור
for area_id in [1, 4, 10]:
    top = result[result['AreaId']==area_id].head(5)
    if len(top):
        print(f"\n--- TOP 5 נהגים ב-AreaId {area_id} ---")
        print(top.to_string(index=False))

# ── בדיקה על שבוע נוסף ────────────────────────────────────────
print("\n\n=== בדיקה שבוע ראשון בדאטא: 2024-02-01 עד 2024-02-07 ===")
result2 = calc_weekly_bonus(df, '2024-02-01', '2024-02-07')
print(result2.to_string(index=False))
print(f"סה\"כ בונוסים: ₪{result2['BonusAmount'].sum():,.0f}")

# ── סיכום לכל השבועות ─────────────────────────────────────────
print("\n\n=== סיכום בונוסים לכל שבועות הדאטאסט ===")
df['Week'] = df['OrderDate'].dt.to_period('W')
all_weeks  = sorted(df['Week'].dropna().unique())

weekly_summary = []
for wk in all_weeks:
    wk_str = str(wk)
    parts  = wk_str.split('/')
    ws, we = parts[0], parts[1]
    r = calc_weekly_bonus(df, ws, we)
    total_bonus = r['BonusAmount'].sum()
    n_drivers   = len(r)
    weekly_summary.append({'שבוע': wk_str, 'נהגים_עם_בונוס': n_drivers, 'סה_כ_בונוס_ILS': total_bonus})

wsum = pd.DataFrame(weekly_summary)
print(wsum.to_string(index=False))
print(f"\nסה\"כ בונוסים כל התקופה: ₪{wsum['סה_כ_בונוס_ILS'].sum():,.0f}")
print(f"ממוצע שבועי:            ₪{wsum['סה_כ_בונוס_ILS'].mean():,.0f}")
print(f"שבוע שיא:              {wsum.loc[wsum['סה_כ_בונוס_ILS'].idxmax(), 'שבוע']} (₪{wsum['סה_כ_בונוס_ILS'].max():,.0f})")
