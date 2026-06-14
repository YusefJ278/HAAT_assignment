import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import arabic_reshaper
from bidi.algorithm import get_display

def heb(text):
    """ממיר טקסט עברי לתצוגה נכונה (RTL) ב-matplotlib."""
    return get_display(arabic_reshaper.reshape(str(text)))

# ── תיקיית פלט ──────────────────────────────────────────────────
OUT = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\charts'
os.makedirs(OUT, exist_ok=True)

# ── צבעי פלטה ────────────────────────────────────────────────────
C_BLUE   = '#1A376C'
C_ACCENT = '#2674C1'
C_GREEN  = '#1B5E20'
C_RED    = '#8B0000'
C_GOLD   = '#8B6000'
C_MGREY  = '#595959'

plt.rcParams.update({
    'font.family':       'Arial',
    'axes.spines.top':   False,
    'axes.spines.right': False,
    'axes.grid':         True,
    'grid.alpha':        0.3,
    'grid.color':        '#CCCCCC',
    'figure.facecolor':  'white',
    'axes.facecolor':    'white',
})

# ── טעינת נתונים ────────────────────────────────────────────────
PATH = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
df    = pd.read_excel(PATH, sheet_name='dataset')
areas = pd.read_excel(PATH, sheet_name='areas')
df    = df.merge(areas, on='AreaId', how='left')

for c in ['OrderDate','ArriveDate','DriverCandidateAssignedDate','PickedDate','StartPrepare','RejectedDateByRestaurant']:
    df[c] = pd.to_datetime(df[c], errors='coerce')

df['Month']     = df['OrderDate'].dt.to_period('M')
df['Hour']      = df['OrderDate'].dt.hour
df['delivered'] = df['ArriveDate'].notna()

# ════════════════════════════════════════════════════════════════
# Q3-A  |  גרף 1: שיעור מסירה לפי חודש
# ════════════════════════════════════════════════════════════════
monthly = df.groupby('Month').agg(
    orders=('Id','count'), delivered=('delivered','sum')
).assign(rate=lambda x: x['delivered']/x['orders']*100)

fig, ax = plt.subplots(figsize=(7, 3.5))
months_lbl = [heb('פבר׳ 2024'), heb('מרץ 2024'), heb('אפר׳ 2024'), heb('מאי 2024')]
rates      = monthly['rate'].values
colors     = [C_ACCENT if r < 67 else C_GREEN for r in rates]
bars = ax.bar(months_lbl, rates, color=colors, width=0.55, zorder=3)

avg = df['delivered'].mean() * 100
ax.axhline(avg, color=C_RED, linestyle='--', linewidth=1.4, label=heb(f'ממוצע: {avg:.1f}%'))

for bar, r in zip(bars, rates):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            f'{r:.1f}%', ha='center', va='bottom', fontsize=11,
            fontweight='bold', color=C_BLUE)

ax.set_ylim(55, 80)
ax.set_ylabel(heb('שיעור מסירה (%)'), fontsize=11, color=C_MGREY)
ax.set_title(heb('שיעור מסירה לפי חודש'), fontsize=13, fontweight='bold', color=C_BLUE, pad=12)
ax.legend(fontsize=10)
ax.tick_params(axis='x', labelsize=11)
ax.tick_params(axis='y', labelsize=10)
plt.tight_layout()
fig.savefig(f'{OUT}/q3a_monthly.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q3a_monthly.png')

# ════════════════════════════════════════════════════════════════
# Q3-A  |  גרף 2: שיעור מסירה לפי אזור
# ════════════════════════════════════════════════════════════════
area_stats = df.groupby('AreaName').agg(
    orders=('Id','count'), delivered=('delivered','sum')
).assign(rate=lambda x: x['delivered']/x['orders']*100).sort_values('rate')

fig, ax = plt.subplots(figsize=(8, 5))
avg = df['delivered'].mean() * 100

def area_color(r):
    if r >= 90: return C_GREEN
    if r >= 70: return C_ACCENT
    return C_RED

colors = [area_color(r) for r in area_stats['rate']]
bars = ax.barh(area_stats.index, area_stats['rate'], color=colors, height=0.65, zorder=3)
ax.axvline(avg, color=C_GOLD, linestyle='--', linewidth=1.5, label=heb(f'ממוצע: {avg:.1f}%'))

for bar, r in zip(bars, area_stats['rate']):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
            f'{r:.1f}%', va='center', fontsize=10, fontweight='bold', color=C_BLUE)

patches = [
    mpatches.Patch(color=C_GREEN,  label=heb('ביצוע גבוה (>90%)')),
    mpatches.Patch(color=C_ACCENT, label=heb('ביצוע בינוני (70-90%)')),
    mpatches.Patch(color=C_RED,    label=heb('ביצוע נמוך (<70%)')),
]
ax.legend(handles=patches, fontsize=9, loc='lower right')
ax.set_xlim(0, 110)
ax.set_xlabel(heb('שיעור מסירה (%)'), fontsize=11, color=C_MGREY)
ax.set_title(heb('שיעור מסירה לפי אזור'), fontsize=13, fontweight='bold', color=C_BLUE, pad=12)
ax.tick_params(axis='y', labelsize=10)
plt.tight_layout()
fig.savefig(f'{OUT}/q3a_area.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q3a_area.png')

# ════════════════════════════════════════════════════════════════
# Q3-B  |  גרף 3: נפח + שיעור מסירה לפי שעה
# ════════════════════════════════════════════════════════════════
hourly = df.groupby('Hour').agg(
    orders=('Id','count'), delivered=('delivered','sum')
).assign(rate=lambda x: x['delivered']/x['orders']*100)

fig, ax1 = plt.subplots(figsize=(10, 4.5))
ax2 = ax1.twinx()

hours      = hourly.index.tolist()
bar_colors = [C_RED if r < 62 else (C_ACCENT if r < 70 else C_GREEN) for r in hourly['rate']]
bars = ax1.bar(hours, hourly['orders'], color=bar_colors, alpha=0.6, width=0.7, zorder=2,
               label=heb('נפח הזמנות'))
ax2.plot(hours, hourly['rate'], color=C_BLUE, linewidth=2.2, marker='o', markersize=5,
         zorder=3, label=heb('שיעור מסירה %'))

peak_h  = hourly['orders'].idxmax()
worst_h = hourly['rate'].idxmin()
ax1.annotate(heb(f'שיא נפח\n{int(hourly.loc[peak_h,"orders"])} הזמנות'),
             xy=(peak_h, hourly.loc[peak_h,'orders']),
             xytext=(peak_h-3, hourly.loc[peak_h,'orders']+80),
             fontsize=8.5, color=C_MGREY,
             arrowprops=dict(arrowstyle='->', color=C_MGREY, lw=1.2))
ax2.annotate(heb(f'שפל מסירה\n{hourly.loc[worst_h,"rate"]:.1f}%'),
             xy=(worst_h, hourly.loc[worst_h,'rate']),
             xytext=(worst_h+1.5, hourly.loc[worst_h,'rate']-4),
             fontsize=8.5, color=C_RED,
             arrowprops=dict(arrowstyle='->', color=C_RED, lw=1.2))

ax1.set_xlabel(heb('שעה ביום'), fontsize=11, color=C_MGREY)
ax1.set_ylabel(heb('נפח הזמנות'), fontsize=11, color=C_MGREY)
ax2.set_ylabel(heb('שיעור מסירה (%)'), fontsize=11, color=C_BLUE)
ax1.set_xticks(hours)
ax1.set_xticklabels([f'{h:02d}:00' for h in hours], rotation=45, ha='right', fontsize=9)
ax2.set_ylim(50, 105)
ax1.set_title(heb('נפח הזמנות ושיעור מסירה לפי שעה'), fontsize=13, fontweight='bold', color=C_BLUE, pad=12)

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1+lines2, labels1+labels2, fontsize=9, loc='upper left')
plt.tight_layout()
fig.savefig(f'{OUT}/q3b_hourly.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q3b_hourly.png')

# ════════════════════════════════════════════════════════════════
# Q4-A  |  גרף 4: TOP 10 + BOTTOM 10 שליחים
# ════════════════════════════════════════════════════════════════
driver_df = df[df['DriverId'].notna()].copy()
driver_stats = driver_df.groupby('DriverId').agg(
    total_assigned=('Id','count'),
    total_delivered=('delivered','sum'),
    first_order=('OrderDate','min'),
    last_order=('OrderDate','max'),
).reset_index()
driver_stats['active_days']   = (driver_stats['last_order'] - driver_stats['first_order']).dt.days + 1
driver_stats['delivery_rate'] = driver_stats['total_delivered'] / driver_stats['total_assigned'] * 100
driver_stats_sorted           = driver_stats.sort_values('total_delivered', ascending=False)

top10    = driver_stats_sorted.head(10).copy()
bottom10 = driver_stats_sorted.tail(10).copy()

fig, (ax_top, ax_bot) = plt.subplots(1, 2, figsize=(12, 5))

# TOP 10
top10_lbl  = [heb(f'שליח {int(d)}') for d in top10['DriverId']]
top10_vals = top10['total_delivered'].values
c_top      = [C_RED if v > 200 else C_ACCENT for v in top10_vals]
ax_top.barh(top10_lbl[::-1], top10_vals[::-1], color=c_top[::-1], height=0.65, zorder=3)
for i, (v, d) in enumerate(zip(top10_vals[::-1], top10['active_days'].values[::-1])):
    ax_top.text(v + 1, i, heb(f'{int(v)} ({int(d)} ימים)'), va='center', fontsize=9, color=C_MGREY)
ax_top.set_title(heb('TOP 10 — הכי הרבה משלוחים'), fontsize=11, fontweight='bold', color=C_BLUE)
ax_top.set_xlabel(heb('הזמנות שנמסרו'), fontsize=10)
ax_top.set_xlim(0, 270)
ax_top.tick_params(axis='y', labelsize=10)

# BOTTOM 10
bot10_lbl  = [heb(f'שליח {int(d)}') for d in bottom10['DriverId']]
bot10_vals = bottom10['total_delivered'].values
ax_bot.barh(bot10_lbl[::-1], bot10_vals[::-1], color=C_MGREY, height=0.65, alpha=0.7, zorder=3)
for i, v in enumerate(bot10_vals[::-1]):
    ax_bot.text(v + 0.02, i, str(int(v)), va='center', fontsize=9, color=C_MGREY)
ax_bot.set_title(heb('BOTTOM 10 — הכי פחות משלוחים'), fontsize=11, fontweight='bold', color=C_MGREY)
ax_bot.set_xlabel(heb('הזמנות שנמסרו'), fontsize=10)
ax_bot.set_xlim(0, 5)
ax_bot.tick_params(axis='y', labelsize=10)

fig.suptitle(heb('דירוג שליחים לפי הזמנות שנמסרו'), fontsize=13, fontweight='bold', color=C_BLUE, y=1.01)
plt.tight_layout()
fig.savefig(f'{OUT}/q4a_ranking.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q4a_ranking.png')

# ════════════════════════════════════════════════════════════════
# Q4-B  |  גרף 5: TOP 10 שליחים לפי יעילות (ללא חריגים קיצוניים)
# ════════════════════════════════════════════════════════════════
driver_stats['efficiency'] = driver_stats['total_delivered'] / driver_stats['active_days']
legit   = driver_stats[(driver_stats['active_days'] >= 7) & (driver_stats['total_delivered'] >= 10)]
top_eff = legit.sort_values('efficiency', ascending=False).head(10)

fig, ax = plt.subplots(figsize=(8, 4.5))
eff_lbl  = [heb(f'שליח {int(d)}') for d in top_eff['DriverId']]
eff_vals = top_eff['efficiency'].values
ax.barh(eff_lbl[::-1], eff_vals[::-1], color=C_GREEN, height=0.65, zorder=3)
for i, (v, tot, days) in enumerate(zip(eff_vals[::-1],
                                        top_eff['total_delivered'].values[::-1],
                                        top_eff['active_days'].values[::-1])):
    ax.text(v + 0.01, i, heb(f'{v:.2f} ({int(tot)} משלוחים / {int(days)} ימים)'),
            va='center', fontsize=9, color=C_MGREY)

avg_eff = legit['efficiency'].mean()
ax.axvline(avg_eff, color=C_GOLD, linestyle='--', linewidth=1.4, label=heb(f'ממוצע: {avg_eff:.2f}'))
ax.set_xlabel(heb('משלוחים מוצלחים ליום פעיל'), fontsize=11, color=C_MGREY)
ax.set_title(heb('TOP 10 שליחים — מדד יעילות מנורמל ליום') + '\n' +
             heb('(שליחים עם 7+ ימים ו-10+ משלוחים)'),
             fontsize=11, fontweight='bold', color=C_BLUE, pad=10)
ax.set_xlim(0, eff_vals.max() * 1.5)
ax.legend(fontsize=9)
ax.tick_params(axis='y', labelsize=10)
plt.tight_layout()
fig.savefig(f'{OUT}/q4b_efficiency.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q4b_efficiency.png')

# ════════════════════════════════════════════════════════════════
# Q4-C  |  גרף 6: התפלגות מדד היעילות + חריגים
# ════════════════════════════════════════════════════════════════
eff_all = driver_stats['efficiency']
Q1_v    = eff_all.quantile(0.25)
Q3_v    = eff_all.quantile(0.75)
IQR_v   = Q3_v - Q1_v
fence   = Q3_v + 3 * IQR_v
n_out   = (eff_all > fence).sum()

fig, (ax_main, ax_zoom) = plt.subplots(1, 2, figsize=(12, 4.5))

ax_main.hist(eff_all[eff_all <= fence], bins=40, color=C_ACCENT, alpha=0.8,
             edgecolor='white', linewidth=0.5, zorder=3, label=heb('שליחים רגילים'))
ax_main.hist(eff_all[eff_all > fence],  bins=5,  color=C_RED,   alpha=0.9,
             edgecolor='white', linewidth=0.5, zorder=4, label=heb(f'חריגים (n={n_out})'))
ax_main.axvline(fence, color=C_RED, linestyle='--', linewidth=1.5,
                label=heb(f'גבול חריגים: {fence:.2f}'))
ax_main.axvline(eff_all.median(), color=C_GREEN, linestyle='-', linewidth=1.5,
                label=heb(f'חציון: {eff_all.median():.3f}'))
ax_main.set_xlabel(heb('משלוחים ליום פעיל'), fontsize=11, color=C_MGREY)
ax_main.set_ylabel(heb('מספר שליחים'), fontsize=11, color=C_MGREY)
ax_main.set_title(heb('התפלגות מדד יעילות — כל השליחים'), fontsize=11, fontweight='bold', color=C_BLUE)
ax_main.legend(fontsize=9)

zoom_data    = eff_all[eff_all <= 2.0]
pct_in_range = zoom_data.count() / len(eff_all) * 100
ax_zoom.hist(zoom_data, bins=30, color=C_ACCENT, alpha=0.8, edgecolor='white', linewidth=0.5, zorder=3)
ax_zoom.axvline(eff_all.median(), color=C_GREEN, linestyle='-', linewidth=1.5,
                label=heb(f'חציון: {eff_all.median():.3f}'))
ax_zoom.axvline(eff_all.mean(), color=C_GOLD, linestyle='--', linewidth=1.5,
                label=heb(f'ממוצע: {eff_all.mean():.3f}'))
ax_zoom.set_xlabel(heb('משלוחים ליום פעיל'), fontsize=11, color=C_MGREY)
ax_zoom.set_ylabel(heb('מספר שליחים'), fontsize=11, color=C_MGREY)
ax_zoom.set_title(heb(f'זום: 0-2 משלוחים/יום ({pct_in_range:.1f}% מהשליחים)'),
                  fontsize=11, fontweight='bold', color=C_BLUE)
ax_zoom.legend(fontsize=9)

plt.suptitle(heb('התפלגות מדד יעילות השליחים'), fontsize=13, fontweight='bold', color=C_BLUE, y=1.02)
plt.tight_layout()
fig.savefig(f'{OUT}/q4c_distribution.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q4c_distribution.png')


# ════════════════════════════════════════════════════════════════
# Q5-A  |  גרף 7: הכנסות לפי אזור + שיעור מסירה (2 פאנלים)
# ════════════════════════════════════════════════════════════════
total_rev = df['Price'].sum()
area_rev  = df.groupby('AreaName').agg(
    revenue=('Price','sum'),
    orders=('Id','count'),
    delivered=('delivered','sum'),
).assign(
    pct=lambda x: x['revenue']/total_rev*100,
    delivery_rate=lambda x: x['delivered']/x['orders']*100,
).sort_values('revenue', ascending=True)   # מיון לפי הכנסה — שמור לשני הפאנלים

fig, (ax_rev, ax_del) = plt.subplots(1, 2, figsize=(13, 5.5),
                                      gridspec_kw={'width_ratios': [1.3, 1]})

# ── פאנל שמאל: הכנסות ──
rev_colors = [C_GREEN if r >= 90 else (C_ACCENT if r >= 70 else C_RED)
              for r in area_rev['delivery_rate']]
bars = ax_rev.barh(area_rev.index, area_rev['revenue'] / 1000,
                   color=rev_colors, height=0.6, zorder=3)
for bar, rev, pct in zip(bars, area_rev['revenue'], area_rev['pct']):
    ax_rev.text(bar.get_width() + 4, bar.get_y() + bar.get_height() / 2,
                f'₪{rev/1000:.0f}K  ({pct:.1f}%)',
                va='center', fontsize=9, color=C_MGREY)
ax_rev.set_xlim(0, 1550)
ax_rev.set_xlabel(heb('הכנסות (אלפי ₪)'), fontsize=11, color=C_MGREY)
ax_rev.set_title(heb('הכנסות לפי אזור'), fontsize=12,
                 fontweight='bold', color=C_BLUE, pad=10)
patches = [
    mpatches.Patch(color=C_GREEN,  label=heb('מסירה >90%')),
    mpatches.Patch(color=C_ACCENT, label=heb('מסירה 70-90%')),
    mpatches.Patch(color=C_RED,    label=heb('מסירה <70%')),
]
ax_rev.legend(handles=patches, fontsize=8.5, loc='lower right')
ax_rev.tick_params(axis='y', labelsize=9.5)

# ── פאנל ימין: שיעור מסירה (אותו סדר אזורים) ──
del_colors = [C_GREEN if r >= 90 else (C_ACCENT if r >= 70 else C_RED)
              for r in area_rev['delivery_rate']]
bars2 = ax_del.barh(area_rev.index, area_rev['delivery_rate'],
                    color=del_colors, height=0.6, zorder=3)
avg_del = df['delivered'].mean() * 100
ax_del.axvline(avg_del, color=C_GOLD, linestyle='--', linewidth=1.5,
               label=heb(f'ממוצע: {avg_del:.1f}%'))
for bar, r in zip(bars2, area_rev['delivery_rate']):
    ax_del.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height() / 2,
                f'{r:.1f}%', va='center', fontsize=9.5,
                fontweight='bold', color=C_BLUE)
ax_del.set_xlim(0, 118)
ax_del.set_xlabel(heb('שיעור מסירה (%)'), fontsize=11, color=C_MGREY)
ax_del.set_title(heb('שיעור מסירה לפי אזור'), fontsize=12,
                 fontweight='bold', color=C_BLUE, pad=10)
ax_del.legend(fontsize=8.5)
ax_del.set_yticklabels([])   # אותן תוויות כבר בפאנל שמאל

fig.suptitle(heb('הכנסות מול שיעור מסירה — לפי אזור'),
             fontsize=13, fontweight='bold', color=C_BLUE, y=1.01)
plt.tight_layout()
fig.savefig(f'{OUT}/q5a_revenue_area.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q5a_revenue_area.png')

# ════════════════════════════════════════════════════════════════
# Q5-B  |  גרף 8: הכנסה vs מסירה — scatter plot
# ════════════════════════════════════════════════════════════════
area_scatter = df.groupby('AreaName').agg(
    revenue=('Price','sum'),
    orders=('Id','count'),
    delivered=('delivered','sum'),
).assign(
    pct=lambda x: x['revenue']/total_rev*100,
    delivery_rate=lambda x: x['delivered']/x['orders']*100,
).reset_index()

fig, ax = plt.subplots(figsize=(8, 5))
sc = ax.scatter(area_scatter['delivery_rate'], area_scatter['revenue']/1000,
                s=area_scatter['orders']/20, alpha=0.75,
                c=area_scatter['revenue'], cmap='Blues', zorder=3,
                edgecolors=C_BLUE, linewidths=0.8)

for _, row in area_scatter.iterrows():
    ax.annotate(row['AreaName'].split('-')[0].strip(),
                xy=(row['delivery_rate'], row['revenue']/1000),
                xytext=(4, 4), textcoords='offset points',
                fontsize=8.5, color=C_MGREY)

# קו מגמה
z = np.polyfit(area_scatter['delivery_rate'], area_scatter['revenue']/1000, 1)
xline = np.linspace(area_scatter['delivery_rate'].min(), area_scatter['delivery_rate'].max(), 100)
ax.plot(xline, np.poly1d(z)(xline), color=C_RED, linestyle='--',
        linewidth=1.5, label=heb('מגמה (r = -0.62)'), zorder=2)

ax.set_xlabel(heb('שיעור מסירה (%)'), fontsize=11, color=C_MGREY)
ax.set_ylabel(heb('הכנסות (אלפי ₪)'), fontsize=11, color=C_MGREY)
ax.set_title(heb('קורלציה שלילית: הכנסה גבוהה = מסירה גרועה'),
             fontsize=12, fontweight='bold', color=C_BLUE, pad=12)
ax.legend(fontsize=9)
ax.text(0.97, 0.97, heb('גודל נקודה = נפח הזמנות'),
        transform=ax.transAxes, ha='right', va='top', fontsize=8.5, color=C_MGREY)
plt.tight_layout()
fig.savefig(f'{OUT}/q5b_scatter.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q5b_scatter.png')

# ════════════════════════════════════════════════════════════════
# Q5-C  |  גרף 9: TOP 10 עסקים לפי הכנסה
# ════════════════════════════════════════════════════════════════
biz_rev = df.groupby('BusinessId').agg(
    revenue=('Price','sum'), orders=('Id','count')
).assign(pct=lambda x: x['revenue']/total_rev*100)\
 .sort_values('revenue', ascending=False).head(10)

fig, ax = plt.subplots(figsize=(8, 4.5))
lbl    = [heb(f'עסק {int(b)}') for b in biz_rev.index]
colors = [C_RED if i < 3 else C_ACCENT for i in range(10)]
bars   = ax.barh(lbl[::-1], biz_rev['revenue'].values[::-1]/1000,
                 color=colors[::-1], height=0.6, zorder=3)

for bar, rev, pct in zip(bars, biz_rev['revenue'].values[::-1], biz_rev['pct'].values[::-1]):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
            heb(f'₪{rev/1000:.0f}K ({pct:.2f}%)'),
            va='center', fontsize=9.5, color=C_MGREY)

top3_pct = biz_rev.head(3)['pct'].sum()
ax.set_title(heb(f'TOP 10 עסקים לפי הכנסה | TOP 3 = {top3_pct:.1f}% מסה"כ'),
             fontsize=12, fontweight='bold', color=C_BLUE, pad=12)
ax.set_xlabel(heb('הכנסות (אלפי ₪)'), fontsize=11, color=C_MGREY)
ax.set_xlim(0, 130)
patches = [
    mpatches.Patch(color=C_RED,    label=heb('TOP 3 עסקים')),
    mpatches.Patch(color=C_ACCENT, label=heb('עסקים 4-10')),
]
ax.legend(handles=patches, fontsize=9)
ax.tick_params(axis='y', labelsize=10)
plt.tight_layout()
fig.savefig(f'{OUT}/q5c_top_businesses.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q5c_top_businesses.png')

# ════════════════════════════════════════════════════════════════
# Q6-A  |  גרף 10: שיעור חזרת לקוחות לפי אזור
# ════════════════════════════════════════════════════════════════
df_sorted = df.sort_values('OrderDate')
df_sorted['is_returning'] = df_sorted.duplicated(subset='UserId', keep='first')
area_ret = df_sorted.groupby('AreaName').agg(
    orders=('Id','count'), returning=('is_returning','sum')
).assign(ret_rate=lambda x: x['returning']/x['orders']*100)\
 .sort_values('ret_rate', ascending=True)

avg_ret = df_sorted['is_returning'].mean() * 100
fig, ax = plt.subplots(figsize=(8, 5))
colors  = [C_GREEN if r >= 30 else (C_ACCENT if r >= 22 else C_RED)
           for r in area_ret['ret_rate']]
bars    = ax.barh(area_ret.index, area_ret['ret_rate'],
                  color=colors, height=0.6, zorder=3)
ax.axvline(avg_ret, color=C_GOLD, linestyle='--', linewidth=1.5,
           label=heb(f'ממוצע: {avg_ret:.1f}%'))

for bar, r in zip(bars, area_ret['ret_rate']):
    ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
            f'{r:.1f}%', va='center', fontsize=10, fontweight='bold', color=C_BLUE)

ax.set_xlim(0, 45)
ax.set_xlabel(heb('שיעור הזמנות מלקוחות חוזרים (%)'), fontsize=11, color=C_MGREY)
ax.set_title(heb('שיעור לקוחות חוזרים לפי אזור'),
             fontsize=12, fontweight='bold', color=C_BLUE, pad=12)
ax.legend(fontsize=9)
ax.tick_params(axis='y', labelsize=10)
plt.tight_layout()
fig.savefig(f'{OUT}/q6a_returning_area.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q6a_returning_area.png')

# ════════════════════════════════════════════════════════════════
# Q6-B  |  גרף 11: קוהורט מרץ + שימור 30 יום לפי אזור
# ════════════════════════════════════════════════════════════════
df_sorted2 = df.sort_values('OrderDate').copy()
df_sorted2['is_first'] = ~df_sorted2.duplicated(subset='UserId', keep='first')
march_first = df_sorted2[
    df_sorted2['is_first'] &
    (df_sorted2['Month'] == pd.Period('2024-03'))
]['UserId'].unique()
cohort_df = df[df['UserId'].isin(march_first)].copy()
first_orders = cohort_df.groupby('UserId')['OrderDate'].min().reset_index()
first_orders.columns = ['UserId', 'first_date']
cohort_df = cohort_df.merge(first_orders, on='UserId')
cohort_df['days'] = (cohort_df['OrderDate'] - cohort_df['first_date']).dt.days
repeat_30 = cohort_df[(cohort_df['days']>0)&(cohort_df['days']<=30)]['UserId'].nunique()
repeat_60 = cohort_df[(cohort_df['days']>0)&(cohort_df['days']<=60)]['UserId'].nunique()
total_c   = len(march_first)

# retention by area
cust_first = df_sorted2[df_sorted2['is_first']][['UserId','OrderDate','AreaName']].copy()
cust_first.columns = ['UserId','first_date','first_area']
all_ord = df[['UserId','OrderDate']].copy()
ca = cust_first.merge(all_ord, on='UserId')
ca['days2'] = (ca['OrderDate'] - ca['first_date']).dt.days
ret30 = ca[(ca['days2']>0)&(ca['days2']<=30)].groupby('first_area')['UserId'].nunique()
tot_c = cust_first.groupby('first_area')['UserId'].nunique()
ret_area = (ret30/tot_c*100).sort_values(ascending=True).dropna()

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

# פאנל שמאל: קוהורט
labels    = [heb('לא חזרו\n(60 יום)'), heb('חזרו\n30–60 יום'), heb('חזרו\nתוך 30 יום')]
vals      = [total_c - repeat_60, repeat_60 - repeat_30, repeat_30]
clrs      = [C_RED, C_ACCENT, C_GREEN]
wedges, texts, autotexts = ax1.pie(
    vals, labels=labels, colors=clrs,
    autopct='%1.1f%%', startangle=90,
    textprops={'fontsize': 10},
    wedgeprops={'edgecolor': 'white', 'linewidth': 2}
)
ax1.set_title(heb(f'קוהורט מרץ 2024 | {total_c:,} לקוחות חדשים'),
              fontsize=11, fontweight='bold', color=C_BLUE, pad=12)

# פאנל ימין: שימור 30 יום לפי אזור
c_ret = [C_GREEN if r >= 15 else (C_ACCENT if r >= 12 else C_RED)
         for r in ret_area.values]
bars2 = ax2.barh(ret_area.index, ret_area.values, color=c_ret, height=0.6, zorder=3)
ax2.axvline(ret_area.mean(), color=C_GOLD, linestyle='--', linewidth=1.5,
            label=heb(f'ממוצע: {ret_area.mean():.1f}%'))
for bar, r in zip(bars2, ret_area.values):
    ax2.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height()/2,
             f'{r:.1f}%', va='center', fontsize=9.5, fontweight='bold', color=C_BLUE)
ax2.set_xlim(0, 28)
ax2.set_xlabel(heb('שיעור שימור תוך 30 יום (%)'), fontsize=10, color=C_MGREY)
ax2.set_title(heb('שיעור שימור 30 יום לפי אזור'),
              fontsize=11, fontweight='bold', color=C_BLUE, pad=12)
ax2.legend(fontsize=9)
ax2.tick_params(axis='y', labelsize=9.5)

plt.suptitle(heb('ניתוח שימור לקוחות'), fontsize=13, fontweight='bold', color=C_BLUE, y=1.01)
plt.tight_layout()
fig.savefig(f'{OUT}/q6b_retention.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q6b_retention.png')


# ════════════════════════════════════════════════════════════════
# Q7  |  גרף 12: מגמה שבועית — נפח + AOV + מסירה
# ════════════════════════════════════════════════════════════════
df['Week'] = df['OrderDate'].dt.to_period('W')
weekly = df.groupby('Week').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
    revenue=('Price','sum'),
).assign(
    delivery_rate=lambda x: x['delivered']/x['orders']*100,
    aov=lambda x: x['revenue']/x['orders'],
).reset_index()
weekly['week_num'] = range(len(weekly))
weeks_lbl = [str(w)[:10] for w in weekly['Week']]

fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(13, 7), sharex=True,
                                gridspec_kw={'height_ratios':[2,1]})

# נפח
bar_clr = [C_RED if r < 55 else (C_GOLD if r < 65 else C_ACCENT)
           for r in weekly['delivery_rate']]
ax1.bar(weekly['week_num'], weekly['orders'], color=bar_clr, alpha=0.8,
        width=0.7, zorder=3, label=heb('נפח הזמנות'))
ax1b = ax1.twinx()
ax1b.plot(weekly['week_num'], weekly['delivery_rate'], color=C_BLUE,
          linewidth=2.2, marker='o', markersize=5, zorder=4,
          label=heb('שיעור מסירה %'))
# סימון אנומליות
anomaly_weeks = [5, 6, 10]  # שבוע 04-10 מרץ, 11-17 מרץ, 08-14 אפריל
for wi in anomaly_weeks:
    if wi < len(weekly):
        ax1.axvspan(wi-0.4, wi+0.4, color=C_RED, alpha=0.12)
ax1.set_ylabel(heb('נפח הזמנות'), fontsize=11, color=C_MGREY)
ax1b.set_ylabel(heb('שיעור מסירה (%)'), fontsize=11, color=C_BLUE)
ax1b.set_ylim(35, 90)
ax1.set_title(heb('מגמה שבועית — נפח הזמנות ושיעור מסירה'),
              fontsize=13, fontweight='bold', color=C_BLUE, pad=10)
lines1, labs1 = ax1.get_legend_handles_labels()
lines2, labs2 = ax1b.get_legend_handles_labels()
ax1.legend(lines1+lines2, labs1+labs2, fontsize=9, loc='upper left')

# AOV
ax2.plot(weekly['week_num'], weekly['aov'], color=C_GOLD,
         linewidth=2.2, marker='s', markersize=5, zorder=3)
ax2.axhline(weekly['aov'].mean(), color=C_MGREY, linestyle='--', linewidth=1.2,
            label=heb(f'ממוצע ₪{weekly["aov"].mean():.2f}'))
ax2.set_ylabel(heb('AOV (₪)'), fontsize=11, color=C_GOLD)
ax2.set_xticks(weekly['week_num'])
ax2.set_xticklabels(weeks_lbl, rotation=45, ha='right', fontsize=7.5)
ax2.set_title(heb('AOV שבועי'), fontsize=10, fontweight='bold', color=C_GOLD)
ax2.legend(fontsize=9)

plt.tight_layout()
fig.savefig(f'{OUT}/q7a_weekly_trend.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q7a_weekly_trend.png')

# ════════════════════════════════════════════════════════════════
# Q7  |  גרף 13: דפוס יום בשבוע
# ════════════════════════════════════════════════════════════════
df['DayOfWeek'] = df['OrderDate'].dt.dayofweek
days_heb_short  = [heb('שני'),heb('שלישי'),heb('רביעי'),heb('חמישי'),
                   heb('שישי'),heb('שבת'),heb('ראשון')]
dow = df.groupby('DayOfWeek').agg(
    orders=('Id','count'),
    delivered=('delivered','sum'),
    revenue=('Price','sum'),
).assign(
    delivery_rate=lambda x: x['delivered']/x['orders']*100,
    aov=lambda x: x['revenue']/x['orders'],
)
n_weeks = df['OrderDate'].dt.to_period('W').nunique()
dow['daily_avg'] = dow['orders'] / (n_weeks)

fig, (axL, axR) = plt.subplots(1, 2, figsize=(12, 4.5))

# עמודות נפח
clr_dow = [C_GOLD if i in [4,5] else C_ACCENT for i in range(7)]
bars = axL.bar(days_heb_short, dow['daily_avg'], color=clr_dow, width=0.6, zorder=3)
for bar, v in zip(bars, dow['daily_avg']):
    axL.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1,
             f'{v:.0f}', ha='center', fontsize=10, fontweight='bold', color=C_BLUE)
axL.set_ylabel(heb('ממוצע הזמנות ליום'), fontsize=11, color=C_MGREY)
axL.set_title(heb('ממוצע הזמנות לפי יום בשבוע'), fontsize=12,
              fontweight='bold', color=C_BLUE, pad=10)
axL.set_ylim(0, 340)

# שיעור מסירה + AOV
ax_twin = axR.twinx()
c_del = [C_GREEN if r >= 67 else C_RED for r in dow['delivery_rate']]
bars2 = axR.bar(days_heb_short, dow['delivery_rate'], color=c_del,
                width=0.6, alpha=0.7, zorder=3)
ax_twin.plot(days_heb_short, dow['aov'], color=C_GOLD, linewidth=2.5,
             marker='D', markersize=7, zorder=4, label=heb('AOV (₪)'))
for bar, r in zip(bars2, dow['delivery_rate']):
    axR.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3,
             f'{r:.1f}%', ha='center', fontsize=9.5, fontweight='bold', color=C_BLUE)
axR.set_ylabel(heb('שיעור מסירה (%)'), fontsize=11, color=C_MGREY)
ax_twin.set_ylabel(heb('AOV (₪)'), fontsize=11, color=C_GOLD)
axR.set_ylim(55, 80)
ax_twin.set_ylim(110, 130)
axR.set_title(heb('מסירה ו-AOV לפי יום בשבוע'), fontsize=12,
              fontweight='bold', color=C_BLUE, pad=10)
ax_twin.legend(fontsize=9)

plt.suptitle(heb('דפוס התנהגות לפי יום בשבוע'), fontsize=13,
             fontweight='bold', color=C_BLUE, y=1.01)
plt.tight_layout()
fig.savefig(f'{OUT}/q7b_dow.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q7b_dow.png')

# ════════════════════════════════════════════════════════════════
# Q8  |  גרף 14: פיזור טוקנים לכל לקוח
# ════════════════════════════════════════════════════════════════
PATH2 = r'C:\Users\Yosef\PycharmProjects\HAAT_assignment\HAAT_DA_Dataset.xlsx'
pm    = pd.read_excel(PATH2, sheet_name='payment_methods_user')
tpu   = pm.groupby('UserId')['Token'].count()

fig, (axA, axB) = plt.subplots(1, 2, figsize=(12, 4.5))

# היסטוגרמה — עד 20 טוקנים
tpu_clip = tpu[tpu <= 20]
bins     = range(1, 22)
axA.hist(tpu_clip, bins=bins, color=C_ACCENT, alpha=0.85, edgecolor='white',
         linewidth=0.5, zorder=3)
axA.axvline(tpu.median(), color=C_GREEN, linestyle='-', linewidth=2,
            label=heb(f'חציון: {tpu.median():.0f}'))
axA.axvline(tpu.mean(), color=C_GOLD, linestyle='--', linewidth=2,
            label=heb(f'ממוצע: {tpu.mean():.1f}'))
n_multi = (tpu > 1).sum()
axA.set_xlabel(heb('מספר טוקנים ללקוח'), fontsize=11, color=C_MGREY)
axA.set_ylabel(heb('מספר לקוחות'), fontsize=11, color=C_MGREY)
axA.set_title(heb(f'פיזור טוקנים (1-20) | {n_multi:,} לקוחות עם >1 טוקן'),
              fontsize=11, fontweight='bold', color=C_BLUE)
axA.legend(fontsize=9)
axA.set_xticks(range(1, 21))

# עמודות השוואה: users/orders/payment
users_ord = df['UserId'].nunique()
users_pm  = pm['UserId'].nunique()
in_both   = len(set(pm['UserId']) & set(df['UserId']))
only_pm   = len(set(pm['UserId']) - set(df['UserId']))
only_ord  = len(set(df['UserId']) - set(pm['UserId']))
dup_tok   = pm['Token'].duplicated().sum()

cats   = [heb('לקוחות\nבהזמנות'), heb('לקוחות\nb-payment'), heb('בשתי\nהטבלאות'),
          heb('רק ב-\npayment'), heb('רק\nבהזמנות'), heb('טוקנים\nכפולים')]
vals   = [users_ord, users_pm, in_both, only_pm, only_ord, dup_tok]
clrs   = [C_ACCENT, C_ACCENT, C_GREEN, C_GOLD, C_GOLD, C_RED]
bars3  = axB.bar(cats, vals, color=clrs, width=0.6, zorder=3)
for bar, v in zip(bars3, vals):
    axB.text(bar.get_x()+bar.get_width()/2, bar.get_height()+50,
             f'{v:,}', ha='center', fontsize=10, fontweight='bold', color=C_BLUE)
axB.set_ylabel(heb('מספר רשומות'), fontsize=11, color=C_MGREY)
axB.set_title(heb('ניתוח איכות נתונים — השוואת טבלאות'),
              fontsize=11, fontweight='bold', color=C_BLUE)
axB.set_ylim(0, max(vals)*1.2)
axB.tick_params(axis='x', labelsize=9)

plt.suptitle(heb('Q8 — בדיקת איכות נתונים: payment_methods_user'),
             fontsize=13, fontweight='bold', color=C_BLUE, y=1.01)
plt.tight_layout()
fig.savefig(f'{OUT}/q8_data_quality.png', dpi=150, bbox_inches='tight')
plt.close()
print('saved: q8_data_quality.png')

print('\nכל הגרפים נשמרו בתיקייה:', OUT)
