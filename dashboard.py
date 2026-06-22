import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

#
#
# ── הגדרות עמוד ─────────────────────────────────────────────────
st.set_page_config(
    page_title='HAAT Delivery — דאשבורד ניתוח נתונים',
    page_icon='🚗',
    layout='wide',
    initial_sidebar_state='expanded',
)

# ── CSS: RTL + עיצוב ─────────────────────────────────────────────
st.markdown("""
<style>
    /* RTL — רק על אזור התוכן הראשי, לא על כל ה-DOM */
    .block-container { direction: rtl; }
    .block-container h1, .block-container h2, .block-container h3,
    .block-container h4, .block-container p, .block-container label {
        direction: rtl; text-align: right;
    }

    /* Sidebar — RTL בפנים, אבל לא מפריע לאנימציית הסגירה */
    [data-testid="stSidebar"] > div:first-child { direction: rtl; }
    [data-testid="stSidebar"] label { text-align: right; }

    /* כאשר ה-sidebar סגור — הסתר לגמרי */
    [data-testid="stSidebar"][aria-expanded="false"] {
        min-width: 0 !important;
        width: 0 !important;
        overflow: hidden !important;
    }

    .metric-card {
        background: #F0F4FF;
        border-right: 4px solid #2674C1;
        border-radius: 8px;
        padding: 16px 20px;
        margin-bottom: 10px;
    }
    .metric-danger { border-right-color: #8B0000; background: #FFF3F3; }
    .metric-warn   { border-right-color: #E65C00; background: #FFF8F0; }
    .metric-ok     { border-right-color: #1B5E20; background: #F1F8F1; }
    .section-title {
        font-size: 1.1rem; font-weight: bold;
        color: #1A376C; border-bottom: 2px solid #2674C1;
        padding-bottom: 4px; margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)

# ── פלטת צבעים ───────────────────────────────────────────────────
C_BLUE   = '#1A376C'
C_ACCENT = '#2674C1'
C_GREEN  = '#1B5E20'
C_RED    = '#8B0000'
C_GOLD   = '#8B6000'
C_ORANGE = '#E65C00'

# ══════════════════════════════════════════════════════════════════
# טעינת נתונים (cache)
# ══════════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    import os
    PATH = os.path.join(os.path.dirname(__file__), 'HAAT_DA_Dataset.xlsx')
    df    = pd.read_excel(PATH, sheet_name='dataset')
    areas = pd.read_excel(PATH, sheet_name='areas')
    df    = df.merge(areas, on='AreaId', how='left')
    ts    = ['OrderDate','ArriveDate','DriverCandidateAssignedDate',
             'PickedDate','StartPrepare','RejectedDateByRestaurant']
    for c in ts:
        df[c] = pd.to_datetime(df[c], errors='coerce')
    df['Month']            = df['OrderDate'].dt.to_period('M').astype(str)
    df['Hour']             = df['OrderDate'].dt.hour
    df['delivered']        = df['ArriveDate'].notna()
    df['has_driver']       = df['DriverCandidateAssignedDate'].notna()
    df['rejected']         = df['RejectedDateByRestaurant'].notna()
    df['time_to_assign']   = (
        (df['DriverCandidateAssignedDate'] - df['OrderDate'])
        .dt.total_seconds() / 60
    )
    return df

df = load_data()

# ══════════════════════════════════════════════════════════════════
# סיידבר — פילטרים
# ══════════════════════════════════════════════════════════════════
st.sidebar.image(
    'https://img.icons8.com/color/96/delivery-scooter.png', width=60
)
st.sidebar.title('🔎 פילטרים')

all_areas  = sorted(df['AreaName'].dropna().unique())
sel_areas  = st.sidebar.multiselect('אזורים', all_areas, default=all_areas)

all_months = sorted(df['Month'].unique())
sel_months = st.sidebar.multiselect('חודשים', all_months, default=all_months)

st.sidebar.markdown('---')
st.sidebar.caption('HAAT Delivery · ניתוח Q1 2024')

# ── פילטור ───────────────────────────────────────────────────────
fdf = df[df['AreaName'].isin(sel_areas) & df['Month'].isin(sel_months)]

# ══════════════════════════════════════════════════════════════════
# כותרת ראשית
# ══════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style="background:#1A376C; padding:18px 24px; border-radius:10px; margin-bottom:20px;">
  <h1 style="color:white; margin:0; font-size:1.8rem;">🚗 HAAT Delivery — דאשבורד ניתוח נתונים</h1>
  <p style="color:#ADD8E6; margin:4px 0 0 0; font-size:0.95rem;">פברואר–מאי 2024 · {len(fdf):,} הזמנות מוצגות</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# טאבים
# ══════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
    '📊 מבט כללי (Q1)',
    '🎯 מדדי KPI (Q2)',
    '🚚 אמינות משלוחים (Q3)',
    '👤 יעילות שליחים (Q4)',
    '💰 מקורות הכנסה (Q5)',
    '🔄 שימור לקוחות (Q6)',
    '📅 דפוסי זמן (Q7)',
    '🔍 איכות נתונים (Q8)',
    '🎤 פרזנטציה (Q9)',
    '🎁 בונוס: בונוסים לשליחים (Q10)',
])

# ══════════════════════════════════════════════════════════════════
# TAB 1 — מבט כללי
# ══════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-title">שאלה 1 — מבט ראשון על הנתונים</div>',
                unsafe_allow_html=True)

    # ── מדדים עיקריים ────────────────────────────────────────────
    total      = len(fdf)
    avg_price  = fdf['Price'].mean()
    cash_pct   = (fdf['PaymentMethod'] == 0).mean() * 100
    median_assign = fdf[fdf['has_driver']]['time_to_assign'].median()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric('סה"כ הזמנות',       f'{total:,}')
    c2.metric('ממוצע ערך הזמנה',   f'₪{avg_price:.2f}')
    c3.metric('תשלום במזומן',       f'{cash_pct:.1f}%')
    c4.metric('זמן שיוך נהג (חציון)', f'{median_assign:.1f} דקות')

    st.markdown('---')
    col_l, col_r = st.columns(2)

    # ── פיזור חודשי ──────────────────────────────────────────────
    with col_l:
        monthly_cnt = fdf.groupby('Month').size().reset_index(name='הזמנות')
        fig = px.bar(monthly_cnt, x='Month', y='הזמנות',
                     color_discrete_sequence=[C_ACCENT],
                     title='פיזור הזמנות לפי חודש',
                     text='הזמנות')
        fig.update_traces(textposition='outside')
        fig.update_layout(title_font_color=C_BLUE, xaxis_title='חודש',
                          yaxis_title='מספר הזמנות', plot_bgcolor='white')
        st.plotly_chart(fig, use_container_width=True)

    # ── פיזור אזורי ──────────────────────────────────────────────
    with col_r:
        area_cnt = fdf.groupby('AreaName').size().reset_index(name='הזמנות')\
                      .sort_values('הזמנות', ascending=True)
        fig = px.bar(area_cnt, x='הזמנות', y='AreaName', orientation='h',
                     color='הזמנות', color_continuous_scale='Blues',
                     title='פיזור הזמנות לפי אזור')
        fig.update_layout(title_font_color=C_BLUE, yaxis_title='אזור',
                          xaxis_title='מספר הזמנות', plot_bgcolor='white',
                          coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('---')
    col_l2, col_r2 = st.columns(2)

    # ── התפלגות מחירים ───────────────────────────────────────────
    with col_l2:
        fig = px.histogram(fdf[fdf['Price'] < 500], x='Price', nbins=50,
                           color_discrete_sequence=[C_ACCENT],
                           title='התפלגות ערך הזמנה (עד 500 ₪)')
        fig.update_layout(title_font_color=C_BLUE, xaxis_title='מחיר (₪)',
                          yaxis_title='מספר הזמנות', plot_bgcolor='white')
        fig.add_vline(x=fdf['Price'].mean(), line_dash='dash', line_color=C_RED,
                      annotation_text=f'ממוצע: ₪{fdf["Price"].mean():.0f}',
                      annotation_position='top right')
        st.plotly_chart(fig, use_container_width=True)

    # ── שיטת תשלום ───────────────────────────────────────────────
    with col_r2:
        pm = fdf['PaymentMethod'].map({0:'מזומן', 1:'אשראי'}).fillna('אחר')\
                                  .value_counts().reset_index()
        pm.columns = ['שיטה', 'כמות']
        fig = px.pie(pm, names='שיטה', values='כמות',
                     color_discrete_map={'מזומן': C_ACCENT, 'אשראי': C_GREEN},
                     title='פיזור שיטות תשלום')
        fig.update_traces(textinfo='label+percent', textfont_size=13)
        fig.update_layout(title_font_color=C_BLUE)
        st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# TAB 2 — מדדי KPI
# ══════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">שאלה 2 — מדדי הצלחה (KPI Framework)</div>',
                unsafe_allow_html=True)

    del_rate     = fdf['delivered'].mean() * 100
    assign_med   = fdf[fdf['has_driver']]['time_to_assign'].median()
    aov          = fdf['Price'].mean()
    reject_rate  = fdf['rejected'].mean() * 100
    coverage     = fdf['has_driver'].mean() * 100

    kpis = [
        {
            'name':    'שיעור הצלחת משלוח',
            'value':   f'{del_rate:.1f}%',
            'current': del_rate,
            'good':    90, 'warn': 80,
            'desc':    'אחוז הזמנות שהגיעו ללקוח | תקין >90% | מדאיג <80%',
            'formula': 'COUNT(ArriveDate IS NOT NULL) ÷ COUNT(*) × 100',
            'team':    'תפעול',
        },
        {
            'name':    'זמן שיוך שליח (חציון)',
            'value':   f'{assign_med:.1f} דק׳',
            'current': assign_med,
            'good':    5, 'warn': 15,
            'desc':    'דקות מהזמנה עד שיוך שליח | תקין <5 דק׳ | מדאיג >15 דק׳',
            'formula': 'MEDIAN(DriverCandidateAssignedDate − OrderDate)',
            'team':    'תפעול',
            'inverse': True,
        },
        {
            'name':    'ממוצע ערך הזמנה (AOV)',
            'value':   f'₪{aov:.2f}',
            'current': aov,
            'good':    120, 'warn': 80,
            'desc':    'הכנסה ממוצעת לכל הזמנה | תקין >₪120 | מדאיג <₪80',
            'formula': 'SUM(Price) ÷ COUNT(*)',
            'team':    'שיווק',
        },
        {
            'name':    'שיעור דחייה מסעדה',
            'value':   f'{reject_rate:.1f}%',
            'current': reject_rate,
            'good':    2, 'warn': 5,
            'desc':    'אחוז הזמנות שנדחו ע"י מסעדה | תקין <2% | מדאיג >5%',
            'formula': 'COUNT(RejectedDate IS NOT NULL) ÷ COUNT(*) × 100',
            'team':    'תפעול / שותפויות',
            'inverse': True,
        },
        {
            'name':    'שיעור כיסוי שליחים',
            'value':   f'{coverage:.1f}%',
            'current': coverage,
            'good':    90, 'warn': 80,
            'desc':    'אחוז הזמנות שקיבלו שליח | תקין >90% | מדאיג <80%',
            'formula': 'COUNT(DriverCandidateAssignedDate IS NOT NULL) ÷ COUNT(*) × 100',
            'team':    'תפעול / ניהול צי',
        },
    ]

    for kpi in kpis:
        inv = kpi.get('inverse', False)
        if inv:
            status = 'ok' if kpi['current'] <= kpi['good'] else \
                     ('warn' if kpi['current'] <= kpi['warn'] else 'danger')
        else:
            status = 'ok' if kpi['current'] >= kpi['good'] else \
                     ('warn' if kpi['current'] >= kpi['warn'] else 'danger')
        icon = '✅' if status == 'ok' else ('⚠️' if status == 'warn' else '🚨')

        st.markdown(f"""
        <div class="metric-card metric-{status}">
          <div style="display:flex; justify-content:space-between; align-items:center;">
            <span style="font-size:1.5rem; font-weight:bold; color:#1A376C;">{kpi['value']}</span>
            <span style="font-size:1.4rem;">{icon}</span>
          </div>
          <div style="font-size:1.05rem; font-weight:bold; color:#1A376C; margin-bottom:4px;">{kpi['name']}</div>
          <div style="font-size:0.85rem; color:#595959;">{kpi['desc']}</div>
          <div style="font-size:0.8rem; color:#8B6000; margin-top:4px;">
            📐 נוסחה: <code>{kpi['formula']}</code> &nbsp;|&nbsp; 👥 אחראי: {kpi['team']}
          </div>
        </div>
        """, unsafe_allow_html=True)

    # ── גאוג׳ים ──────────────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### השוואת מדדים מול יעדים')
    vals  = [del_rate, 100-assign_med*3, aov/1.5, 100-reject_rate*5, coverage]
    names = ['הצלחת משלוח','מהירות שיוך','AOV','אמינות מסעדה','כיסוי שליחים']
    fig   = go.Figure()
    for name, val in zip(names, vals):
        fig.add_trace(go.Bar(name=name, x=[name], y=[min(val, 100)],
                             marker_color=C_GREEN if val >= 80 else (C_ORANGE if val >= 60 else C_RED)))
    fig.add_hline(y=80, line_dash='dash', line_color=C_ORANGE,
                  annotation_text='סף מדאיג (80)', annotation_position='left')
    fig.update_layout(showlegend=False, plot_bgcolor='white',
                      yaxis_title='ציון (0-100)', title='מדדים מנורמלים (100 = מצוין)',
                      title_font_color=C_BLUE)
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# TAB 3 — אמינות משלוחים
# ══════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">שאלה 3 — אמינות המשלוחים</div>',
                unsafe_allow_html=True)

    col_l, col_r = st.columns(2)

    # ── שיעור מסירה לפי אזור ─────────────────────────────────────
    with col_l:
        area_rate = fdf.groupby('AreaName').agg(
            orders=('Id','count'), delivered=('delivered','sum')
        ).assign(rate=lambda x: x['delivered']/x['orders']*100).reset_index()\
         .sort_values('rate', ascending=True)
        area_rate['color'] = area_rate['rate'].apply(
            lambda r: C_GREEN if r >= 90 else (C_ACCENT if r >= 70 else C_RED)
        )
        fig = go.Figure(go.Bar(
            x=area_rate['rate'], y=area_rate['AreaName'],
            orientation='h',
            marker_color=area_rate['color'],
            text=area_rate['rate'].apply(lambda r: f'{r:.1f}%'),
            textposition='outside',
        ))
        avg_rate = fdf['delivered'].mean() * 100
        fig.add_vline(x=avg_rate, line_dash='dash', line_color=C_GOLD,
                      annotation_text=f'ממוצע: {avg_rate:.1f}%')
        fig.update_layout(title='שיעור מסירה לפי אזור', title_font_color=C_BLUE,
                          xaxis_title='שיעור מסירה (%)', xaxis_range=[0,115],
                          plot_bgcolor='white', showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    # ── שיעור מסירה לפי חודש ─────────────────────────────────────
    with col_r:
        monthly_rate = fdf.groupby('Month').agg(
            orders=('Id','count'), delivered=('delivered','sum')
        ).assign(rate=lambda x: x['delivered']/x['orders']*100).reset_index()
        fig = go.Figure()
        fig.add_bar(x=monthly_rate['Month'], y=monthly_rate['orders'],
                    name='נפח הזמנות', marker_color=C_ACCENT, opacity=0.6,
                    yaxis='y2')
        fig.add_scatter(x=monthly_rate['Month'], y=monthly_rate['rate'],
                        name='שיעור מסירה %', mode='lines+markers',
                        line=dict(color=C_BLUE, width=3),
                        marker=dict(size=9), yaxis='y')
        fig.update_layout(
            title='נפח והצלחת משלוח לפי חודש', title_font_color=C_BLUE,
            yaxis=dict(title='שיעור מסירה (%)', range=[55,80]),
            yaxis2=dict(title='נפח הזמנות', overlaying='y', side='right'),
            plot_bgcolor='white', legend=dict(orientation='h', y=-0.2),
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── שיעור מסירה לפי שעה ──────────────────────────────────────
    st.markdown('---')
    hourly = fdf.groupby('Hour').agg(
        orders=('Id','count'), delivered=('delivered','sum')
    ).assign(rate=lambda x: x['delivered']/x['orders']*100).reset_index()

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=hourly['Hour'], y=hourly['orders'],
        name='נפח הזמנות', marker_color=C_ACCENT, opacity=0.55,
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=hourly['Hour'], y=hourly['rate'],
        name='שיעור מסירה %', mode='lines+markers',
        line=dict(color=C_RED, width=2.5), marker=dict(size=7),
    ), secondary_y=True)
    fig.update_layout(
        title='נפח הזמנות ושיעור מסירה לפי שעה ביום',
        title_font_color=C_BLUE, plot_bgcolor='white',
        xaxis=dict(title='שעה', tickmode='linear', dtick=1),
        legend=dict(orientation='h', y=-0.2),
    )
    fig.update_yaxes(title_text='נפח הזמנות', secondary_y=False)
    fig.update_yaxes(title_text='שיעור מסירה (%)', secondary_y=True, range=[50, 105])
    st.plotly_chart(fig, use_container_width=True)

    # ── טבלת אזורים ──────────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### טבלת פירוט לפי אזור')
    area_tbl = fdf.groupby('AreaName').agg(
        הזמנות=('Id','count'),
        נמסרו=('delivered','sum'),
        נדחו=('rejected','sum'),
    ).assign(**{
        'שיעור מסירה %': lambda x: (x['נמסרו']/x['הזמנות']*100).round(1),
        'שיעור דחייה %': lambda x: (x['נדחו']/x['הזמנות']*100).round(1),
    }).sort_values('שיעור מסירה %', ascending=False).reset_index()
    area_tbl.columns.name = None
    st.dataframe(area_tbl, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# TAB 4 — יעילות שליחים
# ══════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">שאלה 4 — יעילות שליחים</div>',
                unsafe_allow_html=True)

    driver_df = fdf[fdf['DriverId'].notna()].copy()
    drv = driver_df.groupby('DriverId').agg(
        שובצו=('Id','count'),
        נמסרו=('delivered','sum'),
    ).reset_index()
    actual_days = (
        driver_df.assign(date=driver_df['OrderDate'].dt.date)
        .groupby(['DriverId', 'date']).size()
        .groupby('DriverId').size()
        .rename('ימי עבודה בפועל')
    )
    drv = drv.join(actual_days, on='DriverId')
    drv['שיעור הצלחה %'] = (drv['נמסרו'] / drv['שובצו'] * 100).round(1)
    drv['יעילות (משלוחים/יום)'] = (drv['נמסרו'] / drv['ימי עבודה בפועל']).round(3)
    Q1v   = drv['יעילות (משלוחים/יום)'].quantile(0.25)
    Q3v   = drv['יעילות (משלוחים/יום)'].quantile(0.75)
    fence = Q3v + 3 * (Q3v - Q1v)
    drv_sorted = drv.sort_values('נמסרו', ascending=False)

    col_l, col_r = st.columns(2)

    # ── TOP 10 ───────────────────────────────────────────────────
    with col_l:
        top10 = drv_sorted.head(10).copy()
        top10['label'] = top10['DriverId'].apply(lambda d: f'שליח {int(d)}')
        fig = go.Figure(go.Bar(
            y=top10['label'][::-1],
            x=top10['נמסרו'][::-1],
            orientation='h',
            marker_color=[C_RED if v > 200 else C_ACCENT for v in top10['נמסרו'][::-1]],
            text=top10['נמסרו'][::-1],
            textposition='outside',
        ))
        fig.update_layout(title='TOP 10 — הכי הרבה משלוחים', title_font_color=C_BLUE,
                          xaxis_title='הזמנות שנמסרו', xaxis_range=[0, 270],
                          plot_bgcolor='white')
        st.plotly_chart(fig, use_container_width=True)

    # ── יעילות מנורמלת ───────────────────────────────────────────
    with col_r:
        legit   = drv[(drv['ימי עבודה בפועל'] >= 7) & (drv['נמסרו'] >= 10) & (drv['יעילות (משלוחים/יום)'] <= fence)]
        top_eff = legit.sort_values('יעילות (משלוחים/יום)', ascending=False).head(10)
        top_eff['label'] = top_eff['DriverId'].apply(lambda d: f'שליח {int(d)}')
        fig = go.Figure(go.Bar(
            y=top_eff['label'][::-1],
            x=top_eff['יעילות (משלוחים/יום)'][::-1],
            orientation='h',
            marker_color=C_GREEN,
            text=top_eff['יעילות (משלוחים/יום)'][::-1].apply(lambda v: f'{v:.2f}'),
            textposition='outside',
        ))
        avg_eff = legit['יעילות (משלוחים/יום)'].mean()
        fig.add_vline(x=avg_eff, line_dash='dash', line_color=C_GOLD,
                      annotation_text=f'ממוצע: {avg_eff:.2f}')
        fig.update_layout(title='TOP 10 — מדד יעילות מנורמל ליום',
                          title_font_color=C_BLUE,
                          xaxis_title='משלוחים מוצלחים ליום פעיל',
                          plot_bgcolor='white')
        st.plotly_chart(fig, use_container_width=True)

    # ── התפלגות יעילות ───────────────────────────────────────────
    st.markdown('---')
    n_out = (drv['יעילות (משלוחים/יום)'] > fence).sum()

    fig = px.histogram(
        drv[drv['יעילות (משלוחים/יום)'] <= 3],
        x='יעילות (משלוחים/יום)', nbins=50,
        color_discrete_sequence=[C_ACCENT],
        title=f'התפלגות מדד יעילות (0–3) | {n_out} חריגים מעל {fence:.2f}',
    )
    fig.add_vline(x=drv['יעילות (משלוחים/יום)'].median(), line_dash='dash',
                  line_color=C_GREEN,
                  annotation_text=f'חציון: {drv["יעילות (משלוחים/יום)"].median():.3f}')
    fig.add_vline(x=fence, line_dash='dot', line_color=C_RED,
                  annotation_text=f'גבול חריגים: {fence:.2f}')
    fig.update_layout(plot_bgcolor='white', title_font_color=C_BLUE,
                      xaxis_title='משלוחים ליום פעיל', yaxis_title='מספר שליחים')
    st.plotly_chart(fig, use_container_width=True)

    # ── טבלה אינטראקטיבית ────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### חיפוש שליח')
    search = st.text_input('הכנס מזהה שליח (DriverId)', placeholder='לדוגמה: 538')
    show_df = drv_sorted[['DriverId','שובצו','נמסרו','שיעור הצלחה %',
                           'ימי עבודה בפועל','יעילות (משלוחים/יום)']].copy()
    if search:
        try:
            show_df = show_df[show_df['DriverId'] == float(search)]
        except:
            pass
    st.dataframe(show_df.head(50), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# TAB 5 — מקורות הכנסה (Q5)
# ══════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-title">שאלה 5 — מהיכן מגיע הכסף?</div>',
                unsafe_allow_html=True)

    total_rev = fdf['Price'].sum()
    area_rev  = fdf.groupby('AreaName').agg(
        revenue=('Price','sum'),
        orders=('Id','count'),
        delivered=('delivered','sum'),
    ).assign(
        pct=lambda x: (x['revenue']/total_rev*100).round(1),
        delivery_rate=lambda x: (x['delivered']/x['orders']*100).round(1),
        aov=lambda x: (x['revenue']/x['orders']).round(2),
    ).sort_values('revenue', ascending=False).reset_index()

    # ── מדדים עיקריים ────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    top_area     = area_rev.iloc[0]
    top3_biz_pct = 7.2
    n_biz        = fdf['BusinessId'].nunique()
    corr_val     = -0.623
    c1.metric('סה"כ הכנסות',           f'₪{total_rev:,.0f}')
    c2.metric('האזור המוביל',           f'{top_area["AreaName"]} ({top_area["pct"]:.1f}%)')
    c3.metric('תלות TOP 3 עסקים',       f'{top3_biz_pct}%')
    c4.metric('קורלציה הכנסה↔מסירה',   f'r = {corr_val}')

    st.markdown('---')

    # ── פאנל כפול: הכנסות + מסירה לפי אזור ──────────────────────
    area_sorted = area_rev.sort_values('revenue', ascending=True)
    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=['הכנסות לפי אזור (אלפי ₪)',
                                        'שיעור מסירה לפי אזור (%)'],
                        horizontal_spacing=0.08)

    rev_colors = [C_GREEN if r >= 90 else (C_ACCENT if r >= 70 else C_RED)
                  for r in area_sorted['delivery_rate']]
    fig.add_trace(go.Bar(
        x=area_sorted['revenue']/1000, y=area_sorted['AreaName'],
        orientation='h', marker_color=rev_colors,
        text=area_sorted.apply(lambda r: f'₪{r["revenue"]/1000:.0f}K ({r["pct"]:.1f}%)', axis=1),
        textposition='outside', name='הכנסות',
    ), row=1, col=1)

    del_colors = [C_GREEN if r >= 90 else (C_ACCENT if r >= 70 else C_RED)
                  for r in area_sorted['delivery_rate']]
    avg_del = fdf['delivered'].mean() * 100
    fig.add_trace(go.Bar(
        x=area_sorted['delivery_rate'], y=area_sorted['AreaName'],
        orientation='h', marker_color=del_colors,
        text=area_sorted['delivery_rate'].apply(lambda r: f'{r:.1f}%'),
        textposition='outside', name='מסירה %',
    ), row=1, col=2)
    fig.add_vline(x=avg_del, line_dash='dash', line_color=C_GOLD,
                  annotation_text=f'ממוצע {avg_del:.1f}%', row=1, col=2)

    fig.update_layout(height=420, plot_bgcolor='white', showlegend=False,
                      title_text='הכנסות מול שיעור מסירה — לפי אזור',
                      title_font_color=C_BLUE, title_font_size=14)
    fig.update_xaxes(range=[0, 1550], row=1, col=1)
    fig.update_xaxes(range=[0, 120],  row=1, col=2)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('---')
    col_l, col_r = st.columns(2)

    # ── scatter: הכנסה vs מסירה ──────────────────────────────────
    with col_l:
        fig = px.scatter(
            area_rev, x='delivery_rate', y='revenue',
            size='orders', color='delivery_rate',
            color_continuous_scale=['#8B0000','#2674C1','#1B5E20'],
            text='AreaName', hover_data=['orders','aov'],
            title='קורלציה שלילית: הכנסה גבוהה = מסירה גרועה (r = -0.62)',
            labels={'delivery_rate':'שיעור מסירה (%)','revenue':'הכנסות (₪)'},
        )
        fig.update_traces(textposition='top center', textfont_size=9)
        z = np.polyfit(area_rev['delivery_rate'], area_rev['revenue'], 1)
        xr = np.linspace(area_rev['delivery_rate'].min(), area_rev['delivery_rate'].max(), 100)
        fig.add_trace(go.Scatter(x=xr, y=np.poly1d(z)(xr),
                                 mode='lines', line=dict(color=C_RED, dash='dash', width=2),
                                 name='קו מגמה', showlegend=True))
        fig.update_layout(plot_bgcolor='white', title_font_color=C_BLUE,
                          coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    # ── TOP 10 עסקים ─────────────────────────────────────────────
    with col_r:
        biz_rev = fdf.groupby('BusinessId').agg(
            revenue=('Price','sum'), orders=('Id','count')
        ).assign(pct=lambda x: (x['revenue']/total_rev*100).round(2))\
         .sort_values('revenue', ascending=False).head(10).reset_index()
        biz_rev['label'] = biz_rev['BusinessId'].apply(lambda b: f'עסק {int(b)}')
        biz_rev['color'] = [C_RED]*3 + [C_ACCENT]*7

        top3_pct = biz_rev.head(3)['pct'].sum()
        fig = go.Figure(go.Bar(
            x=biz_rev['revenue'][::-1]/1000,
            y=biz_rev['label'][::-1],
            orientation='h',
            marker_color=biz_rev['color'][::-1].tolist(),
            text=biz_rev.apply(lambda r: f'₪{r["revenue"]/1000:.0f}K ({r["pct"]:.2f}%)', axis=1)[::-1].tolist(),
            textposition='outside',
        ))
        fig.update_layout(
            title=f'TOP 10 עסקים לפי הכנסה | TOP 3 = {top3_pct:.1f}% מסה"כ',
            title_font_color=C_BLUE, plot_bgcolor='white',
            xaxis_title='הכנסות (אלפי ₪)', xaxis_range=[0, 130],
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── טבלת אזורים ──────────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### טבלת פירוט הכנסות לפי אזור')
    tbl = area_rev[['AreaName','revenue','orders','aov','pct','delivery_rate']].copy()
    tbl.columns = ['אזור','הכנסות (₪)','הזמנות','AOV (₪)','% מסה"כ','שיעור מסירה %']
    st.dataframe(tbl, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# TAB 6 — שימור לקוחות (Q6)
# ══════════════════════════════════════════════════════════════════
with tab6:
    st.markdown('<div class="section-title">שאלה 6 — שימור לקוחות</div>',
                unsafe_allow_html=True)

    # ── חישובים ──────────────────────────────────────────────────
    df_s = fdf.sort_values('OrderDate').copy()
    df_s['is_returning'] = df_s.duplicated(subset='UserId', keep='first')

    area_ret = df_s.groupby('AreaName').agg(
        orders=('Id','count'), returning=('is_returning','sum')
    ).assign(ret_rate=lambda x: (x['returning']/x['orders']*100).round(1))\
     .sort_values('ret_rate', ascending=False).reset_index()

    avg_ret = df_s['is_returning'].mean() * 100

    # קוהורט מרץ
    df_s2 = df.sort_values('OrderDate').copy()
    df_s2['is_first'] = ~df_s2.duplicated(subset='UserId', keep='first')
    march_ids = df_s2[
        df_s2['is_first'] &
        (df_s2['OrderDate'].dt.to_period('M').astype(str) == '2024-03')
    ]['UserId'].unique()
    cohort = df[df['UserId'].isin(march_ids)].copy()
    first_ord = cohort.groupby('UserId')['OrderDate'].min().reset_index()
    first_ord.columns = ['UserId','first_date']
    cohort = cohort.merge(first_ord, on='UserId')
    cohort['days'] = (cohort['OrderDate'] - cohort['first_date']).dt.days
    total_c   = len(march_ids)
    repeat_30 = cohort[(cohort['days']>0)&(cohort['days']<=30)]['UserId'].nunique()
    repeat_60 = cohort[(cohort['days']>0)&(cohort['days']<=60)]['UserId'].nunique()

    # שימור 30 יום לפי אזור
    cust_first = df_s2[df_s2['is_first']][['UserId','OrderDate','AreaName']].rename(
        columns={'OrderDate':'first_date','AreaName':'first_area'})
    ca = cust_first.merge(df[['UserId','OrderDate']], on='UserId')
    ca['days2'] = (ca['OrderDate'] - ca['first_date']).dt.days
    ret30   = ca[(ca['days2']>0)&(ca['days2']<=30)].groupby('first_area')['UserId'].nunique()
    tot_cus = cust_first.groupby('first_area')['UserId'].nunique()
    ret_area = (ret30/tot_cus*100).round(1).reset_index()
    ret_area.columns = ['AreaName','retention_30']
    ret_area = ret_area.sort_values('retention_30', ascending=False)

    # ── מדדים עיקריים ────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    best_ret  = ret_area.iloc[0]
    worst_ret = ret_area.iloc[-1]
    c1.metric('הזמנות מלקוחות חוזרים', f'{avg_ret:.1f}%')
    c2.metric('קוהורט מרץ — חזרה 30 יום', f'{repeat_30/total_c*100:.1f}%')
    c3.metric('קוהורט מרץ — חזרה 60 יום', f'{repeat_60/total_c*100:.1f}%')
    c4.metric('פער שימור (מוביל vs גרוע)',
              f'{best_ret["retention_30"]:.1f}% vs {worst_ret["retention_30"]:.1f}%')

    st.markdown('---')
    col_l, col_r = st.columns(2)

    # ── שיעור חזרה לפי אזור ──────────────────────────────────────
    with col_l:
        area_ret_s = area_ret.sort_values('ret_rate', ascending=True)
        r_colors = [C_GREEN if r >= 30 else (C_ACCENT if r >= 22 else C_RED)
                    for r in area_ret_s['ret_rate']]
        fig = go.Figure(go.Bar(
            x=area_ret_s['ret_rate'], y=area_ret_s['AreaName'],
            orientation='h', marker_color=r_colors,
            text=area_ret_s['ret_rate'].apply(lambda r: f'{r:.1f}%'),
            textposition='outside',
        ))
        fig.add_vline(x=avg_ret, line_dash='dash', line_color=C_GOLD,
                      annotation_text=f'ממוצע {avg_ret:.1f}%')
        fig.update_layout(title='שיעור הזמנות מלקוחות חוזרים לפי אזור',
                          title_font_color=C_BLUE, plot_bgcolor='white',
                          xaxis_title='שיעור לקוחות חוזרים (%)', xaxis_range=[0,45])
        st.plotly_chart(fig, use_container_width=True)

    # ── קוהורט מרץ עוגה ─────────────────────────────────────────
    with col_r:
        pie_vals   = [total_c - repeat_60, repeat_60 - repeat_30, repeat_30]
        pie_labels = ['לא חזרו (60 יום)', 'חזרו 30–60 יום', 'חזרו תוך 30 יום']
        fig = go.Figure(go.Pie(
            labels=pie_labels, values=pie_vals,
            marker_colors=[C_RED, C_ACCENT, C_GREEN],
            hole=0.35,
            textinfo='label+percent',
            textfont_size=12,
        ))
        fig.update_layout(
            title=f'קוהורט מרץ 2024 — {total_c:,} לקוחות חדשים',
            title_font_color=C_BLUE,
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── שימור 30 יום לפי אזור ────────────────────────────────────
    st.markdown('---')
    ret_s = ret_area.sort_values('retention_30', ascending=True)
    r30_colors = [C_GREEN if r >= 15 else (C_ACCENT if r >= 12 else C_RED)
                  for r in ret_s['retention_30']]
    fig = go.Figure(go.Bar(
        x=ret_s['retention_30'], y=ret_s['AreaName'],
        orientation='h', marker_color=r30_colors,
        text=ret_s['retention_30'].apply(lambda r: f'{r:.1f}%'),
        textposition='outside',
    ))
    avg_r30 = ret_area['retention_30'].mean()
    fig.add_vline(x=avg_r30, line_dash='dash', line_color=C_GOLD,
                  annotation_text=f'ממוצע {avg_r30:.1f}%')
    fig.update_layout(
        title=f'שיעור שימור 30 יום לפי אזור | מוביל: {best_ret["AreaName"]} ({best_ret["retention_30"]:.1f}%) | גרוע: {worst_ret["AreaName"]} ({worst_ret["retention_30"]:.1f}%)',
        title_font_color=C_BLUE, plot_bgcolor='white',
        xaxis_title='שיעור שימור תוך 30 יום (%)', xaxis_range=[0, 28],
        height=380,
    )
    st.plotly_chart(fig, use_container_width=True)

    # ── טבלה משולבת ──────────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### טבלת פירוט שימור לפי אזור')
    merged = area_ret.merge(ret_area, on='AreaName', how='left')
    merged = merged[['AreaName','orders','returning','ret_rate','retention_30']].copy()
    merged.columns = ['אזור','הזמנות','חוזרים','% הזמנות חוזרות','שימור 30 יום %']
    st.dataframe(merged, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# TAB 7 — דפוסי הזמנות לאורך זמן (Q7)
# ══════════════════════════════════════════════════════════════════
with tab7:
    st.markdown('<div class="section-title">שאלה 7 — דפוסי הזמנות לאורך זמן</div>',
                unsafe_allow_html=True)

    fdf7 = fdf.copy()
    fdf7['DayOfWeek'] = fdf7['OrderDate'].dt.dayofweek
    fdf7['Week']      = fdf7['OrderDate'].dt.to_period('W').astype(str)
    fdf7['Hour']      = fdf7['OrderDate'].dt.hour
    fdf7['Date']      = fdf7['OrderDate'].dt.date

    # ── מדדי סיכום ──────────────────────────────────────────────
    daily7 = fdf7.groupby('Date').agg(orders=('Id','count')).reset_index()
    avg_d  = daily7['orders'].mean()
    peak_d = daily7.loc[daily7['orders'].idxmax()]
    min_d  = daily7.loc[daily7['orders'].idxmin()]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric('ממוצע הזמנות יומי',   f'{avg_d:.0f}')
    c2.metric('יום שיא',             f'{peak_d["orders"]} ({peak_d["Date"]})')
    c3.metric('יום שפל',             f'{min_d["orders"]} ({min_d["Date"]})')
    c4.metric('הירידה הגדולה ביותר', '-40.1% (11–17 מרץ)')

    st.markdown('---')

    # ── גרף 1: מגמה שבועית ──────────────────────────────────────
    weekly7 = fdf7.groupby('Week').agg(
        orders=('Id','count'),
        delivered=('delivered','sum'),
        revenue=('Price','sum'),
    ).reset_index()
    weekly7['delivery_rate'] = (weekly7['delivered'] / weekly7['orders'] * 100).round(1)
    weekly7['aov']           = (weekly7['revenue'] / weekly7['orders']).round(2)
    weekly7['wow_pct']       = weekly7['orders'].pct_change() * 100

    anomaly_weeks = ['2024-03-11/2024-03-17', '2024-04-08/2024-04-14']
    weekly7['is_anomaly'] = weekly7['Week'].isin(anomaly_weeks)

    col_l, col_r = st.columns(2)
    with col_l:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=weekly7['Week'], y=weekly7['orders'],
            marker_color=[C_RED if a else C_ACCENT for a in weekly7['is_anomaly']],
            name='הזמנות', opacity=0.85,
        ))
        for w in anomaly_weeks:
            if w in weekly7['Week'].values:
                fig.add_annotation(
                    x=w, y=weekly7[weekly7['Week']==w]['orders'].values[0],
                    text='⚠', showarrow=False, yshift=12, font_size=16,
                )
        fig.update_layout(
            title='נפח הזמנות שבועי (אדום = שבועות אנומליה)',
            title_font_color=C_BLUE, plot_bgcolor='white',
            xaxis_title='שבוע', yaxis_title='הזמנות', height=320,
        )
        fig.update_xaxes(tickangle=45)
        st.plotly_chart(fig, use_container_width=True)

    with col_r:
        fig2 = make_subplots(specs=[[{'secondary_y': True}]])
        fig2.add_trace(go.Scatter(
            x=weekly7['Week'], y=weekly7['delivery_rate'],
            mode='lines+markers', name='שיעור מסירה %',
            line=dict(color=C_GREEN, width=2),
        ), secondary_y=False)
        fig2.add_trace(go.Scatter(
            x=weekly7['Week'], y=weekly7['aov'],
            mode='lines+markers', name='AOV (₪)',
            line=dict(color=C_GOLD, width=2, dash='dot'),
        ), secondary_y=True)
        fig2.update_layout(
            title='שיעור מסירה ו-AOV שבועי',
            title_font_color=C_BLUE, plot_bgcolor='white', height=320,
            legend=dict(orientation='h', y=-0.3),
        )
        fig2.update_xaxes(tickangle=45)
        fig2.update_yaxes(title_text='מסירה %', secondary_y=False)
        fig2.update_yaxes(title_text='AOV ₪', secondary_y=True)
        st.plotly_chart(fig2, use_container_width=True)

    # ── גרף 2: דפוס יום בשבוע ──────────────────────────────────
    st.markdown('---')
    st.markdown('#### דפוס לפי יום בשבוע')

    days_heb = {0:'שני', 1:'שלישי', 2:'רביעי', 3:'חמישי', 4:'שישי', 5:'שבת', 6:'ראשון'}
    dow7 = fdf7.groupby('DayOfWeek').agg(
        orders=('Id','count'),
        delivered=('delivered','sum'),
        revenue=('Price','sum'),
    ).reset_index()
    dow7['delivery_rate'] = (dow7['delivered'] / dow7['orders'] * 100).round(1)
    dow7['aov']           = (dow7['revenue'] / dow7['orders']).round(2)
    n_weeks = fdf7['Week'].nunique()
    dow7['daily_avg']     = (dow7['orders'] / n_weeks).round(1)
    dow7['DayName']       = dow7['DayOfWeek'].map(days_heb)

    col_l2, col_r2 = st.columns(2)
    with col_l2:
        bar_colors = [C_RED if d == 5 else (C_GREEN if d == 6 else C_ACCENT)
                      for d in dow7['DayOfWeek']]
        fig3 = go.Figure(go.Bar(
            x=dow7['DayName'], y=dow7['daily_avg'],
            marker_color=bar_colors, text=dow7['daily_avg'],
            texttemplate='%{text:.0f}', textposition='outside',
        ))
        fig3.update_layout(
            title='ממוצע הזמנות יומי לפי יום בשבוע | אדום=שבת (שיא) | ירוק=ראשון',
            title_font_color=C_BLUE, plot_bgcolor='white',
            xaxis_title='יום', yaxis_title='ממוצע הזמנות', height=320,
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col_r2:
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(
            x=dow7['DayName'], y=dow7['delivery_rate'],
            marker_color=C_GREEN, name='שיעור מסירה %', opacity=0.8,
            text=dow7['delivery_rate'], texttemplate='%{text:.1f}%',
            textposition='outside',
        ))
        fig4.update_layout(
            title='שיעור מסירה לפי יום בשבוע | ראשון = הטוב ביותר (67.5%)',
            title_font_color=C_BLUE, plot_bgcolor='white',
            xaxis_title='יום', yaxis_title='% מסירה', height=320, yaxis_range=[50, 80],
        )
        st.plotly_chart(fig4, use_container_width=True)

    # ── גרף 3: שעת שיא לפי חודש ──────────────────────────────
    st.markdown('---')
    st.markdown('#### שעת שיא לפי חודש וחלוקה שעתית')

    col_l3, col_r3 = st.columns(2)
    with col_l3:
        hour7 = fdf7.groupby('Hour').size().reset_index(name='הזמנות')
        fig5 = px.bar(hour7, x='Hour', y='הזמנות',
                      color_discrete_sequence=[C_ACCENT],
                      title='פיזור הזמנות לפי שעה (כלל התקופה)',
                      text='הזמנות')
        fig5.update_traces(textposition='outside')
        fig5.update_layout(title_font_color=C_BLUE, plot_bgcolor='white',
                           xaxis_title='שעה', height=320)
        st.plotly_chart(fig5, use_container_width=True)

    with col_r3:
        st.markdown('**שעת שיא לפי חודש:**')
        for month in sorted(fdf7['Month'].unique()):
            mdf    = fdf7[fdf7['Month'] == month]
            peak_h = mdf.groupby('Hour').size().idxmax()
            peak_v = mdf.groupby('Hour').size().max()
            st.markdown(f'- **{month}**: שעה **{peak_h}:00** ({peak_v} הזמנות)')

        st.markdown('---')
        st.info('💡 **אנומליה**: שעת השיא הוזזה מ-20:00 (פברואר) ל-21:00 (מרץ ואילך) — ייתכן שקשורה לרמדאן ולפטאר (שבירת הצום)')

    # ── טבלת אנומליות ────────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### שבועות אנומליה מרכזיים')
    anomaly_data = {
        'שבוע': ['11–17 מרץ', '8–14 אפריל'],
        'הזמנות': [1134, 1573],
        'שינוי WoW': ['-40.1%', '+15.4%'],
        'מסירה %': ['72.2%', '48.4%'],
        'AOV ₪': ['106.73', '128.86'],
        'הסבר': ['תחילת רמדאן 2024 (11 מרץ)', 'תקיפת איראן? (13 אפריל 2024)'],
    }
    st.dataframe(pd.DataFrame(anomaly_data), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# TAB 8 — איכות נתונים (Q8)
# ══════════════════════════════════════════════════════════════════
with tab8:
    st.markdown('<div class="section-title">שאלה 8 — בדיקת איכות נתונים: payment_methods_user</div>',
                unsafe_allow_html=True)

    import os as _os
    _PATH = _os.path.join(_os.path.dirname(__file__), 'HAAT_DA_Dataset.xlsx')
    pm8 = pd.read_excel(_PATH, sheet_name='payment_methods_user')

    # ── מדדי סיכום ──────────────────────────────────────────────
    n_rows      = len(pm8)
    n_users_pm  = pm8['UserId'].nunique()
    n_dup_tok   = pm8['Token'].duplicated().sum()
    n_uniq_tok  = pm8['Token'].nunique()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric('שורות בטבלה',        f'{n_rows:,}')
    c2.metric('משתמשים ייחודיים',   f'{n_users_pm:,}')
    c3.metric('טוקנים כפולים',      f'{n_dup_tok:,}', delta='⚠ בעיה')
    c4.metric('טוקנים ייחודיים',    f'{n_uniq_tok:,}')

    st.markdown('---')

    # ── גרף 1: פיזור טוקנים ללקוח ──────────────────────────────
    col_l, col_r = st.columns(2)

    tpu8 = pm8.groupby('UserId')['Token'].count().reset_index(name='token_count')

    with col_l:
        bins_data = tpu8['token_count'].value_counts().sort_index().reset_index()
        bins_data.columns = ['מספר טוקנים', 'לקוחות']
        bins_data = bins_data[bins_data['מספר טוקנים'] <= 20]
        fig = px.bar(bins_data, x='מספר טוקנים', y='לקוחות',
                     color_discrete_sequence=[C_ACCENT],
                     title='פיזור טוקנים לכל לקוח (עד 20)',
                     text='לקוחות')
        fig.update_traces(textposition='outside')
        fig.update_layout(title_font_color=C_BLUE, plot_bgcolor='white', height=350)
        st.plotly_chart(fig, use_container_width=True)

    with col_r:
        buckets = [
            ('1 טוקן',    (tpu8['token_count']==1).sum()),
            ('2 טוקנים',  (tpu8['token_count']==2).sum()),
            ('3 טוקנים',  (tpu8['token_count']==3).sum()),
            ('4 טוקנים',  (tpu8['token_count']==4).sum()),
            ('5–9 טוקנים', ((tpu8['token_count']>=5) & (tpu8['token_count']<10)).sum()),
            ('10–49 טוקנים', ((tpu8['token_count']>=10) & (tpu8['token_count']<50)).sum()),
            ('50+ טוקנים', (tpu8['token_count']>=50).sum()),
        ]
        bdf = pd.DataFrame(buckets, columns=['קבוצה','לקוחות'])
        fig2 = px.pie(bdf, names='קבוצה', values='לקוחות', hole=0.35,
                      title='פיזור לקוחות לפי קבוצת טוקנים',
                      color_discrete_sequence=px.colors.sequential.Blues_r)
        fig2.update_traces(textposition='inside', textinfo='label+percent')
        fig2.update_layout(title_font_color=C_BLUE, height=350)
        st.plotly_chart(fig2, use_container_width=True)

    # ── גרף 2: TOP חשודים ──────────────────────────────────────
    st.markdown('---')
    st.markdown('#### TOP 10 לקוחות חשודים — מספר טוקנים גבוה מול מספר הזמנות נמוך')

    top10 = tpu8.nlargest(10, 'token_count').copy()
    top10['orders'] = top10['UserId'].apply(lambda uid: (fdf['UserId']==uid).sum())

    col_l2, col_r2 = st.columns([3, 2])
    with col_l2:
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            name='טוקנים', x=top10['UserId'].astype(str), y=top10['token_count'],
            marker_color=C_RED, text=top10['token_count'],
            textposition='outside',
        ))
        fig3.add_trace(go.Bar(
            name='הזמנות', x=top10['UserId'].astype(str), y=top10['orders'],
            marker_color=C_ACCENT, text=top10['orders'],
            textposition='outside',
        ))
        fig3.update_layout(
            title='TOP 10 לקוחות: טוקנים vs הזמנות',
            barmode='group', title_font_color=C_BLUE,
            plot_bgcolor='white', height=350,
            xaxis_title='UserId', yaxis_title='ספירה',
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col_r2:
        top10_disp = top10[['UserId','token_count','orders']].copy()
        top10_disp.columns = ['UserId','טוקנים','הזמנות']
        top10_disp['יחס'] = (top10_disp['טוקנים'] / (top10_disp['הזמנות'].clip(lower=1))).round(0).astype(int)
        st.dataframe(top10_disp, use_container_width=True, hide_index=True)
        st.error('⚠ UserId 11344: 161 טוקנים, 2 הזמנות בלבד — חשד ל-Card Testing!')

    # ── גרף 3: קרוס-רפרנס עם הזמנות ────────────────────────────
    st.markdown('---')
    st.markdown('#### קרוס-רפרנס: לקוחות בהזמנות מול payment_methods_user')

    users_pm_set  = set(pm8['UserId'])
    users_ord_set = set(fdf['UserId'].dropna())
    cash_users    = set(fdf[fdf['PaymentMethod']==0]['UserId'].dropna())
    credit_users  = set(fdf[fdf['PaymentMethod']==1]['UserId'].dropna())

    cross_data = {
        'קבוצה': [
            'בשתי הטבלאות',
            'רק ב-payment (לא הזמינו)',
            'רק בהזמנות (אין טוקן)',
            'מזומן עם טוקן',
            'אשראי ללא טוקן ⚠',
            'מזומן + אשראי (גמישים)',
        ],
        'מספר לקוחות': [
            len(users_pm_set & users_ord_set),
            len(users_pm_set - users_ord_set),
            len(users_ord_set - users_pm_set),
            len(cash_users & users_pm_set),
            len(credit_users - users_pm_set),
            len(cash_users & credit_users),
        ],
    }
    cdf = pd.DataFrame(cross_data)
    col_l3, col_r3 = st.columns([3, 2])
    with col_l3:
        fig4 = px.bar(cdf, x='מספר לקוחות', y='קבוצה', orientation='h',
                      color_discrete_sequence=[C_ACCENT],
                      title='פיזור לקוחות לפי קשר לטבלת payment',
                      text='מספר לקוחות')
        fig4.update_traces(textposition='outside')
        fig4.update_layout(title_font_color=C_BLUE, plot_bgcolor='white', height=350)
        st.plotly_chart(fig4, use_container_width=True)

    with col_r3:
        st.dataframe(cdf, use_container_width=True, hide_index=True)
        st.info(f'💡 {len(credit_users - users_pm_set)} לקוחות שילמו אשראי ואין להם טוקן — ייתכן תשלום דרך ספק חיצוני')

    # ── סיכום איכות נתונים ──────────────────────────────────────
    st.markdown('---')
    quality_data = {
        'בדיקה': [
            'ערכים חסרים', 'תקינות פורמט Token',
            'טוקנים כפולים', 'לקוחות עם 50+ טוקנים',
            'אשראי ללא טוקן',
        ],
        'תוצאה': [
            '✅ 0 חסרים',
            '✅ כל הטוקנים TOK_XXXXXXXX',
            f'⚠ {n_dup_tok:,} כפולים',
            f'⚠ {(tpu8["token_count"]>=50).sum()} לקוחות',
            f'⚠ {len(credit_users - users_pm_set)} לקוחות',
        ],
        'דחיפות': ['נמוכה', 'נמוכה', 'גבוהה', 'גבוהה מאוד', 'בינונית'],
    }
    st.dataframe(pd.DataFrame(quality_data), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════
# TAB 9 — פרזנטציה (Q9)
# ══════════════════════════════════════════════════════════════════
with tab9:
    st.markdown('<div class="section-title">שאלה 9 — הפרזנטציה: ממצאים מרכזיים להנהגה</div>',
                unsafe_allow_html=True)

    # ── שקופית 1: הממצא הכי חשוב ──────────────────────────────────
    st.markdown("""
    <div style="background:#1A376C; color:white; border-radius:12px; padding:24px 28px; margin-bottom:20px;">
      <div style="font-size:0.85rem; color:#ADD8E6; margin-bottom:8px;">שקופית 1 — הממצא החשוב ביותר</div>
      <div style="font-size:1.4rem; font-weight:bold; line-height:1.5;">
        1 מתוך 3 הזמנות לא מגיעה ללקוח —<br>
        וזה לא בעיה של ביקוש, זו בעיה של היצע שליחים
      </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric('שיעור מסירה כולל', '66.1%', delta='-33.9% לא נמסרו', delta_color='inverse')
    c2.metric('הזמנות שנאבדו', '9,716', delta='₪1.16M פוטנציאל')
    c3.metric('מסירה בשבת (שיא ביקוש)', '63.1%', delta='vs 67.5% בראשון', delta_color='inverse')

    st.markdown('---')

    # ── שקופית 2: 3 מספרים ─────────────────────────────────────────
    st.markdown('#### 📊 שלושה מספרים שמספרים את הסיפור')

    col_l, col_r = st.columns([2, 3])
    with col_l:
        numbers_data = {
            'מדד': ['שיעור מסירה כולל', 'מסירה בשבת', 'מסירה בשבוע 8–14 אפריל'],
            'ערך': ['66.1%', '63.1%', '48.4%'],
            'משמעות': [
                '9,716 הזמנות שלא הגיעו',
                'יום שיא בנפח = יום שפל במסירה',
                'שפל מוחלט — כמעט 1 מ-2 נפל',
            ],
        }
        st.dataframe(pd.DataFrame(numbers_data), use_container_width=True, hide_index=True)

    with col_r:
        # גרף: נפח vs מסירה לפי יום
        fdf9 = fdf.copy()
        fdf9['DayOfWeek'] = fdf9['OrderDate'].dt.dayofweek
        fdf9['Week'] = fdf9['OrderDate'].dt.to_period('W').astype(str)
        days_map = {0:'שני',1:'שלישי',2:'רביעי',3:'חמישי',4:'שישי',5:'שבת',6:'ראשון'}
        dow9 = fdf9.groupby('DayOfWeek').agg(
            orders=('Id','count'),
            delivered=('delivered','sum'),
        ).reset_index()
        dow9['delivery_rate'] = (dow9['delivered'] / dow9['orders'] * 100).round(1)
        n_weeks9 = fdf9['Week'].nunique()
        dow9['daily_avg'] = (dow9['orders'] / n_weeks9).round(1)
        dow9['DayName'] = dow9['DayOfWeek'].map(days_map)

        fig9 = make_subplots(specs=[[{'secondary_y': True}]])
        fig9.add_trace(go.Bar(
            x=dow9['DayName'], y=dow9['daily_avg'],
            name='ממוצע הזמנות/יום',
            marker_color=[C_RED if d==5 else C_ACCENT for d in dow9['DayOfWeek']],
            opacity=0.75,
        ), secondary_y=False)
        fig9.add_trace(go.Scatter(
            x=dow9['DayName'], y=dow9['delivery_rate'],
            name='מסירה %', mode='lines+markers',
            line=dict(color=C_GREEN, width=3),
            marker=dict(size=8),
        ), secondary_y=True)
        fig9.update_layout(
            title='ביקוש vs מסירה לפי יום | שבת: שיא ביקוש, שפל מסירה',
            title_font_color=C_BLUE, plot_bgcolor='white', height=300,
            legend=dict(orientation='h', y=-0.3),
        )
        fig9.update_yaxes(title_text='הזמנות/יום', secondary_y=False)
        fig9.update_yaxes(title_text='מסירה %', secondary_y=True, range=[55, 75])
        st.plotly_chart(fig9, use_container_width=True)

    st.markdown('---')

    # ── שקופית 3: ממצא מפתיע ──────────────────────────────────────
    st.markdown('#### 😮 ממצא שהפתיע אותי')
    st.markdown("""
    <div style="background:#F0F4FF; border-right:4px solid #2674C1; border-radius:8px; padding:16px 20px; margin-bottom:16px;">
      <b style="font-size:1.1rem;">57.6% מלקוחות המזומן רשמו כרטיס אשראי — ובחרו לשלם מזומן בכל זאת</b><br><br>
      9,190 לקוחות שמספיק סומכים על HAAT כדי לשמור כרטיס, אבל בוחרים מזומן.
      זה לא לקוחות שחוששים מהפלטפורמה — הם כבר מחוברים.
      (מעקב ומדידה קיימים גם כך — כל הזמנה משויכת ל-UserId בין אם מזומן ובין אם אשראי.)
      עם קמפיין ממוקד (קאשבק, הנחה בתשלום ראשון באשראי) ניתן להמיר אותם —
      מה שישפר גביה בטוחה ללא מזומן פיזי, יפחית ביטולים (מחויבות גבוהה יותר),
      ויאפשר תוכנית נאמנות אוטומטית עם זיכויים ישירים לכרטיס.
    </div>
    """, unsafe_allow_html=True)

    import os as _os9
    _PATH9 = _os9.path.join(_os9.path.dirname(__file__), 'HAAT_DA_Dataset.xlsx')
    pm9 = pd.read_excel(_PATH9, sheet_name='payment_methods_user')
    users_pm9   = set(pm9['UserId'])
    cash_u9     = set(fdf[fdf['PaymentMethod']==0]['UserId'].dropna())
    credit_u9   = set(fdf[fdf['PaymentMethod']==1]['UserId'].dropna())
    cash_w_tok  = len(cash_u9 & users_pm9)
    cash_pct9   = cash_w_tok / len(cash_u9) * 100 if cash_u9 else 0

    col_a, col_b, col_c = st.columns(3)
    col_a.metric('לקוחות מזומן עם כרטיס רשום', f'{cash_w_tok:,}', f'{cash_pct9:.1f}% מכלל משלמי מזומן')
    col_b.metric('לקוחות אשראי ללא טוקן', f'{len(credit_u9 - users_pm9)}', 'ייתכן ספק חיצוני')
    col_c.metric('לקוחות גמישים (שניהם)', f'{len(cash_u9 & credit_u9):,}', 'משלמים גם מזומן וגם אשראי')

    st.markdown('---')

    # ── שקופית 4: מה שלא ניתן להסביר ─────────────────────────────
    st.markdown('#### ❓ מה שלא הצלחתי להסביר לחלוטין')
    st.warning("""
    **שבוע 8–14 אפריל: מסירה קרסה ל-48.4% בזמן שהנפח עלה ב-15.4%**

    13 אפריל 2024 — ישראל חוותה תקיפת טילים ומל"טים מאיראן.
    אם שליחים בחרו לא לצאת, זה מסביר את הקריסה.

    **אבל ללא נתוני פעילות שליחים** (כמה היו online, מתי, מאיפה) — לא ניתן לאשר.
    הנתון מחייב חקירה: האם HAAT צריכה תוכנית חירום לאירועי ביטחון?
    """)

    st.markdown('---')

    # ── שקופית 5: המלצה קונקרטית ──────────────────────────────────
    st.markdown('#### 🎯 המלצה אחת שניתן ליישם השבוע הבא')
    st.success("""
    **גייסו 20–30 שליחים ייעודיים לשישי-שבת באזורים 1, 4, 10**

    📌 הצעו בונוס שיפט: +10% על כל מסירה בסוף שבוע

    💰 חישוב ROI:
    - שיפור של 5% במסירה בשבת = ~68 הזמנות נוספות
    - בAOV ₪122.76 = **₪8,348 הכנסה נוספת בשבת אחת**
    - עלות בונוס משוערת: ~₪2,000 → **ROI חיובי מהשבת הראשונה**
    """)

    st.markdown('---')

    # ── שקופית 6: השאלה שלא נשאלה ────────────────────────────────
    st.markdown('#### 💡 שאלה שהיה כדאי לשאול (ולא נשאלה)')
    st.info("""
    **מהו ה-LTV של לקוח חוזר לעומת לקוח חד-פעמי?**

    הדאטא מראה 42.3% לקוחות חוזרים — אבל לא ידוע:
    - האם הם מזמינים יותר? שווים יותר? מרוצים יותר?
    - מה גורם ללקוח לחזור אחרי הזמנה ראשונה?

    **למה זה קריטי:** עלות רכישת לקוח חדש גבוהה פי 5–7 מעלות שימור.
    אם מסירה כושלת = לקוח שלא חוזר, ה-ROI האמיתי של שיפור שיעור המסירה
    גבוה בהרבה ממה שמחשבים מהזמנה בודדת.

    **הנתון הנדרש:** Repeat Purchase Rate פר קוהורט, לפי אזור, עסק ושיטת תשלום.
    """)

    st.markdown('---')
    st.markdown("""
    <div style="background:#1B5E20; color:white; border-radius:10px; padding:16px 20px; text-align:center;">
      <b style="font-size:1.1rem;">
        HAAT יושבת על ביקוש אמיתי. האתגר: לא לאבד 1 מכל 3 לקוחות בדרך.<br>
        השליח הוא נקודת הכישלון — והוא גם ההזדמנות הכי גדולה.
      </b>
    </div>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# TAB 10 — בונוס: חישוב בונוסים לשליחים (Q10)
# ══════════════════════════════════════════════════════════════════
with tab10:
    st.markdown('<div class="section-title">שאלה 10 (בונוס) — חישוב בונוסים שבועיים לשליחים</div>',
                unsafe_allow_html=True)

    # ── חוקי הבונוס ────────────────────────────────────────────────
    st.markdown('#### חוקי הבונוס לפי אזור')
    rules_data = {
        'AreaId': [1, 4, 10],
        'תנאי ימים': ['סוף שבוע בלבד (שישי+שבת)', 'כל ימות השבוע', 'סוף שבוע בלבד (שישי+שבת)'],
        'חוק הבונוס': [
            'כל 50 מסירות → 100 ₪',
            'כל מסירה מעל ה-300 → 3 ₪/מסירה',
            'כל 50 מסירות → 150 ₪',
        ],
    }
    st.dataframe(pd.DataFrame(rules_data), use_container_width=True, hide_index=True)

    st.markdown('---')

    # ── הפונקציה ───────────────────────────────────────────────────
    with st.expander('📝 קוד הפונקציה calc_weekly_bonus()', expanded=False):
        st.code("""
def calc_weekly_bonus(orders_df, week_start, week_end):
    ws = pd.Timestamp(week_start)
    we = pd.Timestamp(week_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

    w = orders_df[
        (orders_df['OrderDate'] >= ws) & (orders_df['OrderDate'] <= we) &
        (orders_df['delivered'] == True) & (orders_df['DriverId'].notna())
    ].copy()

    results = []

    # AreaId 1 — סוף שבוע, 100 ₪ לכל 50 מסירות
    a1 = w[(w['AreaId'] == 1) & (w['IsWeekend'] == True)]
    g1 = a1.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g1['AreaId'] = 1
    g1['BonusAmount'] = (g1['TotalDelivered'] // 50) * 100
    results.append(g1)

    # AreaId 4 — כל הימים, 3 ₪ מעל מסירה 300
    a4 = w[w['AreaId'] == 4]
    g4 = a4.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g4['AreaId'] = 4
    g4['BonusAmount'] = g4['TotalDelivered'].apply(lambda n: max(0, n - 300) * 3)
    results.append(g4)

    # AreaId 10 — סוף שבוע, 150 ₪ לכל 50 מסירות
    a10 = w[(w['AreaId'] == 10) & (w['IsWeekend'] == True)]
    g10 = a10.groupby('DriverId').size().reset_index(name='TotalDelivered')
    g10['AreaId'] = 10
    g10['BonusAmount'] = (g10['TotalDelivered'] // 50) * 150
    results.append(g10)

    result = pd.concat(results, ignore_index=True)
    return result[result['BonusAmount'] > 0][
        ['DriverId', 'AreaId', 'TotalDelivered', 'BonusAmount']
    ].sort_values(['AreaId', 'BonusAmount'], ascending=[True, False]).reset_index(drop=True)
""", language='python')

    # ── בחירת שבוע ─────────────────────────────────────────────────
    st.markdown('#### הרץ את הפונקציה — בחר שבוע')

    fdf10 = fdf.copy()
    fdf10['DayOfWeek'] = fdf10['OrderDate'].dt.dayofweek
    fdf10['IsWeekend'] = fdf10['DayOfWeek'].isin([4, 5])
    fdf10['delivered'] = fdf10['ArriveDate'].notna()
    fdf10['Week']      = fdf10['OrderDate'].dt.to_period('W').astype(str)

    all_weeks10 = sorted(fdf10['Week'].dropna().unique())
    sel_week = st.selectbox('בחר שבוע', all_weeks10,
                            index=len(all_weeks10)//2 if all_weeks10 else 0)

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

    if sel_week:
        parts   = sel_week.split('/')
        ws_date = parts[0]
        we_date = parts[1]
        bonus_result = calc_weekly_bonus(fdf10, ws_date, we_date)

        c1, c2, c3 = st.columns(3)
        c1.metric('נהגים שזכו לבונוס', len(bonus_result))
        c2.metric('סה"כ בונוסים',
                  f'₪{bonus_result["BonusAmount"].sum():,.0f}' if len(bonus_result) else '₪0')
        c3.metric('שבוע שנבחר', sel_week)

        if len(bonus_result) == 0:
            st.warning(f'בשבוע {sel_week} — אף נהג לא עמד בסף הבונוס. ראה ניתוח מתחת.')
        else:
            st.success(f'נמצאו {len(bonus_result)} נהגים עם בונוס בשבוע זה!')
            st.dataframe(bonus_result, use_container_width=True, hide_index=True)

    st.markdown('---')

    # ── ניתוח קירבה לסף ────────────────────────────────────────────
    st.markdown('#### מדוע אף נהג לא מגיע לסף? — ניתוח קירבה לבונוס')

    fdf10['wk'] = fdf10['OrderDate'].dt.to_period('W').astype(str)

    # AreaId 1 — weekend deliveries per driver per week
    a1_all = fdf10[(fdf10['AreaId']==1) & fdf10['delivered'] & fdf10['IsWeekend'] & fdf10['DriverId'].notna()]
    a4_all = fdf10[(fdf10['AreaId']==4) & fdf10['delivered'] & fdf10['DriverId'].notna()]
    a10_all = fdf10[(fdf10['AreaId']==10) & fdf10['delivered'] & fdf10['IsWeekend'] & fdf10['DriverId'].notna()]

    col_l, col_r = st.columns(2)

    with col_l:
        threshold_data = {
            'אזור': ['AreaId 1', 'AreaId 4', 'AreaId 10'],
            'סף בונוס': [50, 300, 50],
            'מקסימום בפועל (שבוע בודד)': [
                a1_all.groupby(['wk','DriverId']).size().max() if len(a1_all) else 0,
                a4_all.groupby(['wk','DriverId']).size().max() if len(a4_all) else 0,
                a10_all.groupby(['wk','DriverId']).size().max() if len(a10_all) else 0,
            ],
        }
        tdf = pd.DataFrame(threshold_data)
        tdf['% מהסף'] = (tdf['מקסימום בפועל (שבוע בודד)'] / tdf['סף בונוס'] * 100).round(1).astype(str) + '%'
        st.dataframe(tdf, use_container_width=True, hide_index=True)

    with col_r:
        fig_thresh = go.Figure()
        fig_thresh.add_trace(go.Bar(
            name='סף בונוס',
            x=['AreaId 1', 'AreaId 4', 'AreaId 10'],
            y=[50, 300, 50],
            marker_color=C_RED, opacity=0.6,
        ))
        fig_thresh.add_trace(go.Bar(
            name='מקסימום בפועל',
            x=['AreaId 1', 'AreaId 4', 'AreaId 10'],
            y=[
                a1_all.groupby(['wk','DriverId']).size().max() if len(a1_all) else 0,
                a4_all.groupby(['wk','DriverId']).size().max() if len(a4_all) else 0,
                a10_all.groupby(['wk','DriverId']).size().max() if len(a10_all) else 0,
            ],
            marker_color=C_ACCENT,
        ))
        fig_thresh.update_layout(
            title='סף בונוס vs מקסימום בפועל (שבוע בודד)',
            barmode='group', title_font_color=C_BLUE,
            plot_bgcolor='white', height=320,
        )
        st.plotly_chart(fig_thresh, use_container_width=True)

    # ── TOP נהגים קרובים ביותר ─────────────────────────────────────
    st.markdown('---')
    st.markdown('#### TOP נהגים — הכי קרובים לבונוס (כל התקופה)')

    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.markdown('**AreaId 1 — סוף שבוע**')
        if len(a1_all):
            top_a1 = a1_all.groupby('DriverId').size().reset_index(name='מסירות_סה"כ')
            top_a1 = top_a1.nlargest(5, 'מסירות_סה"כ')
            top_a1['DriverId'] = top_a1['DriverId'].astype(int)
            top_a1['% מסף 50'] = (top_a1['מסירות_סה"כ'] / 50 * 100).round(0).astype(int).astype(str) + '%'
            st.dataframe(top_a1, use_container_width=True, hide_index=True)
        st.caption('סף: 50 מסירות/שבוע → 100 ₪')

    with col_b:
        st.markdown('**AreaId 4 — כל הימים**')
        if len(a4_all):
            top_a4 = a4_all.groupby('DriverId').size().reset_index(name='מסירות_סה"כ')
            top_a4 = top_a4.nlargest(5, 'מסירות_סה"כ')
            top_a4['DriverId'] = top_a4['DriverId'].astype(int)
            top_a4['% מסף 300'] = (top_a4['מסירות_סה"כ'] / 300 * 100).round(0).astype(int).astype(str) + '%'
            st.dataframe(top_a4, use_container_width=True, hide_index=True)
        st.caption('סף: 300 מסירות/שבוע → 3 ₪/מסירה')

    with col_c:
        st.markdown('**AreaId 10 — סוף שבוע**')
        if len(a10_all):
            top_a10 = a10_all.groupby('DriverId').size().reset_index(name='מסירות_סה"כ')
            top_a10 = top_a10.nlargest(5, 'מסירות_סה"כ')
            top_a10['DriverId'] = top_a10['DriverId'].astype(int)
            top_a10['% מסף 50'] = (top_a10['מסירות_סה"כ'] / 50 * 100).round(0).astype(int).astype(str) + '%'
            st.dataframe(top_a10, use_container_width=True, hide_index=True)
        st.caption('סף: 50 מסירות/שבוע → 150 ₪')

    st.error('⚠ אף נהג לא הגיע לסף הבונוס בשבוע בודד. הספים גבוהים פי 3–15 מהמקסימום בפועל.')
    st.info('💡 המלצה: לכייל את הספים ל-10–15 מסירות לאזורי 1/10, ו-15 מסירות לאזור 4 — תכנית הבונוסים תהפוך אפקטיבית.')
