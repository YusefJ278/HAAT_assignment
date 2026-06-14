import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

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
    body, .stApp { direction: rtl; }
    h1, h2, h3, h4, p, div, span, label { direction: rtl; text-align: right; }
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
    [data-testid="stSidebar"] { direction: rtl; }
    [data-testid="stSidebar"] * { direction: rtl; text-align: right; }
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
tab1, tab2, tab3, tab4 = st.tabs([
    '📊 מבט כללי (Q1)',
    '🎯 מדדי KPI (Q2)',
    '🚚 אמינות משלוחים (Q3)',
    '👤 יעילות שליחים (Q4)',
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
        st.plotly_chart(fig, width='stretch')

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
        st.plotly_chart(fig, width='stretch')

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
        st.plotly_chart(fig, width='stretch')

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
        st.plotly_chart(fig, width='stretch')

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
            'team':    'פיננסים',
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
    st.plotly_chart(fig, width='stretch')

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
        st.plotly_chart(fig, width='stretch')

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
        st.plotly_chart(fig, width='stretch')

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
    st.plotly_chart(fig, width='stretch')

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
    st.dataframe(area_tbl, width='stretch', hide_index=True)

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
        first=('OrderDate','min'),
        last=('OrderDate','max'),
    ).reset_index()
    drv['ימים פעילים'] = (drv['last'] - drv['first']).dt.days + 1
    drv['שיעור הצלחה %'] = (drv['נמסרו'] / drv['שובצו'] * 100).round(1)
    drv['יעילות (משלוחים/יום)'] = (drv['נמסרו'] / drv['ימים פעילים']).round(3)
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
        st.plotly_chart(fig, width='stretch')

    # ── יעילות מנורמלת ───────────────────────────────────────────
    with col_r:
        legit   = drv[(drv['ימים פעילים'] >= 7) & (drv['נמסרו'] >= 10)]
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
        st.plotly_chart(fig, width='stretch')

    # ── התפלגות יעילות ───────────────────────────────────────────
    st.markdown('---')
    Q1v   = drv['יעילות (משלוחים/יום)'].quantile(0.25)
    Q3v   = drv['יעילות (משלוחים/יום)'].quantile(0.75)
    fence = Q3v + 3 * (Q3v - Q1v)
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
    st.plotly_chart(fig, width='stretch')

    # ── טבלה אינטראקטיבית ────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### חיפוש שליח')
    search = st.text_input('הכנס מזהה שליח (DriverId)', placeholder='לדוגמה: 538')
    show_df = drv_sorted[['DriverId','שובצו','נמסרו','שיעור הצלחה %',
                           'ימים פעילים','יעילות (משלוחים/יום)']].copy()
    if search:
        try:
            show_df = show_df[show_df['DriverId'] == float(search)]
        except:
            pass
    st.dataframe(show_df.head(50), width='stretch', hide_index=True)
