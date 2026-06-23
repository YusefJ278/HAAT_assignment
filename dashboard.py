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
    df['is_delivery']      = df['WithDriver'] == True
    df['is_pickup']        = df['WithDriver'] == False
    # הזמנת איסוף "הושלמה" אם לא נדחתה ע"י מסעדה והוכנה (PickedDate לרוב לא נרשם — פגם בנתונים)
    df['order_completed']  = (
        (df['is_delivery'] & df['delivered']) |
        (df['is_pickup'] & df['RejectedDateByRestaurant'].isna() & df['StartPrepare'].notna())
    )
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
_order_type_opts = {'כל הזמנות': None, '🚗 משלוח (עם שליח)': True, '🏃 איסוף עצמי (Pickup)': False}
sel_order_type = st.sidebar.radio('סוג הזמנה', list(_order_type_opts.keys()), index=0)

st.sidebar.markdown('---')
st.sidebar.caption('HAAT Delivery · ניתוח Q1 2024')

# ── פילטור ───────────────────────────────────────────────────────
_type_val = _order_type_opts[sel_order_type]
_base = df if _type_val is None else df[df['WithDriver'] == _type_val]
fdf = _base[_base['AreaName'].isin(sel_areas) & _base['Month'].isin(sel_months)]

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
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11 = st.tabs([
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
    '🏙️ מיקוד: אום אל-פחם',
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

    # ── תגלית: פיצול Delivery / Pickup ───────────────────────────
    n_del = fdf['is_delivery'].sum()
    n_pck = fdf['is_pickup'].sum()
    n_rej = fdf['rejected'].sum()
    pct_del = n_del / total * 100 if total else 0
    pct_pck = n_pck / total * 100 if total else 0
    rej_rate = n_rej / total * 100 if total else 0

    st.markdown(f"""
    <div style="background:#FFF3CD; border-right:5px solid #E65C00; border-radius:8px;
                padding:14px 20px; margin-bottom:12px;">
      <b style="font-size:1.05rem; color:#5C3D00;">⚠ תגלית מרכזית: הדאטאסט מכיל שני סוגי הזמנות</b><br>
      <div style="display:flex; gap:32px; margin-top:8px; flex-wrap:wrap;">
        <span>🚗 <b>משלוח (WithDriver=True):</b> {n_del:,} הזמנות ({pct_del:.1f}%) — <b style="color:#1B5E20;">100% הצלחה</b></span>
        <span>🏃 <b>איסוף עצמי (Pickup):</b> {n_pck:,} הזמנות ({pct_pck:.1f}%) — PickedDate לא נרשם (פגם בנתונים)</span>
        <span>❌ <b>דחיות מסעדה:</b> {n_rej:,} ({rej_rate:.1f}%) — הכשל היחיד האמיתי במערכת</span>
      </div>
      <div style="font-size:0.83rem; color:#6B4900; margin-top:6px;">
        שיעור המסירה הישן (66.1%) חושב על כל הדאטאסט ביחד — לא תקין. השתמש בפילטר "סוג הזמנה" בסיידבר.
      </div>
    </div>
    """, unsafe_allow_html=True)

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

    fdf_d        = fdf[fdf['is_delivery']]   # הזמנות משלוח בלבד
    fdf_p        = fdf[fdf['is_pickup']]     # הזמנות איסוף בלבד
    del_rate     = fdf_d['delivered'].mean() * 100 if len(fdf_d) else 0
    assign_med   = fdf_d[fdf_d['has_driver']]['time_to_assign'].median() if len(fdf_d) else 0
    aov          = fdf['Price'].mean()
    reject_rate  = fdf['rejected'].mean() * 100
    pickup_pct   = fdf['is_pickup'].mean() * 100 if len(fdf) else 0

    kpis = [
        {
            'name':    'שיעור הצלחת משלוח (WithDriver=True)',
            'value':   f'{del_rate:.1f}%',
            'current': del_rate,
            'good':    90, 'warn': 80,
            'desc':    'הזמנות עם שליח שהגיעו ללקוח (ArriveDate) — 100% הצלחה ✅ | הנתון הישן 66.1% כלל הזמנות Pickup בטעות',
            'formula': 'COUNT(ArriveDate IS NOT NULL WHERE WithDriver=True) ÷ COUNT(WithDriver=True) × 100',
            'team':    'תפעול',
        },
        {
            'name':    'זמן שיוך שליח (חציון)',
            'value':   f'{assign_med:.1f} דק׳',
            'current': assign_med,
            'good':    5, 'warn': 15,
            'desc':    'דקות מהזמנה עד שיוך שליח (הזמנות משלוח בלבד) | תקין <5 דק׳ | מדאיג >15 דק׳',
            'formula': 'MEDIAN(DriverCandidateAssignedDate − OrderDate) WHERE WithDriver=True',
            'team':    'תפעול',
            'inverse': True,
        },
        {
            'name':    'ממוצע ערך הזמנה (AOV)',
            'value':   f'₪{aov:.2f}',
            'current': aov,
            'good':    120, 'warn': 80,
            'desc':    'הכנסה ממוצעת לכל הזמנה (כלל ההזמנות) | תקין >₪120 | מדאיג <₪80',
            'formula': 'SUM(Price) ÷ COUNT(*)',
            'team':    'שיווק',
        },
        {
            'name':    'שיעור דחייה מסעדה — הכשל האמיתי היחיד',
            'value':   f'{reject_rate:.1f}%',
            'current': reject_rate,
            'good':    2, 'warn': 5,
            'desc':    'אחוז הזמנות (כולל Pickup) שנדחו ע"י מסעדה | תקין <2% | מדאיג >5% | זהו הכשל האופרטיבי האמיתי',
            'formula': 'COUNT(RejectedDate IS NOT NULL) ÷ COUNT(*) × 100',
            'team':    'תפעול / שותפויות',
            'inverse': True,
        },
        {
            'name':    'שיעור הזמנות איסוף עצמי (Pickup)',
            'value':   f'{pickup_pct:.1f}%',
            'current': pickup_pct,
            'good':    0, 'warn': 50,
            'desc':    'אחוז הזמנות WithDriver=False — לקוח אוסף בעצמו | PickedDate לא נרשם לרוב (פגם בנתונים)',
            'formula': 'COUNT(WithDriver=False) ÷ COUNT(*) × 100',
            'team':    'מוצר / נתונים',
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
    vals  = [del_rate, 100-assign_med*3, aov/1.5, 100-reject_rate*5, 100-pickup_pct]
    names = ['הצלחת משלוח','מהירות שיוך','AOV','אמינות מסעדה','% הזמנות משלוח']
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

    # הזמנות משלוח בלבד — שיעור מסירה תקף רק עליהן
    fdf3_d = fdf[fdf['is_delivery']]
    if len(fdf3_d) == 0:
        st.warning('אין הזמנות משלוח בפילטר הנוכחי. בחר "כל הזמנות" או "משלוח" בסיידבר.')
    else:
        st.info(f'⚠ שיעור מסירה מחושב על **הזמנות משלוח בלבד** ({len(fdf3_d):,} מתוך {len(fdf):,}) | '
                f'הזמנות Pickup ({len(fdf)-len(fdf3_d):,}) לא נכללות — ArriveDate אינה רלוונטית להן.')

    col_l, col_r = st.columns(2)

    # ── דחיות מסעדה לפי אזור (המדד המשמעותי) ────────────────────
    with col_l:
        area_rej = fdf.groupby('AreaName').agg(
            orders=('Id','count'), rejected=('rejected','sum')
        ).assign(rej_rate=lambda x: (x['rejected']/x['orders']*100).round(2)).reset_index()\
         .sort_values('rej_rate', ascending=True)
        area_rej['color'] = area_rej['rej_rate'].apply(
            lambda r: C_GREEN if r < 2 else (C_ACCENT if r < 5 else C_RED)
        )
        fig = go.Figure(go.Bar(
            x=area_rej['rej_rate'], y=area_rej['AreaName'],
            orientation='h', marker_color=area_rej['color'],
            text=area_rej['rej_rate'].apply(lambda r: f'{r:.2f}%'),
            textposition='outside',
        ))
        avg_rej = fdf['rejected'].mean() * 100
        fig.add_vline(x=avg_rej, line_dash='dash', line_color=C_GOLD,
                      annotation_text=f'ממוצע: {avg_rej:.2f}%')
        fig.add_vline(x=2, line_dash='dot', line_color=C_GREEN,
                      annotation_text='סף תקין 2%', annotation_position='top right')
        fig.update_layout(title='שיעור דחיית מסעדה לפי אזור — הכשל האמיתי',
                          title_font_color=C_BLUE,
                          xaxis_title='שיעור דחייה (%)',
                          plot_bgcolor='white', showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    # ── דחיות מסעדה לפי חודש ─────────────────────────────────────
    with col_r:
        monthly_rej = fdf.groupby('Month').agg(
            orders=('Id','count'), rejected=('rejected','sum')
        ).assign(rej_rate=lambda x: (x['rejected']/x['orders']*100).round(2)).reset_index()
        fig = go.Figure()
        fig.add_bar(x=monthly_rej['Month'], y=monthly_rej['orders'],
                    name='נפח הזמנות', marker_color=C_ACCENT, opacity=0.6,
                    yaxis='y2')
        fig.add_scatter(x=monthly_rej['Month'], y=monthly_rej['rej_rate'],
                        name='דחיות מסעדה %', mode='lines+markers',
                        line=dict(color=C_RED, width=3),
                        marker=dict(size=9), yaxis='y')
        fig.update_layout(
            title='נפח הזמנות ודחיות מסעדה לפי חודש', title_font_color=C_BLUE,
            yaxis=dict(title='שיעור דחייה (%)'),
            yaxis2=dict(title='נפח הזמנות', overlaying='y', side='right'),
            plot_bgcolor='white', legend=dict(orientation='h', y=-0.2),
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── דחיות מסעדה לפי שעה ──────────────────────────────────────
    st.markdown('---')
    hourly = fdf.groupby('Hour').agg(
        orders=('Id','count'), rejected=('rejected','sum')
    ).assign(rej_rate=lambda x: (x['rejected']/x['orders']*100).round(2)).reset_index()

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=hourly['Hour'], y=hourly['orders'],
        name='נפח הזמנות', marker_color=C_ACCENT, opacity=0.55,
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=hourly['Hour'], y=hourly['rej_rate'],
        name='דחיות מסעדה %', mode='lines+markers',
        line=dict(color=C_RED, width=2.5), marker=dict(size=7),
    ), secondary_y=True)
    fig.update_layout(
        title='נפח הזמנות ודחיות מסעדה לפי שעה ביום',
        title_font_color=C_BLUE, plot_bgcolor='white',
        xaxis=dict(title='שעה', tickmode='linear', dtick=1),
        legend=dict(orientation='h', y=-0.2),
    )
    fig.update_yaxes(title_text='נפח הזמנות', secondary_y=False)
    fig.update_yaxes(title_text='דחיות מסעדה (%)', secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)

    # ── טבלת אזורים ──────────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### טבלת פירוט לפי אזור (משלוח בלבד + דחיות כולל)')
    # מסירה — על הזמנות משלוח בלבד; דחיות — על כלל ההזמנות
    area_del_tbl = fdf3_d.groupby('AreaName').agg(
        הזמנות_משלוח=('Id','count'),
        נמסרו=('delivered','sum'),
    ).assign(**{'שיעור מסירה %': lambda x: (x['נמסרו']/x['הזמנות_משלוח']*100).round(1)})
    area_rej_tbl = fdf.groupby('AreaName').agg(
        הזמנות_כולל=('Id','count'),
        נדחו=('rejected','sum'),
    ).assign(**{'שיעור דחייה %': lambda x: (x['נדחו']/x['הזמנות_כולל']*100).round(1)})
    area_tbl = area_del_tbl.join(area_rej_tbl, how='outer').reset_index()
    area_tbl = area_tbl.sort_values('שיעור מסירה %', ascending=False)
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
        rejected=('rejected','sum'),
    ).assign(
        pct=lambda x: (x['revenue']/total_rev*100).round(1),
        rej_rate=lambda x: (x['rejected']/x['orders']*100).round(2),
        aov=lambda x: (x['revenue']/x['orders']).round(2),
    ).sort_values('revenue', ascending=False).reset_index()

    # ── מדדים עיקריים ────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    top_area     = area_rev.iloc[0]
    top3_biz_pct = 7.2
    n_biz        = fdf['BusinessId'].nunique()
    corr_val     = round(float(np.corrcoef(area_rev['rej_rate'], area_rev['revenue'])[0,1]), 3)
    c1.metric('סה"כ הכנסות',           f'₪{total_rev:,.0f}')
    c2.metric('האזור המוביל',           f'{top_area["AreaName"]} ({top_area["pct"]:.1f}%)')
    c3.metric('תלות TOP 3 עסקים',       f'{top3_biz_pct}%')
    c4.metric('קורלציה הכנסה↔דחייה',   f'r = {corr_val:.3f}')

    st.markdown('---')

    # ── פאנל כפול: הכנסות + דחיית מסעדה לפי אזור ───────────────
    area_sorted = area_rev.sort_values('revenue', ascending=True)
    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=['הכנסות לפי אזור (אלפי ₪)',
                                        'שיעור דחיית מסעדה לפי אזור (%)'],
                        horizontal_spacing=0.08)

    rev_colors = [C_RED if r >= 6 else (C_GOLD if r >= 3 else C_GREEN)
                  for r in area_sorted['rej_rate']]
    fig.add_trace(go.Bar(
        x=area_sorted['revenue']/1000, y=area_sorted['AreaName'],
        orientation='h', marker_color=rev_colors,
        text=area_sorted.apply(lambda r: f'₪{r["revenue"]/1000:.0f}K ({r["pct"]:.1f}%)', axis=1),
        textposition='outside', name='הכנסות',
    ), row=1, col=1)

    rej_colors = [C_RED if r >= 6 else (C_GOLD if r >= 3 else C_GREEN)
                  for r in area_sorted['rej_rate']]
    avg_rej = fdf['rejected'].mean() * 100
    fig.add_trace(go.Bar(
        x=area_sorted['rej_rate'], y=area_sorted['AreaName'],
        orientation='h', marker_color=rej_colors,
        text=area_sorted['rej_rate'].apply(lambda r: f'{r:.2f}%'),
        textposition='outside', name='דחייה %',
    ), row=1, col=2)
    fig.add_vline(x=avg_rej, line_dash='dash', line_color=C_GOLD,
                  annotation_text=f'ממוצע {avg_rej:.2f}%', row=1, col=2)

    fig.update_layout(height=420, plot_bgcolor='white', showlegend=False,
                      title_text='הכנסות מול שיעור דחיית מסעדה — לפי אזור',
                      title_font_color=C_BLUE, title_font_size=14)
    fig.update_xaxes(range=[0, 1550], row=1, col=1)
    fig.update_xaxes(range=[0, 10],   row=1, col=2)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('---')
    col_l, col_r = st.columns(2)

    # ── scatter: הכנסה vs דחיית מסעדה ───────────────────────────
    with col_l:
        fig = px.scatter(
            area_rev, x='rej_rate', y='revenue',
            size='orders', color='rej_rate',
            color_continuous_scale=['#1B5E20','#2674C1','#8B0000'],
            text='AreaName', hover_data=['orders','aov'],
            title=f'קורלציה: הכנסה ↔ דחיית מסעדה (r = {corr_val:.2f})',
            labels={'rej_rate':'שיעור דחייה (%)','revenue':'הכנסות (₪)'},
        )
        fig.update_traces(textposition='top center', textfont_size=9)
        z = np.polyfit(area_rev['rej_rate'], area_rev['revenue'], 1)
        xr = np.linspace(area_rev['rej_rate'].min(), area_rev['rej_rate'].max(), 100)
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
    tbl = area_rev[['AreaName','revenue','orders','aov','pct','rej_rate']].copy()
    tbl.columns = ['אזור','הכנסות (₪)','הזמנות','AOV (₪)','% מסה"כ','דחיית מסעדה %']
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

    fdf7   = fdf.copy()
    fdf7_d = fdf[fdf['is_delivery']].copy()   # לחישוב שיעור מסירה בלבד
    fdf7['DayOfWeek'] = fdf7['OrderDate'].dt.dayofweek
    fdf7['Week']      = fdf7['OrderDate'].dt.to_period('W').astype(str)
    fdf7['Hour']      = fdf7['OrderDate'].dt.hour
    fdf7['Date']      = fdf7['OrderDate'].dt.date
    fdf7_d['DayOfWeek'] = fdf7_d['OrderDate'].dt.dayofweek
    fdf7_d['Week']      = fdf7_d['OrderDate'].dt.to_period('W').astype(str)

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
    _w7_vol = fdf7.groupby('Week').agg(orders=('Id','count'), revenue=('Price','sum')).reset_index()
    _w7_del = fdf7_d.groupby('Week').agg(delivered=('delivered','sum'), orders_d=('Id','count')).reset_index()
    _w7_rej = fdf7.groupby('Week').agg(rejected=('rejected','sum')).reset_index()
    weekly7 = _w7_vol.merge(_w7_del, on='Week', how='left').merge(_w7_rej, on='Week', how='left')
    weekly7['delivery_rate'] = (weekly7['delivered'] / weekly7['orders_d'] * 100).round(1)
    weekly7['rej_rate']      = (weekly7['rejected'] / weekly7['orders'] * 100).round(2)
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
            x=weekly7['Week'], y=weekly7['rej_rate'],
            mode='lines+markers', name='שיעור דחייה %',
            line=dict(color=C_RED, width=2),
        ), secondary_y=False)
        fig2.add_trace(go.Scatter(
            x=weekly7['Week'], y=weekly7['aov'],
            mode='lines+markers', name='AOV (₪)',
            line=dict(color=C_GOLD, width=2, dash='dot'),
        ), secondary_y=True)
        fig2.update_layout(
            title='שיעור דחיית מסעדה ו-AOV שבועי',
            title_font_color=C_BLUE, plot_bgcolor='white', height=320,
            legend=dict(orientation='h', y=-0.3),
        )
        fig2.update_xaxes(tickangle=45)
        fig2.update_yaxes(title_text='דחייה %', secondary_y=False)
        fig2.update_yaxes(title_text='AOV ₪', secondary_y=True)
        st.plotly_chart(fig2, use_container_width=True)

    # ── גרף 2: דפוס יום בשבוע ──────────────────────────────────
    st.markdown('---')
    st.markdown('#### דפוס לפי יום בשבוע')

    days_heb = {0:'שני', 1:'שלישי', 2:'רביעי', 3:'חמישי', 4:'שישי', 5:'שבת', 6:'ראשון'}
    _dow7_vol = fdf7.groupby('DayOfWeek').agg(orders=('Id','count'), revenue=('Price','sum')).reset_index()
    _dow7_del = fdf7_d.groupby('DayOfWeek').agg(delivered=('delivered','sum'), orders_d=('Id','count')).reset_index()
    dow7 = _dow7_vol.merge(_dow7_del, on='DayOfWeek', how='left')
    dow7['delivery_rate'] = (dow7['delivered'] / dow7['orders_d'] * 100).round(1)
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
        # שיעור מסירה על הזמנות משלוח = 100% בכל הימים; מציג דחיות מסעדה לפי יום (מעניין יותר)
        _dow7_rej = fdf7.groupby('DayOfWeek').agg(
            orders_all=('Id','count'), rejected=('rejected','sum')
        ).reset_index()
        _dow7_rej['rej_rate'] = (_dow7_rej['rejected'] / _dow7_rej['orders_all'] * 100).round(2)
        _dow7_rej['DayName'] = _dow7_rej['DayOfWeek'].map(days_heb)
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(
            x=_dow7_rej['DayName'], y=_dow7_rej['rej_rate'],
            marker_color=C_RED, name='שיעור דחייה מסעדה %', opacity=0.85,
            text=_dow7_rej['rej_rate'], texttemplate='%{text:.2f}%',
            textposition='outside',
        ))
        fig4.update_layout(
            title='שיעור דחיית מסעדה לפי יום | (שיעור משלוח = 100% בכל הימים)',
            title_font_color=C_BLUE, plot_bgcolor='white',
            xaxis_title='יום', yaxis_title='% דחיות', height=320,
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
    _aw_idx = weekly7.set_index('Week')
    _aw_rej = {w: (f"{_aw_idx.loc[w, 'rej_rate']:.2f}%" if w in _aw_idx.index else 'N/A')
               for w in anomaly_weeks}
    anomaly_data = {
        'שבוע': ['11–17 מרץ', '8–14 אפריל'],
        'הזמנות': [1134, 1573],
        'שינוי WoW': ['-40.1%', '+15.4%'],
        'דחיית מסעדה %': [_aw_rej[anomaly_weeks[0]], _aw_rej[anomaly_weeks[1]]],
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
    st.markdown('#### TOP 10 לקוחות לפי מספר טוקנים — אנומליה לבדיקה')

    st.warning(
        '⚠ **מגבלת הניתוח:** טבלת payment_methods_user אינה כוללת תאריכים — '
        'היא snapshot היסטורי של כל הכרטיסים שנרשמו אי פעם. '
        'הדאטאסט מכסה 4 חודשים בלבד (פבר–מאי 2024). '
        'לכן אי אפשר להסיק מסקנות חד-משמעיות על הונאה — '
        'ייתכן שלמשתמשים האלה יש מאות הזמנות מחוץ לחלון הזמן.'
    )

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
            name='הזמנות (4 חודשים)', x=top10['UserId'].astype(str), y=top10['orders'],
            marker_color=C_ACCENT, text=top10['orders'],
            textposition='outside',
        ))
        fig3.update_layout(
            title='TOP 10 לקוחות: טוקנים (כל הזמנים) vs הזמנות (פבר–מאי 2024 בלבד)',
            barmode='group', title_font_color=C_BLUE,
            plot_bgcolor='white', height=350,
            xaxis_title='UserId', yaxis_title='ספירה',
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col_r2:
        top10_disp = top10[['UserId','token_count','orders']].copy()
        top10_disp.columns = ['UserId','טוקנים','הזמנות (4 חו\')']
        top10_disp['יחס'] = (top10_disp['טוקנים'] / (top10_disp['הזמנות (4 חו\')'].clip(lower=1))).round(0).astype(int)
        st.dataframe(top10_disp, use_container_width=True, hide_index=True)
        st.info('לאימות חשד — נדרש: היסטוריית הזמנות מלאה + timestamp רישום הטוקן + לוגי IP')

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
    st.markdown('#### סיכום ממצאי איכות נתונים')
    quality_data = {
        'בדיקה': [
            'ערכים חסרים',
            'תקינות פורמט Token',
            'טוקנים כפולים',
            'לקוחות עם 50+ טוקנים (אנומליה)',
            'אשראי ללא טוקן',
            'תאריך רישום טוקן',
        ],
        'תוצאה': [
            '✅ 0 חסרים',
            '✅ כל הטוקנים TOK_XXXXXXXX',
            f'⚠ {n_dup_tok:,} כפולים — באג ברישום',
            f'⚠ {(tpu8["token_count"]>=50).sum()} לקוחות — דורש בדיקה',
            f'⚠ {len(credit_users - users_pm_set)} לקוחות — ספק חיצוני?',
            '❌ חסר — לא ניתן לנתח התנהגות לאורך זמן',
        ],
        'דחיפות': ['נמוכה', 'נמוכה', 'גבוהה', 'בינונית — ראו הסתייגות', 'בינונית', 'גבוהה — מגבלת ניתוח'],
    }
    st.dataframe(pd.DataFrame(quality_data), use_container_width=True, hide_index=True)
    st.info(
        '💡 **צעד הבא קונקרטי:** להוסיף עמודת `CreatedAt` לטבלת payment_methods_user. '
        'זה יאפשר לקבוע אם טוקנים רובים נרשמו בפרק זמן קצר (דגל הונאה אמיתי) '
        'לעומת הצטברות לגיטימית לאורך שנים. '
        'במקביל — לשלוף היסטוריית הזמנות מלאה (מחוץ לחלון 4 החודשים) '
        'עבור 8 הלקוחות עם 50+ טוקנים.'
    )

# ══════════════════════════════════════════════════════════════════
# TAB 9 — פרזנטציה (Q9)
# ══════════════════════════════════════════════════════════════════
with tab9:
    st.markdown('<div class="section-title">שאלה 9 — הפרזנטציה: ממצאים מרכזיים להנהגה</div>',
                unsafe_allow_html=True)

    # ── שקופית 1: הממצא הכי חשוב — מתוקן ─────────────────────────
    fdf9_d = fdf[fdf['is_delivery']]
    fdf9_p = fdf[fdf['is_pickup']]
    _del_rate9  = fdf9_d['delivered'].mean() * 100 if len(fdf9_d) else 0
    _rej_rate9  = fdf['rejected'].mean() * 100
    _pck_pct9   = fdf['is_pickup'].mean() * 100
    _completed9 = fdf['order_completed'].mean() * 100

    st.markdown(f"""
    <div style="background:#1A376C; color:white; border-radius:12px; padding:24px 28px; margin-bottom:20px;">
      <div style="font-size:0.85rem; color:#ADD8E6; margin-bottom:8px;">שקופית 1 — הממצא החשוב ביותר (מתוקן)</div>
      <div style="font-size:1.4rem; font-weight:bold; line-height:1.6;">
        HAAT מצליחה 100% בהזמנות משלוח —<br>
        הכשל היחיד האמיתי הוא דחיות מסעדה (3.9%)<br>
        <span style="font-size:0.95rem; color:#ADD8E6; font-weight:normal;">
          הנתון הישן 66.1% היה שגוי: 34.8% מהדאטאסט הן הזמנות Pickup שלעולם לא היה להן שליח
        </span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric('שיעור הצלחת משלוח (WithDriver=True)', f'{_del_rate9:.1f}%',
              delta='✅ כל הזמנות עם שליח הגיעו')
    c2.metric('שיעור דחיית מסעדה — הכשל האמיתי', f'{_rej_rate9:.1f}%',
              delta='❌ מעל סף תקין 2%', delta_color='inverse')
    c3.metric('הזמנות Pickup (34.8% מהדאטא)', f'{int(_pck_pct9*len(fdf)/100):,}',
              delta='PickedDate לא נרשם — פגם בנתונים')

    st.markdown('---')

    # ── שקופית 2: 3 מספרים מתוקנים ────────────────────────────────
    st.markdown('#### 📊 שלושה מספרים שמספרים את הסיפור (מתוקן)')

    col_l, col_r = st.columns([2, 3])
    with col_l:
        numbers_data = {
            'מדד': ['הצלחת משלוח (WithDriver=True)', 'שיעור דחיית מסעדה', 'Pickup — PickedDate חסר'],
            'ערך': [f'{_del_rate9:.1f}%', f'{_rej_rate9:.1f}%', f'{_pck_pct9:.1f}%'],
            'משמעות': [
                f'כל {len(fdf9_d):,} הזמנות המשלוח הושלמו',
                'הכשל האמיתי — מסעדות דוחות הזמנות',
                '34.8% מהדאטא — PickedDate לא מתועד (פגם)',
            ],
        }
        st.dataframe(pd.DataFrame(numbers_data), use_container_width=True, hide_index=True)

    with col_r:
        # גרף: נפח (כל הזמנות) vs דחיות מסעדה לפי יום
        fdf9 = fdf.copy()
        fdf9['DayOfWeek'] = fdf9['OrderDate'].dt.dayofweek
        fdf9['Week'] = fdf9['OrderDate'].dt.to_period('W').astype(str)
        days_map = {0:'שני',1:'שלישי',2:'רביעי',3:'חמישי',4:'שישי',5:'שבת',6:'ראשון'}
        dow9 = fdf9.groupby('DayOfWeek').agg(
            orders=('Id','count'),
            rejected=('rejected','sum'),
        ).reset_index()
        dow9['rej_rate'] = (dow9['rejected'] / dow9['orders'] * 100).round(2)
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
            x=dow9['DayName'], y=dow9['rej_rate'],
            name='דחיות מסעדה %', mode='lines+markers',
            line=dict(color=C_RED, width=3),
            marker=dict(size=8),
        ), secondary_y=True)
        fig9.update_layout(
            title='ביקוש vs דחיות מסעדה לפי יום | (משלוח=100% הצלחה בכל יום)',
            title_font_color=C_BLUE, plot_bgcolor='white', height=300,
            legend=dict(orientation='h', y=-0.3),
        )
        fig9.update_yaxes(title_text='הזמנות/יום', secondary_y=False)
        fig9.update_yaxes(title_text='דחיות %', secondary_y=True)
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
    **שבוע 8–14 אפריל: דחיות מסעדה עלו בצורה חריגה בזמן שהנפח עלה ב-15.4%**

    13 אפריל 2024 — ישראל חוותה תקיפת טילים ומל"טים מאיראן.
    הדחיות מסביר: מסעדות סגרו מוקדם / לא הצליחו לעמוד בביקוש.
    שיעור ה-Pickup Completion אף הוא ירד באותה תקופה (ניתן לאמת עם PickedDate מלא).

    **אבל ללא נתוני פעילות שליחים ומסעדות** — לא ניתן לאשר את הסיבה הספציפית.
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
        HAAT מצליחה 100% בהזמנות משלוח — זה חוזק מרשים.<br>
        האתגר האמיתי: להוריד את דחיות המסעדה (3.9%) ולתקן מעקב Pickup.<br>
        <span style="font-weight:normal; font-size:0.95rem;">
          אם נתקן PickedDate tracking → נוכל לראות את שיעור ההשלמה האמיתי של 34.8% מהלקוחות
        </span>
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

# ══════════════════════════════════════════════════════════════════
# TAB 11 — מיקוד: אום אל-פחם
# ══════════════════════════════════════════════════════════════════
with tab11:
    st.markdown('<div class="section-title">ניתוח מורחב — Umm al-Fahem (אום אל-פחם) · פבר׳–מאי 2024</div>',
                unsafe_allow_html=True)

    # ── נתוני הבסיס ──────────────────────────────────────────────
    uaf = df[df['AreaName'] == 'Umm al-Fahem'].copy()
    uaf['Month']     = uaf['OrderDate'].dt.to_period('M').astype(str)
    uaf['Week']      = uaf['OrderDate'].dt.to_period('W').astype(str)
    uaf['DayOfWeek'] = uaf['OrderDate'].dt.dayofweek
    uaf['IsWeekend'] = uaf['DayOfWeek'].isin([4, 5])
    uaf['delivered'] = uaf['ArriveDate'].notna()
    uaf_d = uaf[uaf['is_delivery']]   # הזמנות משלוח בלבד לחישוב שיעור מסירה

    MONTHS_ORDER = ['2024-02', '2024-03', '2024-04', '2024-05']
    MONTH_HEB    = {'2024-02': 'פבר׳', '2024-03': 'מרץ', '2024-04': 'אפריל', '2024-05': 'מאי'}
    DOW_HEB      = {0: 'שני', 1: 'שלישי', 2: 'רביעי', 3: 'חמישי', 4: 'שישי', 5: 'שבת', 6: 'ראשון'}

    total_uaf   = len(uaf)
    total_all11 = len(df)
    unique_users_uaf = uaf['UserId'].nunique()

    # ── KPI cards ────────────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    _uaf_del_rate = uaf_d['delivered'].mean()*100 if len(uaf_d) else 0
    _uaf_rej_rate = uaf['rejected'].mean()*100
    c1.metric('סה"כ הזמנות',        f'{total_uaf:,}')
    c2.metric('נתח מכלל הדאטא',     f'{total_uaf/total_all11*100:.1f}%')
    c3.metric('משתמשים ייחודיים',   f'{unique_users_uaf:,}')
    c4.metric('שיעור מסירה (WithDriver=True)', f'{_uaf_del_rate:.1f}%',
              delta=f'⚠ ישן: 58.6% (כלל Pickup)')
    c5.metric('AOV',                f'₪{uaf["Price"].mean():.1f}')

    st.markdown('---')

    # ── 1: הזמנות ומשתמשים לפי חודש ─────────────────────────────
    monthly11 = uaf.groupby('Month').agg(
        orders=('Id', 'count'),
        users=('UserId', 'nunique'),
    ).reset_index()
    _m11_rej = uaf.groupby('Month').agg(rej_n=('rejected','sum'), orders_m=('Id','count')).reset_index()
    _m11_rej['rej_rate_m'] = (_m11_rej['rej_n'] / _m11_rej['orders_m'] * 100).round(2)
    monthly11 = monthly11.merge(_m11_rej[['Month','rej_rate_m']], on='Month', how='left')
    monthly11 = monthly11[monthly11['Month'].isin(MONTHS_ORDER)].copy()
    monthly11['month_label'] = monthly11['Month'].map(MONTH_HEB)

    col_l, col_r = st.columns(2)

    with col_l:
        fig = make_subplots(specs=[[{'secondary_y': True}]])
        fig.add_trace(go.Bar(
            x=monthly11['month_label'], y=monthly11['orders'],
            name='הזמנות', marker_color=C_ACCENT, opacity=0.85,
            text=monthly11['orders'], textposition='outside',
        ), secondary_y=False)
        fig.add_trace(go.Bar(
            x=monthly11['month_label'], y=monthly11['users'],
            name='משתמשים', marker_color=C_GREEN, opacity=0.7,
            text=monthly11['users'], textposition='outside',
        ), secondary_y=False)
        fig.add_trace(go.Scatter(
            x=monthly11['month_label'], y=monthly11['rej_rate_m'],
            name='דחייה%', mode='lines+markers',
            line=dict(color=C_RED, width=3), marker=dict(size=9),
        ), secondary_y=True)
        fig.update_layout(
            title='הזמנות, משתמשים ושיעור דחיית מסעדה לפי חודש',
            title_font_color=C_BLUE, plot_bgcolor='white', barmode='group',
            legend=dict(orientation='h', y=-0.25), height=380,
        )
        fig.update_yaxes(title_text='כמות', secondary_y=False)
        fig.update_yaxes(title_text='דחייה%', secondary_y=True, range=[0, 8])
        st.plotly_chart(fig, use_container_width=True)

    # ── 2: חדשים vs. חוזרים ─────────────────────────────────────
    with col_r:
        seen11 = set()
        new_list11, ret_list11, mlabels11 = [], [], []
        for m in MONTHS_ORDER:
            mu  = set(uaf[uaf['Month'] == m]['UserId'])
            new_u = mu - seen11
            ret_u = mu & seen11
            seen11 |= mu
            new_list11.append(len(new_u))
            ret_list11.append(len(ret_u))
            mlabels11.append(MONTH_HEB[m])

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=mlabels11, y=new_list11,
            name='חדשים', marker_color=C_ACCENT,
            text=new_list11, textposition='inside',
        ))
        fig2.add_trace(go.Bar(
            x=mlabels11, y=ret_list11,
            name='חוזרים', marker_color=C_GREEN,
            text=ret_list11, textposition='inside',
        ))
        ret_pcts = [f'{r/(n+r)*100:.1f}%' if (n+r) > 0 else '0%'
                    for n, r in zip(new_list11, ret_list11)]
        for i, (label, pct) in enumerate(zip(mlabels11, ret_pcts)):
            fig2.add_annotation(
                x=label, y=new_list11[i]+ret_list11[i]+40,
                text=f'{pct} חוזרים', showarrow=False,
                font=dict(color=C_GREEN, size=11),
            )
        fig2.update_layout(
            title='משתמשים חדשים vs. חוזרים לפי חודש',
            title_font_color=C_BLUE, plot_bgcolor='white',
            barmode='stack', legend=dict(orientation='h', y=-0.25),
            height=380, yaxis_range=[0, 2700],
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('---')

    # ── 3: Retention חודשי ───────────────────────────────────────
    col_l2, col_r2 = st.columns(2)

    with col_l2:
        periods11 = MONTHS_ORDER
        ret_rates11, ret_pair_labels11 = [], []
        for i in range(len(periods11) - 1):
            u1 = set(uaf[uaf['Month'] == periods11[i]  ]['UserId'])
            u2 = set(uaf[uaf['Month'] == periods11[i+1]]['UserId'])
            r  = len(u1 & u2) / len(u1) * 100 if u1 else 0
            ret_rates11.append(round(r, 1))
            ret_pair_labels11.append(f"{MONTH_HEB[periods11[i]]} ← {MONTH_HEB[periods11[i+1]]}")

        avg_ret11 = sum(ret_rates11) / len(ret_rates11)
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            x=ret_pair_labels11, y=ret_rates11,
            marker_color=[C_ACCENT, C_ACCENT, C_GREEN],
            text=[f'{v:.1f}%' for v in ret_rates11],
            textposition='outside', width=0.45,
        ))
        fig3.add_hline(y=avg_ret11, line_dash='dash', line_color=C_GOLD,
                       annotation_text=f'ממוצע {avg_ret11:.1f}%',
                       annotation_position='bottom right')
        fig3.update_layout(
            title='Retention חודשי — % שחזרו לחודש הבא (טרנד עולה ✓)',
            title_font_color=C_BLUE, plot_bgcolor='white',
            yaxis_title='Retention %', yaxis_range=[0, 22],
            showlegend=False, height=340,
        )
        st.plotly_chart(fig3, use_container_width=True)

    # ── 4: תדירות הזמנה למשתמש ──────────────────────────────────
    with col_r2:
        user_freq11 = uaf.groupby('UserId')['Id'].count()
        freq_labels11 = ['1 הזמנה', '2 הזמנות', '3–4', '5–9', '10+']
        freq_counts11 = [
            (user_freq11 == 1).sum(),
            (user_freq11 == 2).sum(),
            ((user_freq11 >= 3) & (user_freq11 <= 4)).sum(),
            ((user_freq11 >= 5) & (user_freq11 <= 9)).sum(),
            (user_freq11 >= 10).sum(),
        ]
        total_u11 = sum(freq_counts11)
        fig4 = go.Figure(go.Bar(
            x=freq_labels11, y=freq_counts11,
            marker_color=[C_RED, C_ACCENT, C_GOLD, C_GREEN, C_BLUE],
            text=[f'{v:,}<br>({v/total_u11*100:.1f}%)' for v in freq_counts11],
            textposition='outside',
        ))
        fig4.update_layout(
            title=f'התפלגות תדירות הזמנה למשתמש | סה"כ {total_u11:,} משתמשים',
            title_font_color=C_BLUE, plot_bgcolor='white',
            yaxis_title='מספר משתמשים', showlegend=False,
            yaxis_range=[0, max(freq_counts11) * 1.25], height=340,
        )
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown('---')

    # ── 5: פיזור יומי ────────────────────────────────────────────
    dow11 = uaf.groupby('DayOfWeek').agg(
        orders=('Id', 'count'),
        users=('UserId', 'nunique'),
        rej_n=('rejected', 'sum'),
    ).reset_index()
    dow11['rej_rate'] = (dow11['rej_n'] / dow11['orders'] * 100).round(2)
    dow11['day_name'] = dow11['DayOfWeek'].map(DOW_HEB)
    day_order = ['שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת', 'ראשון']
    dow11['day_name'] = pd.Categorical(dow11['day_name'], categories=day_order, ordered=True)
    dow11 = dow11.sort_values('day_name')

    col_l3, col_r3 = st.columns(2)

    with col_l3:
        day_colors = [C_ACCENT if d not in [4, 5] else C_RED for d in dow11['DayOfWeek']]
        fig5 = make_subplots(specs=[[{'secondary_y': True}]])
        fig5.add_trace(go.Bar(
            x=dow11['day_name'], y=dow11['orders'],
            name='הזמנות', marker_color=day_colors, opacity=0.85,
            text=dow11['orders'], textposition='outside',
        ), secondary_y=False)
        fig5.add_trace(go.Scatter(
            x=dow11['day_name'], y=dow11['rej_rate'],
            name='דחייה%', mode='lines+markers',
            line=dict(color=C_RED, width=2.5), marker=dict(size=8),
        ), secondary_y=True)
        fig5.update_layout(
            title='הזמנות ודחיית מסעדה לפי יום בשבוע | אדום = סוף שבוע',
            title_font_color=C_BLUE, plot_bgcolor='white',
            legend=dict(orientation='h', y=-0.25), height=360,
        )
        fig5.update_yaxes(title_text='הזמנות', secondary_y=False)
        fig5.update_yaxes(title_text='דחייה%', secondary_y=True, range=[0, 8])
        st.plotly_chart(fig5, use_container_width=True)

    # ── 6: אירוע אפריל 8–14 ─────────────────────────────────────
    with col_r3:
        ev_windows = {
            'לפני (1–7 אפריל)':    uaf[(uaf['OrderDate'] >= '2024-04-01') & (uaf['OrderDate'] <= '2024-04-07')],
            'האירוע (8–14 אפריל)': uaf[(uaf['OrderDate'] >= '2024-04-08') & (uaf['OrderDate'] <= '2024-04-14')],
            'אחרי (15–21 אפריל)':  uaf[(uaf['OrderDate'] >= '2024-04-15') & (uaf['OrderDate'] <= '2024-04-21')],
        }
        ev_labels11  = list(ev_windows.keys())
        ev_orders11  = [len(v) for v in ev_windows.values()]
        ev_rej11     = [round(v['rejected'].mean() * 100, 2) for v in ev_windows.values()]
        ev_colors11  = [C_ACCENT, C_RED, C_GREEN]

        fig6 = make_subplots(specs=[[{'secondary_y': True}]])
        fig6.add_trace(go.Bar(
            x=ev_labels11, y=ev_orders11,
            name='הזמנות', marker_color=ev_colors11,
            text=ev_orders11, textposition='outside',
        ), secondary_y=False)
        fig6.add_trace(go.Scatter(
            x=ev_labels11, y=ev_rej11,
            name='דחייה%', mode='lines+markers',
            line=dict(color=C_GOLD, width=3), marker=dict(size=10),
        ), secondary_y=True)
        for xi, yi in zip(ev_labels11, ev_rej11):
            fig6.add_annotation(x=xi, y=yi, text=f'{yi:.1f}%',
                                yshift=18, showarrow=False,
                                font=dict(color=C_GOLD, size=12))
        fig6.update_layout(
            title='אירוע 8–14 אפריל | נפח הזמנות ושיעור דחיית מסעדה',
            title_font_color=C_BLUE, plot_bgcolor='white',
            legend=dict(orientation='h', y=-0.25),
            height=360, yaxis_range=[300, 560],
        )
        fig6.update_yaxes(title_text='הזמנות', secondary_y=False)
        fig6.update_yaxes(title_text='דחייה%', secondary_y=True, range=[0, 10])
        st.plotly_chart(fig6, use_container_width=True)

    st.markdown('---')

    # ── 7: טרנד שבועי ────────────────────────────────────────────
    _w11_vol = uaf.groupby('Week').agg(orders=('Id','count'), users=('UserId','nunique'), rej_n=('rejected','sum')).reset_index()
    weekly11 = _w11_vol.sort_values('Week')
    weekly11 = weekly11[weekly11['Week'] >= '2024-02'].copy()
    weekly11['rej_rate'] = (weekly11['rej_n'] / weekly11['orders'] * 100).round(2)

    fig7 = make_subplots(specs=[[{'secondary_y': True}]])
    fig7.add_trace(go.Scatter(
        x=weekly11['Week'], y=weekly11['orders'],
        name='הזמנות', mode='lines+markers',
        line=dict(color=C_ACCENT, width=2.5), marker=dict(size=6),
        fill='tozeroy', fillcolor='rgba(38,116,193,0.12)',
    ), secondary_y=False)
    fig7.add_trace(go.Scatter(
        x=weekly11['Week'], y=weekly11['users'],
        name='משתמשים', mode='lines+markers',
        line=dict(color=C_GREEN, width=2, dash='dot'), marker=dict(size=5),
    ), secondary_y=False)
    fig7.add_trace(go.Scatter(
        x=weekly11['Week'], y=weekly11['rej_rate'],
        name='דחייה%', mode='lines+markers',
        line=dict(color=C_GOLD, width=2), marker=dict(size=5),
    ), secondary_y=True)

    event_week11 = weekly11[weekly11['Week'].str.startswith('2024-04-08')]
    if len(event_week11):
        _ew = event_week11['Week'].values[0]
        fig7.add_shape(type='line', x0=_ew, x1=_ew, y0=0, y1=1,
                       xref='x', yref='paper',
                       line=dict(color=C_RED, dash='dash', width=1.5))
        fig7.add_annotation(x=_ew, y=1.02, xref='x', yref='paper',
                            text='8–14 אפריל', showarrow=False,
                            font=dict(color=C_RED, size=10), xanchor='right')

    fig7.update_layout(
        title='טרנד שבועי — הזמנות, משתמשים ושיעור דחיית מסעדה',
        title_font_color=C_BLUE, plot_bgcolor='white', height=380,
        legend=dict(orientation='h', y=-0.2),
        xaxis=dict(tickangle=35),
    )
    fig7.update_yaxes(title_text='הזמנות / משתמשים', secondary_y=False)
    fig7.update_yaxes(title_text='דחייה%', secondary_y=True, range=[0, 10])
    st.plotly_chart(fig7, use_container_width=True)

    st.markdown('---')

    # ── 8: השוואת ערים ───────────────────────────────────────────
    _bc11_vol = df.groupby('AreaName').agg(
        orders=('Id', 'count'), users=('UserId', 'nunique'), aov=('Price', 'mean')
    ).reset_index()
    _bc11_rej = df.groupby('AreaName').agg(rej_n=('rejected','sum'), orders_r=('Id','count')).reset_index()
    _bc11_rej['rej_rate'] = (_bc11_rej['rej_n'] / _bc11_rej['orders_r'] * 100).round(2)
    by_city11 = _bc11_vol.merge(_bc11_rej[['AreaName','rej_rate']], on='AreaName', how='left')
    by_city11 = by_city11.sort_values('orders', ascending=True)
    by_city11['is_uaf'] = by_city11['AreaName'] == 'Umm al-Fahem'

    col_c1, col_c2 = st.columns(2)

    with col_c1:
        ord_colors = [C_RED if u else C_ACCENT for u in by_city11['is_uaf']]
        fig8 = go.Figure(go.Bar(
            x=by_city11['orders'], y=by_city11['AreaName'],
            orientation='h', marker_color=ord_colors,
            text=by_city11['orders'], textposition='outside',
        ))
        fig8.update_layout(
            title='הזמנות לפי עיר | אום אל-פחם = פי 2 מהשנייה',
            title_font_color=C_BLUE, plot_bgcolor='white',
            xaxis_title='הזמנות', xaxis_range=[0, 11000], height=380,
        )
        st.plotly_chart(fig8, use_container_width=True)

    with col_c2:
        rej_colors11 = [C_RED if u else C_GREEN for u in by_city11['is_uaf']]
        avg_rej11    = df['rejected'].mean() * 100
        fig9 = go.Figure(go.Bar(
            x=by_city11['rej_rate'], y=by_city11['AreaName'],
            orientation='h', marker_color=rej_colors11,
            text=by_city11['rej_rate'].apply(lambda v: f'{v:.2f}%'),
            textposition='outside',
        ))
        fig9.add_vline(x=avg_rej11, line_dash='dash', line_color=C_GOLD,
                       annotation_text=f'ממוצע {avg_rej11:.2f}%')
        fig9.update_layout(
            title='שיעור דחיית מסעדה לפי עיר | אום אל-פחם בצבע אדום',
            title_font_color=C_BLUE, plot_bgcolor='white',
            xaxis_title='שיעור דחייה%', xaxis_range=[0, 10], height=380,
        )
        st.plotly_chart(fig9, use_container_width=True)

    # ── תובנות מרכזיות ───────────────────────────────────────────
    st.markdown('---')
    st.markdown('#### 💡 תובנות מרכזיות — אום אל-פחם')
    power_users11 = (user_freq11 >= 5).sum()
    col_i1, col_i2 = st.columns(2)
    with col_i1:
        st.info(f"""
**74.8% הזמינו פעם אחת בלבד** — הזדמנות שימור עצומה.
עם Retention שעולה מ-13.6% לכ-15.4%, המגמה חיובית אבל רחוקה מהפוטנציאל.
Power Users (≥5 הזמנות): {power_users11} בלבד ({power_users11/unique_users_uaf*100:.1f}%).
        """)
    with col_i2:
        st.warning(f"""
**שיעור מסירה (WithDriver=True): {_uaf_del_rate:.1f}%** — (הנתון הישן 58.6% כלל Pickup).
שיעור דחיית מסעדה: {_uaf_rej_rate:.2f}% — הכשל האמיתי.
שבת = שיא ביקוש (1,581 הזמנות) — נדרש ניתוח נפרד לפי WithDriver.
אם תתקן PickedDate tracking ב-Pickup — אום אל-פחם (30.7% מהדאטא) תיתן תמונה מלאה.
        """)
