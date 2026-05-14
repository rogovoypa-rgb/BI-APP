import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# ФУНКЦИИ ДЛЯ ФОРМАТИРОВАНИЯ ЧИСЕЛ
# ==========================================
def format_number(value):
    if pd.isna(value):
        return "0"
    try:
        return f"{int(value):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(value)

def format_float(value, decimals=1):
    if pd.isna(value):
        return "0"
    try:
        rounded = round(value, decimals)
        if decimals == 0:
            formatted = str(int(rounded))
        else:
            integer_part = int(rounded)
            fractional_part = abs(int(round((rounded - integer_part) * 10**decimals)))
            formatted = f"{integer_part},{fractional_part:0{decimals}d}"
        return formatted
    except (ValueError, TypeError):
        return str(value)

# ==========================================
# 1. НАСТРОЙКА СТРАНИЦЫ
# ==========================================
st.set_page_config(page_title="BI Портал продаж", layout="wide")
st.title("📊 BI Портал аналитики продаж")

# ==========================================
# 2. ЗАГРУЗКА ДАННЫХ
# ==========================================
@st.cache_data
def load_data():
    df = pd.read_excel('sales_data.xlsx')
    
    df['Дата'] = pd.to_datetime(df['Документ.Дата'], dayfirst=True, errors='coerce')
    df['Месяц'] = df['Дата'].dt.strftime('%Y-%m')
    df['Год'] = df['Дата'].dt.year
    
    df['Валовая_прибыль'] = df['Сумма без НДС'] - df['Себестоимость']
    df['Рентабельность_%'] = (df['Валовая_прибыль'] / df['Сумма без НДС'] * 100).fillna(0)
    
    df = df.rename(columns={
        'Сумма': 'Выручка',
        'Контрагент': 'Контрагент',
        'Номенклатура': 'Номенклатура'
    })
    
    df['Период.Месяц'] = pd.to_numeric(df['Период.Месяц'], errors='coerce')
    df = df.dropna(subset=['Период.Месяц'])
    df['Период.Месяц'] = df['Период.Месяц'].astype(int)
    
    return df

df = load_data()

# ==========================================
# 3. ПОЛУЧАЕМ ДОСТУПНЫЕ ГОДЫ
# ==========================================
available_years = sorted(df['Год'].dropna().unique())
if len(available_years) == 0:
    available_years = [2024]

# ==========================================
# 4. НАЗВАНИЯ МЕСЯЦЕВ
# ==========================================
month_names = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
    5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
    9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}

# ==========================================
# 5. БОКОВАЯ ПАНЕЛЬ С ФИЛЬТРАМИ
# ==========================================
st.sidebar.header("🔍 Фильтры")

selected_year = st.sidebar.selectbox("📅 Выберите год", available_years)
df_year = df[df['Год'] == selected_year]

available_months_num = sorted(df_year['Период.Месяц'].unique())
available_months_display = [month_names[m] for m in available_months_num]

selected_month_display = st.sidebar.selectbox("Выберите месяц", available_months_display)
selected_month_num = available_months_num[available_months_display.index(selected_month_display)]

all_customers = sorted(df_year['Контрагент'].dropna().unique())
selected_customers = st.sidebar.multiselect(
    "Выберите контрагентов",
    all_customers,
    default=all_customers[:5] if len(all_customers) > 5 else all_customers
)

df_filtered = df_year[(df_year['Период.Месяц'] == selected_month_num) & (df_year['Контрагент'].isin(selected_customers))]

# ==========================================
# 6. ГОДОВЫЕ МЕТРИКИ
# ==========================================
st.divider()
st.subheader(f"📈 ИТОГИ ЗА {selected_year} ГОД")

year_revenue = df_year['Выручка'].sum()
year_profit = df_year['Валовая_прибыль'].sum()
year_margin = (year_profit / year_revenue * 100) if year_revenue > 0 else 0
year_quantity = df_year['Количество'].sum()

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("💰 Годовая выручка", f"{format_number(year_revenue)} ₽")
with col2:
    st.metric("📈 Годовая прибыль", f"{format_number(year_profit)} ₽")
with col3:
    st.metric("🎯 Рентабельность", f"{format_float(year_margin, 1)}%")
with col4:
    st.metric("📦 Продано (шт)", f"{format_number(year_quantity)}")

if len(available_years) > 1 and selected_year > min(available_years):
    prev_year = selected_year - 1
    if prev_year in available_years:
        df_prev = df[df['Год'] == prev_year]
        prev_revenue = df_prev['Выручка'].sum()
        if prev_revenue > 0:
            change = ((year_revenue - prev_revenue) / prev_revenue * 100)
            st.caption(f"📊 Изменение выручки vs {prev_year}: {format_float(change, 1)}%")

st.divider()

# ==========================================
# 7. ПОМЕСЯЧНАЯ РАЗБИВКА
# ==========================================
st.subheader(f"📅 ПОМЕСЯЧНАЯ РАЗБИВКА ЗА {selected_year} ГОД")

monthly = df_year.groupby('Период.Месяц').agg({
    'Выручка': 'sum',
    'Валовая_прибыль': 'sum',
    'Количество': 'sum'
}).reset_index()
monthly['Название'] = monthly['Период.Месяц'].map(month_names)
monthly['Рентабельность'] = (monthly['Валовая_прибыль'] / monthly['Выручка'] * 100).fillna(0)
monthly = monthly.sort_values('Период.Месяц')

for _, row in monthly.iterrows():
    c = st.columns([1.5, 1, 1, 1, 1])
    with c[0]:
        st.markdown(f"**{row['Название']}**")
    with c[1]:
        st.metric("Выручка", f"{format_number(row['Выручка'])} ₽", label_visibility="collapsed")
    with c[2]:
        st.metric("Прибыль", f"{format_number(row['Валовая_прибыль'])} ₽", label_visibility="collapsed")
    with c[3]:
        st.metric("Рентабельность", f"{format_float(row['Рентабельность'], 1)}%", label_visibility="collapsed")
    with c[4]:
        st.metric("Кол-во", f"{format_number(row['Количество'])}", label_visibility="collapsed")

st.divider()

# ==========================================
# 8. АНАЛИЗ ТОП-5 КОНТРАГЕНТОВ
# ==========================================
st.subheader(f"🏆 ТОП-5 КОНТРАГЕНТОВ ЗА {selected_year}")

cust_rev = df_year.groupby('Контрагент')['Выручка'].sum().reset_index()
cust_rev = cust_rev.sort_values('Выручка', ascending=False)
top5 = cust_rev.head(5)['Контрагент'].tolist()

monthly_cust = df_year.groupby(['Контрагент', 'Период.Месяц'])['Выручка'].sum().reset_index()

table_data = []
for c in top5:
    row = {'Контрагент': c}
    row['Год'] = cust_rev[cust_rev['Контрагент'] == c]['Выручка'].values[0]
    for m in available_months_num:
        val = monthly_cust[(monthly_cust['Контрагент'] == c) & (monthly_cust['Период.Месяц'] == m)]['Выручка'].sum()
        row[month_names[m]] = val
    table_data.append(row)

other_rev = cust_rev[~cust_rev['Контрагент'].isin(top5)]['Выручка'].sum()
other_row = {'Контрагент': '📦 ОСТАЛЬНЫЕ'}
other_row['Год'] = other_rev
for m in available_months_num:
    val = monthly_cust[(~monthly_cust['Контрагент'].isin(top5)) & (monthly_cust['Период.Месяц'] == m)]['Выручка'].sum()
    other_row[month_names[m]] = val
table_data.append(other_row)

df_top5 = pd.DataFrame(table_data)

def fmt(x):
    return f"{int(x):,}".replace(",", " ") if x > 0 else "0"

html = '<table style="width:100%; border-collapse:collapse">'
html += '<tr style="background:#2E86AB; color:white">'
html += '<th style="padding:8px">Контрагент</th><th>💰 За год</th>'
for m in available_months_num:
    html += f'<th style="padding:8px">{month_names[m][:3]}</th>'
html += '</tr>'

for _, row in df_top5.iterrows():
    html += '<tr>'
    html += f'<td style="padding:6px; font-weight:bold">{row["Контрагент"]}</td>'
    html += f'<td style="padding:6px; font-weight:bold">{fmt(row["Год"])} ₽</td>'
    for m in available_months_num:
        val = row[month_names[m]]
        html += f'<td style="padding:6px; font-size:12px">{fmt(val)} ₽</td>'
    html += '</tr>'
html += '</table>'

st.markdown(html, unsafe_allow_html=True)

total_top5 = cust_rev[cust_rev['Контрагент'].isin(top5)]['Выручка'].sum()
st.caption(f"📊 Топ-5: {format_number(total_top5)} ₽ ({format_float(total_top5/year_revenue*100,1)}% от общей выручки)")
st.divider()

# ==========================================
# 9. ДЕТАЛИ ЗА ВЫБРАННЫЙ МЕСЯЦ
# ==========================================
st.subheader(f"📊 ДЕТАЛИ ЗА {selected_month_display} {selected_year}")

m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("💰 Выручка", f"{format_number(df_filtered['Выручка'].sum())} ₽")
with m2:
    st.metric("📈 Прибыль", f"{format_number(df_filtered['Валовая_прибыль'].sum())} ₽")
with m3:
    marg = df_filtered['Валовая_прибыль'].sum() / df_filtered['Выручка'].sum() * 100 if df_filtered['Выручка'].sum() > 0 else 0
    st.metric("🎯 Рентабельность", f"{format_float(marg, 1)}%")
with m4:
    st.metric("📦 Продано (шт)", f"{format_number(df_filtered['Количество'].sum())}")

# ==========================================
# 10. ГРАФИКИ
# ==========================================
c1, c2 = st.columns(2)
with c1:
    top10 = df_filtered.groupby('Контрагент')['Выручка'].sum().nlargest(10).reset_index()
    if not top10.empty:
        fig = px.bar(top10, x='Выручка', y='Контрагент', orientation='h', title='Топ-10 контрагентов')
        st.plotly_chart(fig, use_container_width=True)

with c2:
    comp = df_filtered.groupby('Контрагент')[['Выручка', 'Валовая_прибыль']].sum().nlargest(10, 'Выручка').reset_index()
    if not comp.empty:
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(name='Выручка', x=comp['Контрагент'], y=comp['Выручка'], marker_color='#2E86AB'))
        fig2.add_trace(go.Bar(name='Прибыль', x=comp['Контрагент'], y=comp['Валовая_прибыль'], marker_color='#52B788'))
        fig2.update_layout(title='Выручка vs Прибыль', barmode='group')
        st.plotly_chart(fig2, use_container_width=True)

# ==========================================
# 11. ТАБЛИЦА ДАННЫХ
# ==========================================
st.subheader("📋 Детальные данные")
show_cols = ['Дата', 'Контрагент', 'Номенклатура', 'Выручка', 'Валовая_прибыль', 'Рентабельность_%', 'Количество']
show_cols = [c for c in show_cols if c in df_filtered.columns]
if not df_filtered.empty:
    st.dataframe(df_filtered[show_cols].head(100), use_container_width=True)
    csv = df_filtered[show_cols].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
    st.download_button("📥 Скачать CSV", csv, f"data_{selected_year}_{selected_month_num}.csv", "text/csv")
else:
    st.warning("Нет данных")

st.caption(f"📅 {selected_month_display} {selected_year} | Записей: {format_number(len(df_filtered))}")
