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
    
    # Обработка столбца 'Период.Месяц' сразу при загрузке
    # Преобразуем в числовой формат, строки типа 'Итого' станут NaN
    df['Период.Месяц'] = pd.to_numeric(df['Период.Месяц'], errors='coerce')
    
    return df

df = load_data()

# Убираем строки с NaN в 'Период.Месяц' (бывший 'Итого')
df = df.dropna(subset=['Период.Месяц'])
df['Период.Месяц'] = df['Период.Месяц'].astype(int)

# ==========================================
# 3. ПОЛУЧАЕМ ДОСТУПНЫЕ ГОДЫ
# ==========================================
available_years = sorted(df['Год'].dropna().unique())
if len(available_years) == 0:
    available_years = [2024]

# ==========================================
# 4. БОКОВАЯ ПАНЕЛЬ С ФИЛЬТРАМИ
# ==========================================
st.sidebar.header("🔍 Фильтры")

selected_year = st.sidebar.selectbox("📅 Выберите год", available_years)

df_year = df[df['Год'] == selected_year]

# Получаем доступные месяцы для выбранного года (из 'Период.Месяц')
available_months_num = sorted(df_year['Период.Месяц'].unique())
# Преобразуем номера месяцев в формат для отображения
month_names = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
    5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
    9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}
available_months_display = [f"{month_names[m]} {selected_year}" for m in available_months_num]
month_display_to_num = {f"{month_names[m]} {selected_year}": m for m in available_months_num}

selected_month_display = st.sidebar.selectbox("Выберите месяц", available_months_display)
selected_month_num = month_display_to_num[selected_month_display]

all_customers = sorted(df_year['Контрагент'].dropna().unique())
selected_customers = st.sidebar.multiselect(
    "Выберите контрагентов",
    all_customers,
    default=all_customers[:5] if len(all_customers) > 5 else all_customers
)

df_filtered = df_year[(df_year['Период.Месяц'] == selected_month_num) & (df_year['Контрагент'].isin(selected_customers))]

# ==========================================
# 5. ГОДОВЫЕ МЕТРИКИ
# ==========================================
st.divider()
st.subheader(f"📈 ИТОГИ ЗА {selected_year} ГОД")

year_revenue = df_year['Выручка'].sum()
year_profit = df_year['Валовая_прибыль'].sum()
year_margin = (year_profit / year_revenue * 100) if year_revenue > 0 else 0
year_quantity = df_year['Количество'].sum()

col_y1, col_y2, col_y3, col_y4 = st.columns(4)

with col_y1:
    st.metric("💰 Годовая выручка", f"{format_number(year_revenue)} ₽")
with col_y2:
    st.metric("📈 Годовая валовая прибыль", f"{format_number(year_profit)} ₽")
with col_y3:
    st.metric("🎯 Годовая рентабельность", f"{format_float(year_margin, 1)}%")
with col_y4:
    st.metric("📦 Продано за год (шт)", f"{format_number(year_quantity)}")

if len(available_years) > 1 and selected_year > min(available_years):
    prev_year = selected_year - 1
    if prev_year in available_years:
        df_prev = df[df['Год'] == prev_year]
        prev_revenue = df_prev['Выручка'].sum()
        revenue_change = ((year_revenue - prev_revenue) / prev_revenue * 100) if prev_revenue > 0 else 0
        st.caption(f"📊 Изменение выручки относительно {prev_year} года: {format_float(revenue_change, 1)}%")

st.divider()

# ==========================================
# 6. ПОМЕСЯЧНАЯ РАЗБИВКА
# ==========================================
st.subheader(f"📅 ПОМЕСЯЧНАЯ РАЗБИВКА ЗА {selected_year} ГОД")

monthly_summary = df_year.groupby('Период.Месяц').agg({
    'Выручка': 'sum',
    'Валовая_прибыль': 'sum',
    'Количество': 'sum'
}).reset_index()

monthly_summary['Название_месяца'] = monthly_summary['Период.Месяц'].map(month_names)
monthly_summary['Рентабельность_%'] = (monthly_summary['Валовая_прибыль'] / monthly_summary['Выручка'] * 100).fillna(0)
monthly_summary = monthly_summary.sort_values('Период.Месяц')

def render_small_metric(label, value, suffix=""):
    st.markdown(
        f"""
        <div style='
            background-color: #F0F2F6;
            border-radius: 10px;
            padding: 10px;
            text-align: center;
        '>
            <div style='font-size: 14px; color: #666; margin-bottom: 5px;'>{label}</div>
            <div style='font-size: 20px; font-weight: bold; color: #1f1f1f;'>{value}{suffix}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

for idx, row in monthly_summary.iterrows():
    cols = st.columns([1.5, 1, 1, 1, 1])
    with cols[0]:
        st.markdown(f"<div style='font-weight: bold; font-size: 16px; padding-top: 12px;'>{row['Название_месяца']}</div>", unsafe_allow_html=True)
    with cols[1]:
        render_small_metric("Выручка", format_number(row['Выручка']), " ₽")
    with cols[2]:
        render_small_metric("Прибыль", format_number(row['Валовая_прибыль']), " ₽")
    with cols[3]:
        render_small_metric("Рентабельность", format_float(row['Рентабельность_%'], 1), "%")
    with cols[4]:
        render_small_metric("Кол-во (шт)", format_number(row['Количество']))

st.markdown("---")
total_cols = st.columns([1.5, 1, 1, 1, 1])
with total_cols[0]:
    st.markdown("<div style='font-weight: bold; font-size: 16px;'>📊 ИТОГО</div>", unsafe_allow_html=True)
with total_cols[1]:
    st.markdown(f"<div style='font-size: 16px;'><b>{format_number(year_revenue)} ₽</b></div>", unsafe_allow_html=True)
with total_cols[2]:
    st.markdown(f"<div style='font-size: 16px;'><b>{format_number(year_profit)} ₽</b></div>", unsafe_allow_html=True)
with total_cols[3]:
    st.markdown(f"<div style='font-size: 16px;'><b>{format_float(year_margin, 1)}%</b></div>", unsafe_allow_html=True)
with total_cols[4]:
    st.markdown(f"<div style='font-size: 16px;'><b>{format_number(year_quantity)}</b></div>", unsafe_allow_html=True)

st.divider()

# Остальные блоки (7-11) остаются без изменений...
# (продолжение следует, так как сообщение слишком длинное)
