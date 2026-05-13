import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

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
    
    # Работа с датой (столбец 'Документ.Дата')
    df['Дата'] = pd.to_datetime(df['Документ.Дата'], dayfirst=True, errors='coerce')
    df['Месяц_цифра'] = df['Дата'].dt.month
    df['Месяц'] = df['Дата'].dt.strftime('%Y-%m')
    df['Год'] = df['Дата'].dt.year
    df['Название_месяца'] = df['Дата'].dt.strftime('%B')  # Полное название месяца
    
    # Расчёт прибыли и рентабельности
    df['Валовая_прибыль'] = df['Сумма без НДС'] - df['Себестоимость']
    df['Рентабельность_%'] = (df['Валовая_прибыль'] / df['Сумма без НДС'] * 100).fillna(0)
    
    # Переименуем для удобства
    df = df.rename(columns={
        'Сумма': 'Выручка',
        'Контрагент': 'Контрагент',
        'Номенклатура': 'Номенклатура'
    })
    
    return df

df = load_data()

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

# Фильтр: выбор года
selected_year = st.sidebar.selectbox("📅 Выберите год", available_years)

# Фильтруем данные по году
df_year = df[df['Год'] == selected_year]

# Выбор месяца (из отфильтрованного года)
months = sorted(df_year['Месяц'].dropna().unique())
selected_month = st.sidebar.selectbox("Выберите месяц", months)

# Выбор контрагентов
all_customers = sorted(df_year['Контрагент'].dropna().unique())
selected_customers = st.sidebar.multiselect(
    "Выберите контрагентов",
    all_customers,
    default=all_customers[:5] if len(all_customers) > 5 else all_customers
)

# Применяем фильтры для детального просмотра
df_filtered = df_year[(df_year['Месяц'] == selected_month) & (df_year['Контрагент'].isin(selected_customers))]

# ==========================================
# 5. ГОДОВЫЕ МЕТРИКИ (крупно)
# ==========================================
st.divider()
st.subheader(f"📈 ИТОГИ ЗА {selected_year} ГОД")

# Рассчитываем годовые показатели
year_revenue = df_year['Выручка'].sum()
year_profit = df_year['Валовая_прибыль'].sum()
year_margin = (year_profit / year_revenue * 100) if year_revenue > 0 else 0
year_quantity = df_year['Количество'].sum()

# Отображаем 4 метрики в ряд (крупный размер)
col_y1, col_y2, col_y3, col_y4 = st.columns(4)

with col_y1:
    st.metric("💰 Годовая выручка", f"{year_revenue:,.0f} ₽")

with col_y2:
    st.metric("📈 Годовая валовая прибыль", f"{year_profit:,.0f} ₽")

with col_y3:
    st.metric("🎯 Годовая рентабельность", f"{year_margin:.1f}%")

with col_y4:
    st.metric("📦 Продано за год (шт)", f"{year_quantity:,.0f}")

# Изменение по сравнению с предыдущим годом
if len(available_years) > 1 and selected_year > min(available_years):
    prev_year = selected_year - 1
    if prev_year in available_years:
        df_prev = df[df['Год'] == prev_year]
        prev_revenue = df_prev['Выручка'].sum()
        revenue_change = ((year_revenue - prev_revenue) / prev_revenue * 100) if prev_revenue > 0 else 0
        st.caption(f"📊 Изменение выручки относительно {prev_year} года: {revenue_change:+.1f}%")

st.divider()

# ==========================================
# 6. НОВЫЙ БЛОК: ПОМЕСЯЧНАЯ РАЗБИВКА (уменьшенный размер)
# ==========================================
st.subheader(f"📅 ПОМЕСЯЧНАЯ РАЗБИВКА ЗА {selected_year} ГОД")

# Группируем данные по месяцам
monthly_summary = df_year.groupby(['Месяц', 'Название_месяца']).agg({
    'Выручка': 'sum',
    'Валовая_прибыль': 'sum',
    'Количество': 'sum'
}).reset_index()

# Рассчитываем рентабельность по месяцам
monthly_summary['Рентабельность_%'] = (monthly_summary['Валовая_прибыль'] / monthly_summary['Выручка'] * 100).fillna(0)

# Сортируем по месяцам (по дате, а не по алфавиту)
monthly_summary['Месяц_сортировка'] = monthly_summary['Месяц']
monthly_summary = monthly_summary.sort_values('Месяц_сортировка')

# Функция для отображения уменьшенных метрик (на 30% меньше)
def render_small_metric(label, value, help_text=None):
    st.markdown(
        f"""
        <div style='
            background-color: #F0F2F6;
            border-radius: 10px;
            padding: 10px;
            text-align: center;
        '>
            <div style='
                font-size: 14px;
                color: #666;
                margin-bottom: 5px;
            '>{label}</div>
            <div style='
                font-size: 20px;
                font-weight: bold;
                color: #1f1f1f;
            '>{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# Отображаем таблицу с помесячными данными
for idx, row in monthly_summary.iterrows():
    # Название месяца
    month_name = row['Название_месяца']
    
    # Создаём строку с 4 колонками
    cols = st.columns([1.5, 1, 1, 1, 1])  # Первая колонка шире (название месяца)
    
    with cols[0]:
        st.markdown(f"<div style='font-weight: bold; font-size: 16px; padding-top: 12px;'>{month_name}</div>", unsafe_allow_html=True)
    
    with cols[1]:
        render_small_metric("Выручка", f"{row['Выручка']:,.0f} ₽")
    
    with cols[2]:
        render_small_metric("Прибыль", f"{row['Валовая_прибыль']:,.0f} ₽")
    
    with cols[3]:
        render_small_metric("Рентабельность", f"{row['Рентабельность_%']:.1f}%")
    
    with cols[4]:
        render_small_metric("Кол-во (шт)", f"{row['Количество']:,.0f}")

# Добавляем строку-итог
st.markdown("---")
total_cols = st.columns([1.5, 1, 1, 1, 1])
with total_cols[0]:
    st.markdown("<div style='font-weight: bold; font-size: 16px;'>📊 ИТОГО</div>", unsafe_allow_html=True)
with total_cols[1]:
    st.markdown(f"<div style='font-size: 16px;'><b>{year_revenue:,.0f} ₽</b></div>", unsafe_allow_html=True)
with total_cols[2]:
    st.markdown(f"<div style='font-size: 16px;'><b>{year_profit:,.0f} ₽</b></div>", unsafe_allow_html=True)
with total_cols[3]:
    st.markdown(f"<div style='font-size: 16px;'><b>{year_margin:.1f}%</b></div>", unsafe_allow_html=True)
with total_cols[4]:
    st.markdown(f"<div style='font-size: 16px;'><b>{year_quantity:,.0f}</b></div>", unsafe_allow_html=True)

st.divider()

# ==========================================
# 7. ОСНОВНЫЕ МЕТРИКИ (за выбранный месяц)
# ==========================================
st.subheader(f"📊 ДЕТАЛИ ЗА {selected_month}")

col1, col2, col3, col4 = st.columns(4)

with col1:
    total_revenue = df_filtered['Выручка'].sum()
    st.metric("💰 Выручка", f"{total_revenue:,.0f} ₽")

with col2:
    total_profit = df_filtered['Валовая_прибыль'].sum()
    st.metric("📈 Валовая прибыль", f"{total_profit:,.0f} ₽")

with col3:
    margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
    st.metric("🎯 Рентабельность", f"{margin:.1f}%")

with col4:
    total_quantity = df_filtered['Количество'].sum()
    st.metric("📦 Продано (шт)", f"{total_quantity:,.0f}")

# ==========================================
# 8. ГРАФИКИ
# ==========================================

# Два графика в ряд
col_left, col_right = st.columns(2)

with col_left:
    # Топ-10 контрагентов
    top_customers = df_filtered.groupby('Контрагент')['Выручка'].sum().nlargest(10).reset_index()
    if not top_customers.empty:
        fig1 = px.bar(top_customers, x='Выручка', y='Контрагент', orientation='h',
                      title='Топ-10 контрагентов по выручке',
                      color='Выручка', color_continuous_scale='Blues',
                      labels={'Выручка': 'Выручка (₽)', 'Контрагент': ''})
        fig1.update_layout(height=500)
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.info("Нет данных для отображения")

with col_right:
    # Сравнение выручки и прибыли по контрагентам
    compare = df_filtered.groupby('Контрагент')[['Выручка', 'Валовая_прибыль']].sum().nlargest(10, 'Выручка').reset_index()
    if not compare.empty:
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(name='Выручка', x=compare['Контрагент'], y=compare['Выручка'], 
                              marker_color='#2E86AB', text=compare['Выручка'].apply(lambda x: f'{x:,.0f}'),
                              textposition='outside'))
        fig2.add_trace(go.Bar(name='Валовая прибыль', x=compare['Контрагент'], y=compare['Валовая_прибыль'], 
                              marker_color='#52B788', text=compare['Валовая_прибыль'].apply(lambda x: f'{x:,.0f}'),
                              textposition='outside'))
        fig2.update_layout(title='Сравнение выручки и прибыли',
                           barmode='group',
                           xaxis_tickangle=-45,
                           height=500)
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Нет данных для отображения")

# Третий график на всю ширину
st.subheader("📈 Динамика рентабельности по месяцам")
margin_by_month = df_year[df_year['Контрагент'].isin(selected_customers)].groupby('Месяц')['Рентабельность_%'].mean().reset_index()
if not margin_by_month.empty:
    fig3 = px.line(margin_by_month, x='Месяц', y='Рентабельность_%', 
                   markers=True, line_shape='linear',
                   title=f'Средняя рентабельность выбранных контрагентов за {selected_year} год',
                   labels={'Рентабельность_%': 'Рентабельность (%)', 'Месяц': ''})
    fig3.update_layout(hovermode='x unified', height=450)
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("Нет данных для отображения динамики")

# ==========================================
# 9. ТАБЛИЦА С ДАННЫМИ
# ==========================================
st.subheader("📋 Детальные данные по продажам")
display_cols = ['Дата', 'Контрагент', 'Номенклатура', 'Выручка', 'Валовая_прибыль', 'Рентабельность_%', 'Количество']
display_cols = [col for col in display_cols if col in df_filtered.columns]

if not df_filtered.empty:
    st.dataframe(df_filtered[display_cols].head(100), use_container_width=True)
    
    # Кнопка скачивания
    csv = df_filtered[display_cols].to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Скачать данные (CSV)",
        data=csv,
        file_name=f'sales_data_{selected_year}_{selected_month}.csv',
        mime='text/csv',
    )
else:
    st.warning("Нет данных для отображения. Попробуйте изменить фильтры.")

# ==========================================
# 10. ИТОГОВАЯ СТАТИСТИКА
# ==========================================
st.divider()
st.caption(f"📅 Данные за {selected_month} {selected_year} | Всего записей: {len(df_filtered)} | Обновлено: автоматически при запуске")
