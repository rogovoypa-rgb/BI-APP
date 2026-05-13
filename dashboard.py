import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# ФУНКЦИИ ДЛЯ ФОРМАТИРОВАНИЯ ЧИСЕЛ
# ==========================================
def format_number(value):
    """Форматирует число с пробелами между разрядами"""
    if pd.isna(value):
        return "0"
    try:
        return f"{int(value):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(value)

def format_float(value, decimals=1):
    """Форматирует число с плавающей точкой (запятая как разделитель)"""
    if pd.isna(value):
        return "0"
    try:
        # Округляем
        rounded = round(value, decimals)
        # Преобразуем в строку
        if decimals == 0:
            formatted = str(int(rounded))
        else:
            # Разделяем целую и дробную части
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
    
    # Работа с датой (столбец 'Документ.Дата')
    df['Дата'] = pd.to_datetime(df['Документ.Дата'], dayfirst=True, errors='coerce')
    df['Месяц'] = df['Дата'].dt.strftime('%Y-%m')
    df['Год'] = df['Дата'].dt.year
    
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
# 5. ГОДОВЫЕ МЕТРИКИ
# ==========================================
st.divider()
st.subheader(f"📈 ИТОГИ ЗА {selected_year} ГОД")

# Рассчитываем годовые показатели
year_revenue = df_year['Выручка'].sum()
year_profit = df_year['Валовая_прибыль'].sum()
year_margin = (year_profit / year_revenue * 100) if year_revenue > 0 else 0
year_quantity = df_year['Количество'].sum()

# Отображаем 4 метрики в ряд
col_y1, col_y2, col_y3, col_y4 = st.columns(4)

with col_y1:
    st.metric("💰 Годовая выручка", f"{format_number(year_revenue)} ₽")

with col_y2:
    st.metric("📈 Годовая валовая прибыль", f"{format_number(year_profit)} ₽")

with col_y3:
    st.metric("🎯 Годовая рентабельность", f"{format_float(year_margin, 1)}%")

with col_y4:
    st.metric("📦 Продано за год (шт)", f"{format_number(year_quantity)}")

# Изменение по сравнению с предыдущим годом
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

# Создаём словарь для преобразования номера месяца в название
month_names_full = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
    5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
    9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}

# Добавляем название месяца в df_year (на основе столбца 'Период.Месяц')
df_year_monthly = df_year.copy()
df_year_monthly = df_year_monthly[df_year_monthly['Период.Месяц'] != 'Итого'].copy()
df_year_monthly['Период.Месяц'] = pd.to_numeric(df_year_monthly['Период.Месяц'], errors='coerce')
df_year_monthly = df_year_monthly.dropna(subset=['Период.Месяц'])
df_year_monthly['Название_месяца'] = df_year_monthly['Период.Месяц'].map(month_names_full)

# Группируем данные по месяцам
monthly_summary = df_year_monthly.groupby(['Период.Месяц', 'Название_месяца']).agg({
    'Выручка': 'sum',
    'Валовая_прибыль': 'sum',
    'Количество': 'sum'
}).reset_index()

# Рассчитываем рентабельность по месяцам
monthly_summary['Рентабельность_%'] = (monthly_summary['Валовая_прибыль'] / monthly_summary['Выручка'] * 100).fillna(0)

# Сортируем по номеру месяца
monthly_summary = monthly_summary.sort_values('Период.Месяц')

# Функция для отображения уменьшенных метрик
def render_small_metric(label, value, suffix=""):
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
            '>{value}{suffix}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# Отображаем таблицу с помесячными данными
for idx, row in monthly_summary.iterrows():
    month_name = row['Название_месяца']
    
    cols = st.columns([1.5, 1, 1, 1, 1])
    
    with cols[0]:
        st.markdown(f"<div style='font-weight: bold; font-size: 16px; padding-top: 12px;'>{month_name}</div>", unsafe_allow_html=True)
    
    with cols[1]:
        render_small_metric("Выручка", format_number(row['Выручка']), " ₽")
    
    with cols[2]:
        render_small_metric("Прибыль", format_number(row['Валовая_прибыль']), " ₽")
    
    with cols[3]:
        render_small_metric("Рентабельность", format_float(row['Рентабельность_%'], 1), "%")
    
    with cols[4]:
        render_small_metric("Кол-во (шт)", format_number(row['Количество']))

# Итоговая строка
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

# ==========================================
# 7. АНАЛИЗ ВЫРУЧКИ ПО ТОП-5 КОНТРАГЕНТАМ
# ==========================================
st.subheader(f"🏆 АНАЛИЗ ВЫРУЧКИ ПО ТОП-5 КОНТРАГЕНТАМ ЗА {selected_year} ГОД")

# Очищаем данные
df_year_clean = df_year[df_year['Период.Месяц'] != 'Итого'].copy()
df_year_clean['Период.Месяц'] = pd.to_numeric(df_year_clean['Период.Месяц'], errors='coerce')
df_year_clean = df_year_clean.dropna(subset=['Период.Месяц'])

# Словарь для преобразования номера месяца в название
month_names = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
    5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
    9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}

df_year_clean['Название_месяца'] = df_year_clean['Период.Месяц'].map(month_names)

# Рассчитываем выручку по контрагентам за год
customer_revenue = df_year_clean.groupby('Контрагент')['Выручка'].sum().reset_index()
customer_revenue = customer_revenue.sort_values('Выручка', ascending=False)

# Топ-5 контрагентов
top5_customers = customer_revenue.head(5)['Контрагент'].tolist()

# Помесячная выручка
monthly_customer_revenue = df_year_clean.groupby(['Контрагент', 'Период.Месяц', 'Название_месяца'])['Выручка'].sum().reset_index()

month_order = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 
               'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']

# Строим таблицу
table_data = []

for customer in top5_customers:
    row = {'Контрагент': customer}
    yearly = customer_revenue[customer_revenue['Контрагент'] == customer]['Выручка'].values[0]
    row['Выручка за год'] = yearly
    
    for month in month_order:
        month_value = monthly_customer_revenue[
            (monthly_customer_revenue['Контрагент'] == customer) & 
            (monthly_customer_revenue['Название_месяца'] == month)
        ]['Выручка'].sum()
        row[month] = month_value
    table_data.append(row)

# Остальные
other_customers = customer_revenue[~customer_revenue['Контрагент'].isin(top5_customers)]
other_yearly = other_customers['Выручка'].sum()

row_other = {'Контрагент': '📦 ОСТАЛЬНЫЕ (все прочие контрагенты)'}
row_other['Выручка за год'] = other_yearly

for month in month_order:
    month_value = monthly_customer_revenue[
        (~monthly_customer_revenue['Контрагент'].isin(top5_customers)) & 
        (monthly_customer_revenue['Название_месяца'] == month)
    ]['Выручка'].sum()
    row_other[month] = month_value

table_data.append(row_other)

df_top5_table = pd.DataFrame(table_data)

# Форматируем числа с пробелами для отображения в HTML
def format_with_spaces(x):
    if pd.isna(x):
        return "0"
    try:
        return f"{int(x):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(x)

# Создаём HTML-таблицу для кастомного отображения
html_table = """
<style>
.top5-table {
    width: 100%;
    border-collapse: collapse;
    font-family: 'Segoe UI', Arial, sans-serif;
}
.top5-table th {
    background-color: #2E86AB;
    color: white;
    padding: 10px;
    text-align: center;
    font-weight: bold;
    position: sticky;
    top: 0;
}
.top5-table td {
    padding: 8px;
    text-align: right;
    border-bottom: 1px solid #e0e0e0;
}
.top5-table td:first-child {
    text-align: left;
    font-weight: bold;
    background-color: #f8f9fa;
}
.top5-table tr:last-child td:first-child {
    background-color: #e8f0fe;
}
.top5-table td:last-child {
    text-align: right;
}
.top5-table tr:last-child td {
    font-weight: bold;
    background-color: #f0f2f6;
}
.month-cell {
    font-size: 12px;
}
.year-cell {
    font-weight: bold;
}
</style>

<table class="top5-table">
    <thead>
        <tr>
            <th>Контрагент / Покупатель</th>
            <th>💰 Выручка за год</th>
"""

# Добавляем заголовки месяцев
for month in month_order:
    short_name = month[:3]  # Янв, Фев, Мар и т.д.
    html_table += f"<th>{short_name}</th>"

html_table += """
        </tr>
    </thead>
    <tbody>
"""

# Добавляем строки с данными
for idx, row in df_top5_table.iterrows():
    html_table += "<tr>"
    
    # Контрагент
    html_table += f"<td style='font-weight: bold;'>{row['Контрагент']}</td>"
    
    # Выручка за год (жирный шрифт)
    html_table += f"<td style='font-weight: bold;'>{format_with_spaces(row['Выручка за год'])} ₽</td>"
    
    # Месяцы (уменьшенный шрифт 12px)
    for month in month_order:
        value = row[month]
        html_table += f"<td class='month-cell'>{format_with_spaces(value)} ₽</td>"
    
    html_table += "</tr>"

html_table += """
    </tbody>
</table>
"""

# Отображаем HTML-таблицу
st.markdown(html_table, unsafe_allow_html=True)

# Итоговая статистика
total_top5 = customer_revenue[customer_revenue['Контрагент'].isin(top5_customers)]['Выручка'].sum()
total_all = year_revenue
total_others = total_all - total_top5

st.caption(f"📊 Из {format_number(total_all)} ₽ общей выручки за {selected_year} год:")
st.caption(f"   • Топ-5 контрагентов: {format_number(total_top5)} ₽ ({format_float(total_top5/total_all*100, 1)}%)")
st.caption(f"   • Остальные контрагенты: {format_number(total_others)} ₽ ({format_float(total_others/total_all*100, 1)}%)")

st.divider()

# ==========================================
# 8. ОСНОВНЫЕ МЕТРИКИ (за выбранный месяц)
# ==========================================
st.subheader(f"📊 ДЕТАЛИ ЗА {selected_month}")

col1, col2, col3, col4 = st.columns(4)

with col1:
    total_revenue = df_filtered['Выручка'].sum()
    st.metric("💰 Выручка", f"{format_number(total_revenue)} ₽")

with col2:
    total_profit = df_filtered['Валовая_прибыль'].sum()
    st.metric("📈 Валовая прибыль", f"{format_number(total_profit)} ₽")

with col3:
    margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
    st.metric("🎯 Рентабельность", f"{format_float(margin, 1)}%")

with col4:
    total_quantity = df_filtered['Количество'].sum()
    st.metric("📦 Продано (шт)", f"{format_number(total_quantity)}")

# ==========================================
# 9. ГРАФИКИ
# ==========================================

# Два графика в ряд
col_left, col_right = st.columns(2)

with col_left:
    top_customers_chart = df_filtered.groupby('Контрагент')['Выручка'].sum().nlargest(10).reset_index()
    if not top_customers_chart.empty:
        fig1 = px.bar(top_customers_chart, x='Выручка', y='Контрагент', orientation='h',
                      title='Топ-10 контрагентов по выручке',
                      color='Выручка', color_continuous_scale='Blues',
                      labels={'Выручка': 'Выручка (₽)', 'Контрагент': ''})
        fig1.update_layout(height=500)
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.info("Нет данных для отображения")

with col_right:
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

# Третий график
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
# 10. ТАБЛИЦА С ДАННЫМИ
# ==========================================
st.subheader("📋 Детальные данные по продажам")
display_cols = ['Дата', 'Контрагент', 'Номенклатура', 'Выручка', 'Валовая_прибыль', 'Рентабельность_%', 'Количество']
display_cols = [col for col in display_cols if col in df_filtered.columns]

if not df_filtered.empty:
    # Форматируем копию таблицы для отображения
    df_display = df_filtered[display_cols].copy()
    df_display['Выручка'] = df_display['Выручка'].apply(lambda x: f"{format_number(x)} ₽")
    df_display['Валовая_прибыль'] = df_display['Валовая_прибыль'].apply(lambda x: f"{format_number(x)} ₽")
    df_display['Рентабельность_%'] = df_display['Рентабельность_%'].apply(lambda x: f"{format_float(x, 1)}%")
    df_display['Количество'] = df_display['Количество'].apply(format_number)
    
    st.dataframe(df_display.head(100), use_container_width=True)
    
    # Кнопка скачивания
    csv = df_filtered[display_cols].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
    st.download_button(
        label="📥 Скачать данные (CSV)",
        data=csv,
        file_name=f'sales_data_{selected_year}_{selected_month}.csv',
        mime='text/csv',
    )
else:
    st.warning("Нет данных для отображения. Попробуйте изменить фильтры.")

# ==========================================
# 11. ИТОГОВАЯ СТАТИСТИКА
# ==========================================
st.divider()
st.caption(f"📅 Данные за {selected_month} {selected_year} | Всего записей: {format_number(len(df_filtered))} | Обновлено: автоматически при запуске")
