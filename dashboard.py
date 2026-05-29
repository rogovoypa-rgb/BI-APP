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
# 1. ЗАГРУЗКА ДАННЫХ
# ==========================================
@st.cache_data
def load_sales_data():
    df = pd.read_excel('sales_data.xlsx')
    
    # Работа с датой (столбец 'Документ.Дата')
    df['Дата'] = pd.to_datetime(df['Документ.Дата'], dayfirst=True, errors='coerce')
    df['Месяц'] = df['Дата'].dt.strftime('%Y-%m')
    df['Год'] = df['Дата'].dt.year
    
    # Расчёт прибыли и рентабельности
    # ВАЖНО: 
    # - Выручка без НДС берётся из столбца 'Сумма без НДС' (столбец R)
    # - Себестоимость без НДС берётся из столбца 'Себестоимость без НДС' (столбец T)
    df['Валовая_прибыль'] = df['Сумма без НДС'] - df['Себестоимость без НДС']
    df['Рентабельность_%'] = (df['Валовая_прибыль'] / df['Сумма без НДС'] * 100).fillna(0)
    
    # Переименуем для удобства
    df = df.rename(columns={
        'Сумма без НДС': 'Выручка_без_НДС',
        'Себестоимость без НДС': 'Себестоимость',
        'Контрагент': 'Контрагент',
        'Номенклатура': 'Номенклатура'
    })
    
    # Обработка столбца 'Период.Месяц'
    df['Период.Месяц'] = pd.to_numeric(df['Период.Месяц'], errors='coerce')
    df = df.dropna(subset=['Период.Месяц'])
    df['Период.Месяц'] = df['Период.Месяц'].astype(int)
    
    return df

@st.cache_data
def load_logistics_data():
    try:
        # Читаем файл с заголовками в первой строке
        df = pd.read_excel('logistics_data.xlsx', header=0)
        
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        # Фильтруем строки с флагом 1 (только нужные для анализа)
        if 'Строка содержит данные' in df.columns:
            df = df[df['Строка содержит данные'] == 1]
        
        # Преобразуем числовые колонки
        numeric_cols = ['Плановая цена PLM', 'Фактическая цена PLM', 'Доставка до PLM', 
                        'Стоимость доставки 1 паллета', 'Кол-во паллет']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df
    except FileNotFoundError:
        return pd.DataFrame()
    except Exception as e:
        st.warning(f"Ошибка загрузки логистики: {e}")
        return pd.DataFrame()

sales_df = load_sales_data()
logistics_df = load_logistics_data()

# ==========================================
# 2. НАЗВАНИЯ МЕСЯЦЕВ
# ==========================================
month_names = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
    5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
    9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}

# ==========================================
# 3. НАВИГАЦИЯ ПО СТРАНИЦАМ
# ==========================================
st.set_page_config(page_title="BI Портал", layout="wide")

st.sidebar.title("📊 Навигация")
page = st.sidebar.radio(
    "Выберите раздел",
    ["📈 Продажи", "🚚 Логистика"]
)

# ==========================================
# СТРАНИЦА 1: ПРОДАЖИ
# ==========================================
if page == "📈 Продажи":
    st.title("📊 BI Портал аналитики продаж")
    
    available_years = sorted(sales_df['Год'].dropna().unique())
    if len(available_years) == 0:
        available_years = [2024]
    
    selected_year = st.sidebar.selectbox("📅 Выберите год", available_years)
    df_year = sales_df[sales_df['Год'] == selected_year]
    
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
    
    # Годовые метрики
    st.divider()
    st.subheader(f"📈 ИТОГИ ЗА {selected_year} ГОД")
    
    year_revenue = df_year['Выручка_без_НДС'].sum()
    year_profit = df_year['Валовая_прибыль'].sum()
    year_margin = (year_profit / year_revenue * 100) if year_revenue > 0 else 0
    year_quantity = df_year['Количество'].sum()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("💰 Выручка без НДС", f"{format_number(year_revenue)} ₽")
    with col2:
        st.metric("📈 Валовая прибыль", f"{format_number(year_profit)} ₽")
    with col3:
        st.metric("🎯 Рентабельность", f"{format_float(year_margin, 1)}%")
    with col4:
        st.metric("📦 Продано (шт)", f"{format_number(year_quantity)}")
    
    st.divider()
    
    # Помесячная разбивка
    st.subheader(f"📅 ПОМЕСЯЧНАЯ РАЗБИВКА ЗА {selected_year} ГОД")
    
    monthly = df_year.groupby('Период.Месяц').agg({
        'Выручка_без_НДС': 'sum',
        'Валовая_прибыль': 'sum',
        'Количество': 'sum'
    }).reset_index()
    monthly['Название'] = monthly['Период.Месяц'].map(month_names)
    monthly['Рентабельность'] = (monthly['Валовая_прибыль'] / monthly['Выручка_без_НДС'] * 100).fillna(0)
    monthly = monthly.sort_values('Период.Месяц')
    
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
    
    for _, row in monthly.iterrows():
        cols = st.columns([1.5, 1, 1, 1, 1])
        with cols[0]:
            st.markdown(f"<div style='font-weight: bold; font-size: 16px; padding-top: 12px;'>{row['Название']}</div>", unsafe_allow_html=True)
        with cols[1]:
            render_small_metric("Выручка без НДС", format_number(row['Выручка_без_НДС']), " ₽")
        with cols[2]:
            render_small_metric("Прибыль", format_number(row['Валовая_прибыль']), " ₽")
        with cols[3]:
            render_small_metric("Рентабельность", format_float(row['Рентабельность'], 1), "%")
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
    
    # Топ-5 контрагентов
    st.subheader(f"🏆 ТОП-5 КОНТРАГЕНТОВ ЗА {selected_year}")
    
    cust_rev = df_year.groupby('Контрагент')['Выручка_без_НДС'].sum().reset_index()
    cust_rev = cust_rev.sort_values('Выручка_без_НДС', ascending=False)
    top5 = cust_rev.head(5)['Контрагент'].tolist()
    
    monthly_cust = df_year.groupby(['Контрагент', 'Период.Месяц'])['Выручка_без_НДС'].sum().reset_index()
    
    table_data = []
    for c in top5:
        row = {'Контрагент': c}
        row['Год'] = cust_rev[cust_rev['Контрагент'] == c]['Выручка_без_НДС'].values[0]
        for m in available_months_num:
            val = monthly_cust[(monthly_cust['Контрагент'] == c) & (monthly_cust['Период.Месяц'] == m)]['Выручка_без_НДС'].sum()
            row[month_names[m]] = val
        table_data.append(row)
    
    other_rev = cust_rev[~cust_rev['Контрагент'].isin(top5)]['Выручка_без_НДС'].sum()
    other_row = {'Контрагент': '📦 ОСТАЛЬНЫЕ'}
    other_row['Год'] = other_rev
    for m in available_months_num:
        val = monthly_cust[(~monthly_cust['Контрагент'].isin(top5)) & (monthly_cust['Период.Месяц'] == m)]['Выручка_без_НДС'].sum()
        other_row[month_names[m]] = val
    table_data.append(other_row)
    
    df_top5 = pd.DataFrame(table_data)
    
    def fmt(x):
        return f"{int(x):,}".replace(",", " ") if x > 0 else "0"
    
    html = '<table style="width:100%; border-collapse:collapse">'
    html += '<tr style="background:#2E86AB; color:white">'
    html += '<th style="padding:8px">Контрагент</th><th>💰 Выручка без НДС за год</th>'
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
    html += '<tr>'
    
    st.markdown(html, unsafe_allow_html=True)
    
    total_top5 = cust_rev[cust_rev['Контрагент'].isin(top5)]['Выручка_без_НДС'].sum()
    st.caption(f"📊 Топ-5: {format_number(total_top5)} ₽ ({format_float(total_top5/year_revenue*100,1)}% от общей выручки без НДС)")
    st.divider()
    
    # Детали за месяц
    st.subheader(f"📊 ДЕТАЛИ ЗА {selected_month_display} {selected_year}")
    
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("💰 Выручка без НДС", f"{format_number(df_filtered['Выручка_без_НДС'].sum())} ₽")
    with m2:
        st.metric("📈 Валовая прибыль", f"{format_number(df_filtered['Валовая_прибыль'].sum())} ₽")
    with m3:
        marg = df_filtered['Валовая_прибыль'].sum() / df_filtered['Выручка_без_НДС'].sum() * 100 if df_filtered['Выручка_без_НДС'].sum() > 0 else 0
        st.metric("🎯 Рентабельность", f"{format_float(marg, 1)}%")
    with m4:
        st.metric("📦 Продано (шт)", f"{format_number(df_filtered['Количество'].sum())}")
    
    # Графики
    c1, c2 = st.columns(2)
    with c1:
        top10 = df_filtered.groupby('Контрагент')['Выручка_без_НДС'].sum().nlargest(10).reset_index()
        if not top10.empty:
            fig = px.bar(top10, x='Выручка_без_НДС', y='Контрагент', orientation='h', 
                         title='Топ-10 контрагентов по выручке без НДС',
                         labels={'Выручка_без_НДС': 'Выручка без НДС (₽)'})
            st.plotly_chart(fig, use_container_width=True)
    
    with c2:
        comp = df_filtered.groupby('Контрагент')[['Выручка_без_НДС', 'Валовая_прибыль']].sum().nlargest(10, 'Выручка_без_НДС').reset_index()
        if not comp.empty:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(name='Выручка без НДС', x=comp['Контрагент'], y=comp['Выручка_без_НДС'], marker_color='#2E86AB'))
            fig2.add_trace(go.Bar(name='Валовая прибыль', x=comp['Контрагент'], y=comp['Валовая_прибыль'], marker_color='#52B788'))
            fig2.update_layout(title='Выручка без НДС vs Валовая прибыль', barmode='group')
            st.plotly_chart(fig2, use_container_width=True)
    
    # Таблица данных
    st.subheader("📋 Детальные данные")
    show_cols = ['Дата', 'Контрагент', 'Номенклатура', 'Выручка_без_НДС', 'Валовая_прибыль', 'Рентабельность_%', 'Количество']
    show_cols = [c for c in show_cols if c in df_filtered.columns]
    if not df_filtered.empty:
        # Переименовываем для отображения в таблице
        df_display = df_filtered[show_cols].copy()
        df_display = df_display.rename(columns={'Выручка_без_НДС': 'Выручка без НДС'})
        df_display['Выручка без НДС'] = df_display['Выручка без НДС'].apply(lambda x: f"{format_number(x)} ₽")
        df_display['Валовая_прибыль'] = df_display['Валовая_прибыль'].apply(lambda x: f"{format_number(x)} ₽")
        df_display['Рентабельность_%'] = df_display['Рентабельность_%'].apply(lambda x: f"{format_float(x, 1)}%")
        df_display['Количество'] = df_display['Количество'].apply(format_number)
        
        st.dataframe(df_display.head(100), use_container_width=True)
        
        csv = df_filtered[show_cols].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
        st.download_button("📥 Скачать CSV", csv, f"data_{selected_year}_{selected_month_num}.csv", "text/csv")
    else:
        st.warning("Нет данных")
    
    st.caption(f"📅 {selected_month_display} {selected_year} | Записей: {format_number(len(df_filtered))}")

# ==========================================
# СТРАНИЦА 2: ЛОГИСТИКА
# ==========================================
elif page == "🚚 Логистика":
    st.title("🚚 Аналитика затрат на логистику")
    
    if logistics_df.empty:
        st.warning("⚠️ Файл 'logistics_data.xlsx' не найден или пуст.")
        st.info("📌 Пожалуйста, добавьте файл с данными по логистике в папку с приложением.")
        
        # Диагностика: показываем, какие файлы есть в папке
        import os
        with st.expander("🔧 Диагностика"):
            st.write("Файлы в текущей папке:")
            for f in os.listdir('.'):
                if f.endswith('.xlsx'):
                    st.write(f"  - {f}")
    else:
        # ==========================================
        # ПОДГОТОВКА ДАННЫХ
        # ==========================================
        df_log = logistics_df.copy()
        
        # Преобразуем даты
        if 'Дата заказа' in df_log.columns:
            df_log['Дата'] = pd.to_datetime(df_log['Дата заказа'], errors='coerce')
        elif 'Дата отгрузки' in df_log.columns:
            df_log['Дата'] = pd.to_datetime(df_log['Дата отгрузки'], errors='coerce')
        else:
            df_log['Дата'] = pd.Timestamp.now()
        
        df_log['Год'] = df_log['Дата'].dt.year
        df_log['Месяц'] = df_log['Дата'].dt.month
        df_log['Месяц_название'] = df_log['Месяц'].map(month_names)
        
        # Основные колонки для анализа
        df_log['План'] = pd.to_numeric(df_log['Плановая цена PLM'], errors='coerce')
        df_log['Факт'] = pd.to_numeric(df_log['Фактическая цена PLM'], errors='coerce')
        df_log['Доставка_до_PLM'] = pd.to_numeric(df_log['Доставка до PLM'], errors='coerce')
        df_log['Кол_во_паллет'] = pd.to_numeric(df_log['Кол-во паллет'], errors='coerce')
        df_log['Отклонение'] = df_log['Факт'] - df_log['План']
        
        # ==========================================
        # ФИЛЬТРЫ
        # ==========================================
        st.sidebar.header("🚚 Фильтры логистики")
        
        available_years = sorted(df_log['Год'].dropna().unique())
        if len(available_years) == 0:
            available_years = [2024]
        selected_year_log = st.sidebar.selectbox("📅 Выберите год", available_years, key="log_year")
        
        df_year_log = df_log[df_log['Год'] == selected_year_log]
        
        available_months = sorted(df_year_log['Месяц'].dropna().unique())
        available_months_display = [month_names[m] for m in available_months]
        selected_month_display = st.sidebar.selectbox("Выберите месяц", available_months_display, key="log_month")
        selected_month_num = available_months[available_months_display.index(selected_month_display)]
        
        # Города
        if 'Город' in df_year_log.columns:
            all_cities = sorted(df_year_log['Город'].dropna().unique())
            selected_cities = st.sidebar.multiselect(
                "🏙️ Города",
                all_cities,
                default=all_cities[:5] if len(all_cities) > 5 else all_cities,
                key="log_cities"
            )
        else:
            selected_cities = []
        
        # Фильтруем данные
        mask = (df_year_log['Месяц'] == selected_month_num)
        if selected_cities:
            mask = mask & (df_year_log['Город'].isin(selected_cities))
        df_filtered_log = df_year_log[mask]
        
        # ==========================================
        # ГОДОВЫЕ МЕТРИКИ
        # ==========================================
        st.divider()
        st.subheader(f"📦 ИТОГИ ЛОГИСТИКИ ЗА {selected_year_log} ГОД")
        
        year_plan = df_year_log['План'].sum()
        year_fact = df_year_log['Факт'].sum()
        year_deviation = year_fact - year_plan
        year_deviation_pct = (year_deviation / year_plan * 100) if year_plan > 0 else 0
        year_delivery = df_year_log['Доставка_до_PLM'].sum()
        year_orders = len(df_year_log)
        year_pallets = df_year_log['Кол_во_паллет'].sum()
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("💰 Фактические затраты", f"{format_number(year_fact)} ₽")
        with c2:
            st.metric("📋 Плановые затраты", f"{format_number(year_plan)} ₽")
        with c3:
            delta_color = "normal" if year_deviation <= 0 else "inverse"
            st.metric("📊 Отклонение", f"{format_number(year_deviation)} ₽", 
                     delta=f"{format_float(year_deviation_pct, 1)}%")
        with c4:
            st.metric("📦 Кол-во паллет", f"{format_number(year_pallets)}")
        
        # Вторая строка метрик
        c5, c6, c7, c8 = st.columns(4)
        with c5:
            st.metric("🚛 Доставка до PLM", f"{format_number(year_delivery)} ₽")
        with c6:
            st.metric("📋 Кол-во заказов", f"{format_number(year_orders)}")
        with c7:
            avg_cost = year_fact / year_orders if year_orders > 0 else 0
            st.metric("💰 Средняя цена заказа", f"{format_number(avg_cost)} ₽")
        with c8:
            avg_pallet = year_fact / year_pallets if year_pallets > 0 else 0
            st.metric("📦 Средняя цена паллеты", f"{format_number(avg_pallet)} ₽")
        
        st.divider()
        
        # ==========================================
        # ПОМЕСЯЧНАЯ РАЗБИВКА
        # ==========================================
        st.subheader(f"📅 ПОМЕСЯЧНАЯ РАЗБИВКА ЗА {selected_year_log} ГОД")
        
        monthly = df_year_log.groupby('Месяц').agg({
            'План': 'sum',
            'Факт': 'sum',
            'Доставка_до_PLM': 'sum',
            'Кол_во_паллет': 'sum'
        }).reset_index()
        monthly['Название'] = monthly['Месяц'].map(month_names)
        monthly['Отклонение'] = monthly['Факт'] - monthly['План']
        monthly = monthly.sort_values('Месяц')
        
        for _, row in monthly.iterrows():
            cols = st.columns([1.5, 1, 1, 1, 1, 1])
            with cols[0]:
                st.markdown(f"**{row['Название']}**")
            with cols[1]:
                st.metric("План", f"{format_number(row['План'])} ₽", label_visibility="collapsed")
            with cols[2]:
                st.metric("Факт", f"{format_number(row['Факт'])} ₽", label_visibility="collapsed")
            with cols[3]:
                dev_color = "🟢" if row['Отклонение'] <= 0 else "🔴"
                st.metric("Отклонение", f"{dev_color} {format_number(row['Отклонение'])} ₽", label_visibility="collapsed")
            with cols[4]:
                st.metric("Доставка до PLM", f"{format_number(row['Доставка_до_PLM'])} ₽", label_visibility="collapsed")
            with cols[5]:
                st.metric("Паллет", format_number(row['Кол_во_паллет']), label_visibility="collapsed")
        
        st.divider()
        
        # ==========================================
        # ГРАФИКИ
        # ==========================================
        col1, col2 = st.columns(2)
        
        with col1:
            # Топ городов по затратам
            if 'Город' in df_year_log.columns:
                city_costs = df_year_log.groupby('Город')['Факт'].sum().nlargest(10).reset_index()
                if not city_costs.empty:
                    fig = px.bar(city_costs, x='Факт', y='Город', orientation='h',
                                 title='Топ городов по фактическим затратам',
                                 color='Факт', color_continuous_scale='Reds',
                                 labels={'Факт': 'Затраты (₽)', 'Город': ''})
                    st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Динамика план vs факт по месяцам
            if not monthly.empty:
                fig2 = go.Figure()
                fig2.add_trace(go.Scatter(x=monthly['Название'], y=monthly['План'], 
                                          mode='lines+markers', name='План', line=dict(color='blue')))
                fig2.add_trace(go.Scatter(x=monthly['Название'], y=monthly['Факт'], 
                                          mode='lines+markers', name='Факт', line=dict(color='red')))
                fig2.update_layout(title='Динамика план vs факт по месяцам',
                                   xaxis_title='Месяц', yaxis_title='Затраты (₽)')
                st.plotly_chart(fig2, use_container_width=True)
        
        # Третий график
        if not monthly.empty:
            st.subheader("📊 Отклонение факта от плана по месяцам")
            fig3 = px.bar(monthly, x='Название', y='Отклонение',
                          title='Отклонение (Факт - План)',
                          color='Отклонение', color_continuous_scale='RdYlGn',
                          labels={'Отклонение': 'Отклонение (₽)', 'Название': 'Месяц'})
            st.plotly_chart(fig3, use_container_width=True)
        
        # ==========================================
        # ДЕТАЛИ ЗА ВЫБРАННЫЙ МЕСЯЦ
        # ==========================================
        st.divider()
        st.subheader(f"📊 ДЕТАЛИ ЗА {selected_month_display} {selected_year_log}")
        
        d1, d2, d3, d4 = st.columns(4)
        with d1:
            st.metric("💰 Факт", f"{format_number(df_filtered_log['Факт'].sum())} ₽")
        with d2:
            st.metric("📋 План", f"{format_number(df_filtered_log['План'].sum())} ₽")
        with d3:
            dev = df_filtered_log['Факт'].sum() - df_filtered_log['План'].sum()
            st.metric("📊 Отклонение", f"{format_number(dev)} ₽")
        with d4:
            st.metric("📦 Паллет", f"{format_number(df_filtered_log['Кол_во_паллет'].sum())}")
        
        # Таблица с детальными данными
        st.subheader("📋 Детальные данные по логистике")
        display_cols = ['Дата заказа', 'Город', 'Плановая цена PLM', 'Фактическая цена PLM', 
                        'Доставка до PLM', 'Кол-во паллет', 'Категория', 'Номенклатура']
        display_cols = [c for c in display_cols if c in df_filtered_log.columns]
        
        if not df_filtered_log.empty and display_cols:
            df_display = df_filtered_log[display_cols].copy()
            for col in ['Плановая цена PLM', 'Фактическая цена PLM', 'Доставка до PLM']:
                if col in df_display.columns:
                    df_display[col] = df_display[col].apply(lambda x: f"{format_number(x)} ₽" if pd.notna(x) else "0 ₽")
            
            st.dataframe(df_display.head(100), use_container_width=True)
            
            csv_log = df_filtered_log[display_cols].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
            st.download_button("📥 Скачать CSV (логистика)", csv_log, f"logistics_{selected_year_log}_{selected_month_num}.csv", "text/csv")
        else:
            st.warning("Нет данных за выбранный период")
        
        st.caption(f"📅 {selected_month_display} {selected_year_log} | Записей: {len(df_filtered_log)}")
