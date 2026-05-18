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

@st.cache_data
def load_logistics_data():
    try:
        df = pd.read_excel('logistics_data.xlsx')
        df['Дата'] = pd.to_datetime(df['Дата'], dayfirst=True, errors='coerce')
        df['Месяц'] = df['Дата'].dt.strftime('%Y-%m')
        df['Год'] = df['Дата'].dt.year
        df['Месяц_цифра'] = df['Дата'].dt.month
        return df
    except FileNotFoundError:
        return pd.DataFrame()  # Пустой DataFrame, если файла нет

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

# Боковая панель с навигацией
st.sidebar.title("📊 Навигация")
page = st.sidebar.radio(
    "Выберите раздел",
    ["📈 Продажи", "🚚 Логистика"]
)

# ==========================================
# СТРАНИЦА 1: ПРОДАЖИ (ваш существующий код)
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
    
    st.divider()
    
    # Помесячная разбивка
    st.subheader(f"📅 ПОМЕСЯЧНАЯ РАЗБИВКА ЗА {selected_year} ГОД")
    
    monthly = df_year.groupby('Период.Месяц').agg({
        'Выручка': 'sum',
        'Валовая_прибыль': 'sum',
        'Количество': 'sum'
    }).reset_index()
    monthly['Название'] = monthly['Период.Месяц'].map(month_names)
    monthly['Рентабельность'] = (monthly['Валовая_прибыль'] / monthly['Выручка'] * 100).fillna(0)
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
            render_small_metric("Выручка", format_number(row['Выручка']), " ₽")
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
    
    # Детали за месяц
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
    
    # Графики
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
    
    # Таблица данных
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

# ==========================================
# СТРАНИЦА 2: ЛОГИСТИКА
# ==========================================
elif page == "🚚 Логистика":
    st.title("🚚 Аналитика затрат на логистику")
    
    if logistics_df.empty:
        st.warning("⚠️ Файл 'logistics_data.xlsx' не найден. Пожалуйста, добавьте файл с данными по логистике в папку с приложением.")
        st.info("📌 Ожидаемая структура файла: колонки 'Дата', 'Контрагент', 'Транспортная_компания', 'Услуга', 'Стоимость_без_НДС', 'НДС', 'Итого'")
    else:
        # Фильтры для логистики
        available_years_log = sorted(logistics_df['Год'].dropna().unique())
        selected_year_log = st.sidebar.selectbox("📅 Выберите год", available_years_log)
        
        df_year_log = logistics_df[logistics_df['Год'] == selected_year_log]
        
        available_months_log = sorted(df_year_log['Месяц_цифра'].dropna().unique())
        available_months_display_log = [month_names[m] for m in available_months_log]
        
        selected_month_display_log = st.sidebar.selectbox("Выберите месяц", available_months_display_log)
        selected_month_num_log = available_months_log[available_months_display_log.index(selected_month_display_log)]
        
        all_carriers = sorted(df_year_log['Транспортная_компания'].dropna().unique())
        selected_carriers = st.sidebar.multiselect(
            "🚛 Транспортные компании",
            all_carriers,
            default=all_carriers[:3] if len(all_carriers) > 3 else all_carriers
        )
        
        df_filtered_log = df_year_log[
            (df_year_log['Месяц_цифра'] == selected_month_num_log) & 
            (df_year_log['Транспортная_компания'].isin(selected_carriers))
        ]
        
        # Годовые метрики логистики
        st.divider()
        st.subheader(f"📦 ИТОГИ ЛОГИСТИКИ ЗА {selected_year_log} ГОД")
        
        year_cost = df_year_log['Итого'].sum()
        year_cost_no_nds = df_year_log['Стоимость_без_НДС'].sum()
        year_nds = df_year_log['НДС'].sum()
        year_operations = len(df_year_log)
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("💰 Общие затраты", f"{format_number(year_cost)} ₽")
        with c2:
            st.metric("📉 Без НДС", f"{format_number(year_cost_no_nds)} ₽")
        with c3:
            st.metric("🧾 НДС", f"{format_number(year_nds)} ₽")
        with c4:
            st.metric("📋 Кол-во операций", f"{format_number(year_operations)}")
        
        st.divider()
        
        # Помесячная разбивка логистики
        st.subheader(f"📅 ПОМЕСЯЧНАЯ РАЗБИВКА ЗА {selected_year_log} ГОД")
        
        monthly_log = df_year_log.groupby('Месяц_цифра').agg({
            'Итого': 'sum',
            'Стоимость_без_НДС': 'sum',
            'НДС': 'sum'
        }).reset_index()
        monthly_log['Название'] = monthly_log['Месяц_цифра'].map(month_names)
        monthly_log = monthly_log.sort_values('Месяц_цифра')
        
        for _, row in monthly_log.iterrows():
            cols = st.columns([1.5, 1, 1, 1])
            with cols[0]:
                st.markdown(f"**{row['Название']}**")
            with cols[1]:
                st.metric("Итого", f"{format_number(row['Итого'])} ₽", label_visibility="collapsed")
            with cols[2]:
                st.metric("Без НДС", f"{format_number(row['Стоимость_без_НДС'])} ₽", label_visibility="collapsed")
            with cols[3]:
                st.metric("НДС", f"{format_number(row['НДС'])} ₽", label_visibility="collapsed")
        
        st.divider()
        
        # Графики по логистике
        col1, col2 = st.columns(2)
        
        with col1:
            # Топ транспортных компаний по затратам
            top_carriers = df_year_log.groupby('Транспортная_компания')['Итого'].sum().nlargest(10).reset_index()
            if not top_carriers.empty:
                fig = px.bar(top_carriers, x='Итого', y='Транспортная_компания', orientation='h',
                             title='Топ транспортных компаний по затратам',
                             color='Итого', color_continuous_scale='Reds')
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Динамика затрат по месяцам
            if not monthly_log.empty:
                fig2 = px.line(monthly_log, x='Название', y='Итого', 
                               markers=True, title='Динамика затрат на логистику по месяцам')
                st.plotly_chart(fig2, use_container_width=True)
        
        # Детали за выбранный месяц
        st.subheader(f"📊 ДЕТАЛИ ЗА {selected_month_display_log} {selected_year_log}")
        
        mm1, mm2, mm3, mm4 = st.columns(4)
        with mm1:
            st.metric("💰 Затраты", f"{format_number(df_filtered_log['Итого'].sum())} ₽")
        with mm2:
            st.metric("📊 Без НДС", f"{format_number(df_filtered_log['Стоимость_без_НДС'].sum())} ₽")
        with mm3:
            st.metric("🧾 НДС", f"{format_number(df_filtered_log['НДС'].sum())} ₽")
        with mm4:
            st.metric("📋 Операций", f"{len(df_filtered_log)}")
        
        # Таблица данных по логистике
        st.subheader("📋 Детальные данные по логистике")
        show_cols_log = ['Дата', 'Контрагент', 'Транспортная_компания', 'Услуга', 'Стоимость_без_НДС', 'НДС', 'Итого']
        show_cols_log = [c for c in show_cols_log if c in df_filtered_log.columns]
        
        if not df_filtered_log.empty:
            st.dataframe(df_filtered_log[show_cols_log].head(100), use_container_width=True)
            csv_log = df_filtered_log[show_cols_log].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
            st.download_button("📥 Скачать CSV (логистика)", csv_log, f"logistics_{selected_year_log}_{selected_month_num_log}.csv", "text/csv")
        else:
            st.warning("Нет данных за выбранный период")
        
        st.caption(f"📅 {selected_month_display_log} {selected_year_log} | Записей: {len(df_filtered_log)}")
