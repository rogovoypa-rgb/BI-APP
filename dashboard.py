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
    df['Месяц_цифра'] = df['Дата'].dt.month
    
    df['Валовая_прибыль'] = df['Сумма без НДС'] - df['Себестоимость без НДС']
    df['Рентабельность_%'] = (df['Валовая_прибыль'] / df['Сумма без НДС'] * 100).fillna(0)
    
    df = df.rename(columns={
        'Сумма без НДС': 'Выручка_без_НДС',
        'Себестоимость без НДС': 'Себестоимость',
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
        df = pd.read_excel('logistics_data.xlsx', header=0)
        
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        if 'Строка содержит данные' in df.columns:
            df = df[df['Строка содержит данные'] == 1]
        
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
    ["📈 Продажи", "🚚 Логистика", "📊 Анализ себестоимости"]
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
    html += '</td>'
    
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
    
    total_top5 = cust_rev[cust_rev['Контрагент'].isin(top5)]['Выручка_без_НДС'].sum()
    st.caption(f"📊 Топ-5: {format_number(total_top5)} ₽ ({format_float(total_top5/year_revenue*100,1)}% от общей выручки без НДС)")
    st.divider()
    
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
    
    st.subheader("📋 Детальные данные")
    show_cols = ['Дата', 'Контрагент', 'Номенклатура', 'Выручка_без_НДС', 'Валовая_прибыль', 'Рентабельность_%', 'Количество']
    show_cols = [c for c in show_cols if c in df_filtered.columns]
    if not df_filtered.empty:
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
        
        import os
        with st.expander("🔧 Диагностика"):
            st.write("Файлы в текущей папке:")
            for f in os.listdir('.'):
                if f.endswith('.xlsx'):
                    st.write(f"  - {f}")
    else:
        # ПОДГОТОВКА ДАННЫХ
        df_log = logistics_df.copy()
        
        if 'Дата заказа' in df_log.columns:
            df_log['Дата'] = pd.to_datetime(df_log['Дата заказа'], errors='coerce')
        elif 'Дата отгрузки' in df_log.columns:
            df_log['Дата'] = pd.to_datetime(df_log['Дата отгрузки'], errors='coerce')
        else:
            df_log['Дата'] = pd.Timestamp.now()
        
        df_log['Год'] = df_log['Дата'].dt.year
        df_log['Месяц'] = df_log['Дата'].dt.month
        df_log['Месяц_название'] = df_log['Месяц'].map(month_names)
        
        if 'Дата заказа' in df_log.columns:
            df_log['Номер заказа'] = df_log['Дата заказа'].dt.strftime('%Y%m%d') if pd.api.types.is_datetime64_any_dtype(df_log['Дата заказа']) else df_log['Дата заказа'].astype(str)
        else:
            df_log['Номер заказа'] = df_log['Дата'].dt.strftime('%Y%m%d')
        
        df_log['План'] = pd.to_numeric(df_log['Плановая цена PLM'], errors='coerce')
        df_log['Факт'] = pd.to_numeric(df_log['Фактическая цена PLM'], errors='coerce')
        df_log['Доставка_до_PLM'] = pd.to_numeric(df_log['Доставка до PLM'], errors='coerce')
        df_log['Кол_во_паллет'] = pd.to_numeric(df_log['Кол-во паллет'], errors='coerce')
        
        percent_col = None
        for col in df_log.columns:
            if 'доля' in str(col).lower() and 'объем' in str(col).lower():
                percent_col = col
                break
        
        if percent_col:
            df_log['Процент_доставки'] = pd.to_numeric(df_log[percent_col], errors='coerce')
        else:
            df_log['Процент_доставки'] = 1.0
        
        df_log['Затраты_PLM_на_SKU'] = df_log['Факт'] * df_log['Процент_доставки'].fillna(1)
        df_log['Отклонение'] = df_log['Факт'] - df_log['План']
        
        # СТОИМОСТЬ ТОВАРА В ЗАКАЗЕ (столбец 25)
        if 'Стоимость товара в заказе' in df_log.columns:
            df_log['Стоимость_товара_в_заказе'] = pd.to_numeric(df_log['Стоимость товара в заказе'], errors='coerce')
        else:
            df_log['Стоимость_товара_в_заказе'] = 0
        
        # ВКЛАДКИ НА СТРАНИЦЕ ЛОГИСТИКИ
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Общая аналитика", "📋 Логистика месяц+город+SKU+заказ", "📋 Логистика месяц+город+заказ", "📊 Сводка по городам"])
        
        # ВКЛАДКА 1: ОБЩАЯ АНАЛИТИКА
        with tab1:
            st.header("📦 Общая аналитика логистики")
            
            # Фильтр по году (остаётся вверху)
            col_filter1 = st.columns(1)[0]
            with col_filter1:
                available_years = sorted(df_log['Год'].dropna().unique())
                if len(available_years) == 0:
                    available_years = [2024]
                selected_year_log = st.selectbox("📅 Выберите год", available_years, key="log_year")
            
            # Фильтруем данные по году
            df_year_log = df_log[df_log['Год'] == selected_year_log]
            
            # ГОДОВЫЕ МЕТРИКИ
            st.divider()
            st.subheader(f"📦 ИТОГИ ЛОГИСТИКИ ЗА {selected_year_log} ГОД")
            
            year_plan = df_year_log['План'].sum()
            year_fact = df_year_log['Факт'].sum()
            year_deviation = year_fact - year_plan
            year_deviation_pct = (year_deviation / year_plan * 100) if year_plan > 0 else 0
            year_delivery = df_year_log['Доставка_до_PLM'].sum()
            year_orders = df_year_log['Номер заказа'].nunique()
            year_pallets = df_year_log['Кол_во_паллет'].sum()
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("💰 Фактические затраты PLM", f"{format_number(year_fact)} ₽")
            with c2:
                st.metric("📋 Плановые затраты PLM", f"{format_number(year_plan)} ₽")
            with c3:
                st.metric("📊 Отклонение", f"{format_number(year_deviation)} ₽", delta=f"{format_float(year_deviation_pct, 1)}%")
            with c4:
                st.metric("📦 Кол-во паллет", f"{format_number(year_pallets)}")
            
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
            
            year_sku_costs = df_year_log['Затраты_PLM_на_SKU'].sum()
            st.metric("🎯 Затраты PLM на доставку SKU", f"{format_number(year_sku_costs)} ₽")
            
            st.divider()
            
            # ПОМЕСЯЧНАЯ РАЗБИВКА
            st.subheader(f"📅 ПОМЕСЯЧНАЯ РАЗБИВКА ЗА {selected_year_log} ГОД")
            
            monthly = df_year_log.groupby('Месяц').agg({
                'План': 'sum',
                'Факт': 'sum',
                'Доставка_до_PLM': 'sum',
                'Кол_во_паллет': 'sum',
                'Затраты_PLM_на_SKU': 'sum',
                'Номер заказа': 'nunique'
            }).reset_index()
            monthly['Название'] = monthly['Месяц'].map(month_names)
            monthly['Отклонение'] = monthly['Факт'] - monthly['План']
            monthly = monthly.sort_values('Месяц')
            monthly = monthly.rename(columns={'Номер заказа': 'Кол-во заказов'})
            
            for _, row in monthly.iterrows():
                cols = st.columns([1.5, 1, 1, 1, 1, 1, 1])
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
                    st.metric("Затраты на SKU", f"{format_number(row['Затраты_PLM_на_SKU'])} ₽", label_visibility="collapsed")
                with cols[5]:
                    st.metric("Доставка до PLM", f"{format_number(row['Доставка_до_PLM'])} ₽", label_visibility="collapsed")
                with cols[6]:
                    st.metric("Паллет", format_number(row['Кол_во_паллет']), label_visibility="collapsed")
            
            st.divider()
            
            # ГРАФИКИ (без фильтрации по месяцам и городам)
            col1, col2 = st.columns(2)
            
            with col1:
                if 'Город' in df_year_log.columns:
                    city_costs = df_year_log.groupby('Город')['Факт'].sum().nlargest(10).reset_index()
                    if not city_costs.empty:
                        fig = px.bar(city_costs, x='Факт', y='Город', orientation='h',
                                     title='Топ городов по фактическим затратам PLM',
                                     color='Факт', color_continuous_scale='Reds',
                                     labels={'Факт': 'Затраты (₽)', 'Город': ''})
                        st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                if not monthly.empty:
                    fig2 = go.Figure()
                    fig2.add_trace(go.Scatter(x=monthly['Название'], y=monthly['План'], 
                                              mode='lines+markers', name='План', line=dict(color='blue')))
                    fig2.add_trace(go.Scatter(x=monthly['Название'], y=monthly['Факт'], 
                                              mode='lines+markers', name='Факт', line=dict(color='red')))
                    fig2.update_layout(title='Динамика план vs факт по месяцам',
                                       xaxis_title='Месяц', yaxis_title='Затраты (₽)')
                    st.plotly_chart(fig2, use_container_width=True)
            
            if not monthly.empty:
                st.subheader("📊 Отклонение факта от плана по месяцам")
                fig3 = px.bar(monthly, x='Название', y='Отклонение',
                              title='Отклонение (Факт - План)',
                              color='Отклонение', color_continuous_scale='RdYlGn',
                              labels={'Отклонение': 'Отклонение (₽)', 'Название': 'Месяц'})
                st.plotly_chart(fig3, use_container_width=True)
            
            # ==========================================
            # ДЕТАЛИ ЗА МЕСЯЦ (с выбором месяца и города)
            # ==========================================
            st.divider()
            st.subheader("📊 ДЕТАЛИ ЗА МЕСЯЦ")
            
            # Фильтры для детализации
            col_filter2, col_filter3 = st.columns(2)
            
            with col_filter2:
                available_months = sorted(df_year_log['Месяц'].dropna().unique())
                available_months_display = [month_names[m] for m in available_months]
                selected_month_display = st.selectbox("Выберите месяц", available_months_display, key="detail_month")
                selected_month_num = available_months[available_months_display.index(selected_month_display)]
            
            with col_filter3:
                df_month_log = df_year_log[df_year_log['Месяц'] == selected_month_num]
                if 'Город' in df_month_log.columns:
                    all_cities = sorted(df_month_log['Город'].dropna().unique())
                    selected_cities = st.multiselect(
                        "Выберите города",
                        all_cities,
                        default=all_cities[:5] if len(all_cities) > 5 else all_cities,
                        key="detail_cities"
                    )
                else:
                    selected_cities = []
            
            # Фильтруем данные
            mask = (df_year_log['Месяц'] == selected_month_num)
            if selected_cities:
                mask = mask & (df_year_log['Город'].isin(selected_cities))
            df_detail_month = df_year_log[mask]
            
            # Метрики за выбранный месяц
            d1, d2, d3, d4 = st.columns(4)
            with d1:
                st.metric("💰 Факт PLM", f"{format_number(df_detail_month['Факт'].sum())} ₽")
            with d2:
                st.metric("📋 План PLM", f"{format_number(df_detail_month['План'].sum())} ₽")
            with d3:
                dev = df_detail_month['Факт'].sum() - df_detail_month['План'].sum()
                st.metric("📊 Отклонение", f"{format_number(dev)} ₽")
            with d4:
                st.metric("🎯 Затраты на SKU", f"{format_number(df_detail_month['Затраты_PLM_на_SKU'].sum())} ₽")
            
            # Таблица детализации по заказам за выбранный месяц и город
            if not df_detail_month.empty:
                st.subheader("📋 Детализация по заказам")
                
                order_detail = df_detail_month.groupby('Номер заказа').agg({
                    'Затраты_PLM_на_SKU': 'sum',
                    'Дата заказа': 'first',
                    'Город': 'first'
                }).reset_index()
                order_detail = order_detail.sort_values('Дата заказа', ascending=False)
                
                def format_order_num(order_num):
                    try:
                        s = str(order_num)
                        if len(s) == 8:
                            return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"
                        return s
                    except:
                        return str(order_num)
                
                order_detail['Номер заказа (дата)'] = order_detail['Номер заказа'].apply(format_order_num)
                
                display_orders = order_detail[['Номер заказа (дата)', 'Город', 'Затраты_PLM_на_SKU']].copy()
                display_orders['Затраты_PLM_на_SKU'] = display_orders['Затраты_PLM_на_SKU'].apply(lambda x: f"{format_number(x)} ₽")
                
                st.dataframe(display_orders, use_container_width=True, hide_index=True)
                
                csv_detail = order_detail[['Номер заказа', 'Дата заказа', 'Город', 'Затраты_PLM_на_SKU']].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
                st.download_button("📥 Скачать детализацию по заказам (CSV)", csv_detail, f"logistics_detail_{selected_year_log}_{selected_month_num}.csv", "text/csv")
            else:
                st.info("Нет данных за выбранный период")
        
        # ВКЛАДКА 2: ЛОГИСТИКА МЕСЯЦ+ГОРОД+SKU+ЗАКАЗ
        with tab2:
            st.subheader("📋 Детализация логистики по заказам")
            
            col_filter1, col_filter2, col_filter3, col_filter4 = st.columns(4)
            
            with col_filter1:
                years_available = sorted(df_log['Год'].dropna().unique())
                selected_year_tab2 = st.selectbox("📅 Год", years_available, key="tab2_year")
            
            with col_filter2:
                df_year_tab2 = df_log[df_log['Год'] == selected_year_tab2]
                months_available = sorted(df_year_tab2['Месяц'].dropna().unique())
                months_display_tab2 = [month_names[m] for m in months_available]
                selected_month_display_tab2 = st.selectbox("📅 Месяц", months_display_tab2, key="tab2_month")
                selected_month_tab2 = months_available[months_display_tab2.index(selected_month_display_tab2)]
            
            with col_filter3:
                df_month_tab2 = df_year_tab2[df_year_tab2['Месяц'] == selected_month_tab2]
                cities_available = sorted(df_month_tab2['Город'].dropna().unique())
                selected_city_tab2 = st.selectbox("🏙️ Город", cities_available, key="tab2_city")
            
            with col_filter4:
                df_city_tab2 = df_month_tab2[df_month_tab2['Город'] == selected_city_tab2]
                skus_available = sorted(df_city_tab2['Номенклатура'].dropna().unique())
                selected_sku_tab2 = st.selectbox("📦 SKU (Номенклатура)", skus_available, key="tab2_sku")
            
            df_detail = df_log[
                (df_log['Год'] == selected_year_tab2) &
                (df_log['Месяц'] == selected_month_tab2) &
                (df_log['Город'] == selected_city_tab2) &
                (df_log['Номенклатура'] == selected_sku_tab2)
            ]
            
            if df_detail.empty:
                st.warning("Нет данных за выбранный период по указанному SKU")
            else:
                order_summary = df_detail.groupby('Номер заказа').agg({
                    'Затраты_PLM_на_SKU': 'sum',
                    'Кол_во_паллет': 'sum',
                    'Дата заказа': 'first',
                    'Факт': 'sum',
                    'План': 'sum'
                }).reset_index()
                order_summary = order_summary.sort_values('Дата заказа', ascending=False)
                
                def format_order_num(order_num):
                    try:
                        s = str(order_num)
                        if len(s) == 8:
                            return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"
                        return s
                    except:
                        return str(order_num)
                
                order_summary['Номер заказа (дата)'] = order_summary['Номер заказа'].apply(format_order_num)
                
                st.subheader(f"📊 Результаты для SKU: **{selected_sku_tab2}**")
                st.caption(f"📍 Город: {selected_city_tab2} | 📅 {selected_month_display_tab2} {selected_year_tab2}")
                
                m1, m2, m3 = st.columns(3)
                with m1:
                    total_orders = len(order_summary)
                    st.metric("📋 Кол-во заказов", total_orders)
                with m2:
                    total_cost = order_summary['Затраты_PLM_на_SKU'].sum()
                    st.metric("💰 Общие затраты PLM на доставку SKU", f"{format_number(total_cost)} ₽")
                with m3:
                    avg_cost = total_cost / total_orders if total_orders > 0 else 0
                    st.metric("📊 Средние затраты на заказ", f"{format_number(avg_cost)} ₽")
                
                st.divider()
                st.subheader("📋 Детализация по заказам")
                
                display_df = order_summary[['Номер заказа (дата)', 'Кол_во_паллет', 'Факт', 'План', 'Затраты_PLM_на_SKU']].copy()
                display_df['Затраты_PLM_на_SKU'] = display_df['Затраты_PLM_на_SKU'].apply(lambda x: f"{format_number(x)} ₽")
                display_df['Факт'] = display_df['Факт'].apply(lambda x: f"{format_number(x)} ₽")
                display_df['План'] = display_df['План'].apply(lambda x: f"{format_number(x)} ₽")
                display_df['Кол_во_паллет'] = display_df['Кол_во_паллет'].apply(format_number)
                
                st.dataframe(display_df, use_container_width=True, hide_index=True)
                
                if len(order_summary) > 1:
                    st.subheader("📊 Динамика затрат по заказам")
                    fig_orders = px.bar(order_summary, x='Номер заказа (дата)', y='Затраты_PLM_на_SKU',
                                        title='Затраты PLM на доставку SKU по каждому заказу',
                                        labels={'Затраты_PLM_на_SKU': 'Затраты (₽)', 'Номер заказа (дата)': 'Номер заказа'})
                    st.plotly_chart(fig_orders, use_container_width=True)
                
                csv_detail = order_summary[['Номер заказа', 'Дата заказа', 'Кол_во_паллет', 'Факт', 'План', 'Затраты_PLM_на_SKU']].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
                st.download_button("📥 Скачать детализацию по заказам (CSV)", csv_detail, f"logistics_details_{selected_year_tab2}_{selected_month_tab2}_{selected_city_tab2}_{selected_sku_tab2}.csv", "text/csv")
        
        # ВКЛАДКА 3: ЛОГИСТИКА МЕСЯЦ+ГОРОД+ЗАКАЗ
        with tab3:
            st.subheader("📋 Сводка по заказу")
            st.caption("Сумма всех 'Затраты PLM на доставку SKU' по выбранному заказу")
            
            col_filter1, col_filter2, col_filter3, col_filter4 = st.columns(4)
            
            with col_filter1:
                years_available = sorted(df_log['Год'].dropna().unique())
                selected_year_tab3 = st.selectbox("📅 Год", years_available, key="tab3_year")
            
            with col_filter2:
                df_year_tab3 = df_log[df_log['Год'] == selected_year_tab3]
                months_available = sorted(df_year_tab3['Месяц'].dropna().unique())
                months_display_tab3 = [month_names[m] for m in months_available]
                selected_month_display_tab3 = st.selectbox("📅 Месяц", months_display_tab3, key="tab3_month")
                selected_month_tab3 = months_available[months_display_tab3.index(selected_month_display_tab3)]
            
            with col_filter3:
                df_month_tab3 = df_year_tab3[df_year_tab3['Месяц'] == selected_month_tab3]
                cities_available = sorted(df_month_tab3['Город'].dropna().unique())
                selected_city_tab3 = st.selectbox("🏙️ Город", cities_available, key="tab3_city")
            
            with col_filter4:
                df_city_tab3 = df_month_tab3[df_month_tab3['Город'] == selected_city_tab3]
                orders_available = sorted(df_city_tab3['Номер заказа'].dropna().unique(), reverse=True)
                
                def format_order_num(order_num):
                    try:
                        s = str(order_num)
                        if len(s) == 8:
                            return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"
                        return s
                    except:
                        return str(order_num)
                
                orders_display = [format_order_num(o) for o in orders_available]
                selected_order_display = st.selectbox("📋 Номер заказа (дата)", orders_display, key="tab3_order")
                selected_order_tab3 = orders_available[orders_display.index(selected_order_display)]
            
            df_order = df_log[
                (df_log['Год'] == selected_year_tab3) &
                (df_log['Месяц'] == selected_month_tab3) &
                (df_log['Город'] == selected_city_tab3) &
                (df_log['Номер заказа'] == selected_order_tab3)
            ]
            
            if df_order.empty:
                st.warning("Нет данных за выбранный период")
            else:
                total_sku_cost = df_order['Затраты_PLM_на_SKU'].sum()
                unique_skus = df_order['Номенклатура'].nunique()
                
                st.subheader(f"📊 Сводка по заказу **{selected_order_display}**")
                st.caption(f"📍 Город: {selected_city_tab3} | 📅 {selected_month_display_tab3} {selected_year_tab3}")
                
                m1, m2 = st.columns(2)
                with m1:
                    st.metric("💰 Сумма затрат PLM на доставку SKU", f"{format_number(total_sku_cost)} ₽")
                with m2:
                    st.metric("📦 Количество уникальных SKU", unique_skus)
                
                st.divider()
                st.subheader("📋 Детализация по SKU в заказе")
                
                sku_summary = df_order.groupby('Номенклатура').agg({
                    'Затраты_PLM_на_SKU': 'sum',
                    'Кол_во_паллет': 'sum'
                }).reset_index()
                sku_summary = sku_summary.sort_values('Затраты_PLM_на_SKU', ascending=False)
                
                display_sku = sku_summary.copy()
                display_sku['Затраты_PLM_на_SKU'] = display_sku['Затраты_PLM_на_SKU'].apply(lambda x: f"{format_number(x)} ₽")
                display_sku['Кол_во_паллет'] = display_sku['Кол_во_паллет'].apply(format_number)
                
                st.dataframe(display_sku, use_container_width=True, hide_index=True)
                
                if len(sku_summary) > 1:
                    st.subheader("📊 Затраты PLM на доставку SKU по каждому SKU")
                    fig_sku = px.bar(sku_summary, x='Номенклатура', y='Затраты_PLM_на_SKU',
                                     title=f'Распределение затрат по SKU в заказе {selected_order_display}',
                                     labels={'Затраты_PLM_на_SKU': 'Затраты (₽)', 'Номенклатура': 'SKU'})
                    fig_sku.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig_sku, use_container_width=True)
                
                csv_order = sku_summary.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
                st.download_button("📥 Скачать детализацию по заказу (CSV)", csv_order, f"logistics_order_{selected_year_tab3}_{selected_month_tab3}_{selected_city_tab3}_{selected_order_tab3}.csv", "text/csv")
        
        # ВКЛАДКА 4: СВОДКА ПО ГОРОДАМ
        with tab4:
            st.subheader("📊 Сводка затрат PLM на доставку SKU по городам")
            st.caption("Сумма 'Затраты PLM на доставку SKU' и 'Стоимость товара в заказе' по всем заказам за выбранный месяц")
            
            col_filter1, col_filter2, col_filter3 = st.columns(3)
            
            with col_filter1:
                years_available = sorted(df_log['Год'].dropna().unique())
                selected_year_tab4 = st.selectbox("📅 Год", years_available, key="tab4_year")
            
            with col_filter2:
                df_year_tab4 = df_log[df_log['Год'] == selected_year_tab4]
                months_available = sorted(df_year_tab4['Месяц'].dropna().unique())
                months_display_tab4 = [month_names[m] for m in months_available]
                selected_month_display_tab4 = st.selectbox("📅 Месяц", months_display_tab4, key="tab4_month")
                selected_month_tab4 = months_available[months_display_tab4.index(selected_month_display_tab4)]
            
            with col_filter3:
                df_month_tab4 = df_year_tab4[df_year_tab4['Месяц'] == selected_month_tab4]
                cities_available = sorted(df_month_tab4['Город'].dropna().unique())
                selected_city_tab4 = st.selectbox("🏙️ Город", ["Все города"] + cities_available, key="tab4_city")
            
            mask = (df_log['Год'] == selected_year_tab4) & (df_log['Месяц'] == selected_month_tab4)
            if selected_city_tab4 != "Все города":
                mask = mask & (df_log['Город'] == selected_city_tab4)
            
            df_filtered_tab4 = df_log[mask]
            
            if df_filtered_tab4.empty:
                st.warning("Нет данных за выбранный период")
            else:
                if selected_city_tab4 != "Все города":
                    # Группируем по заказам для конкретного города
                    order_summary = df_filtered_tab4.groupby('Номер заказа').agg({
                        'Затраты_PLM_на_SKU': 'sum',
                        'Стоимость_товара_в_заказе': 'sum',
                        'Дата заказа': 'first'
                    }).reset_index()
                    order_summary = order_summary.sort_values('Дата заказа', ascending=False)
                    
                    def format_order_num(order_num):
                        try:
                            s = str(order_num)
                            if len(s) == 8:
                                return f"{s[6:8]}.{s[4:6]}.{s[0:4]}"
                            return s
                        except:
                            return str(order_num)
                    
                    order_summary['Номер заказа (дата)'] = order_summary['Номер заказа'].apply(format_order_num)
                    
                    total_cost = order_summary['Затраты_PLM_на_SKU'].sum()
                    total_goods = order_summary['Стоимость_товара_в_заказе'].sum()
                    
                    st.subheader(f"📊 Результаты для города: **{selected_city_tab4}**")
                    st.caption(f"📅 {selected_month_display_tab4} {selected_year_tab4}")
                    
                    m1, m2, m3 = st.columns(3)
                    with m1:
                        st.metric("💰 Затраты PLM на доставку SKU", f"{format_number(total_cost)} ₽")
                    with m2:
                        st.metric("📦 Стоимость товара в заказе", f"{format_number(total_goods)} ₽")
                    with m3:
                        st.metric("📋 Количество заказов", len(order_summary))
                    
                    if total_goods > 0:
                        cost_percent = (total_cost / total_goods * 100)
                        st.metric("📊 Доля затрат PLM в стоимости товара", f"{format_float(cost_percent, 1)}%")
                    
                    st.divider()
                    st.subheader("📋 Детализация по заказам")
                    
                    display_orders = order_summary[['Номер заказа (дата)', 'Затраты_PLM_на_SKU', 'Стоимость_товара_в_заказе']].copy()
                    display_orders['Затраты_PLM_на_SKU'] = display_orders['Затраты_PLM_на_SKU'].apply(lambda x: f"{format_number(x)} ₽")
                    display_orders['Стоимость_товара_в_заказе'] = display_orders['Стоимость_товара_в_заказе'].apply(lambda x: f"{format_number(x)} ₽")
                    
                    st.dataframe(display_orders, use_container_width=True, hide_index=True)
                    
                    if len(order_summary) > 1:
                        st.subheader("📊 Динамика затрат по заказам")
                        fig_orders = px.bar(order_summary, x='Номер заказа (дата)', y='Затраты_PLM_на_SKU',
                                            title=f'Затраты PLM на доставку SKU по заказам в городе {selected_city_tab4}',
                                            labels={'Затраты_PLM_на_SKU': 'Затраты (₽)', 'Номер заказа (дата)': 'Номер заказа'})
                        st.plotly_chart(fig_orders, use_container_width=True)
                    
                    csv_data = order_summary[['Номер заказа', 'Дата заказа', 'Затраты_PLM_на_SKU', 'Стоимость_товара_в_заказе']].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
                    st.download_button("📥 Скачать данные по городу (CSV)", csv_data, f"logistics_city_{selected_year_tab4}_{selected_month_tab4}_{selected_city_tab4}.csv", "text/csv")
                
                else:
                    # Сводка по всем городам
                    city_summary = df_filtered_tab4.groupby('Город').agg({
                        'Затраты_PLM_на_SKU': 'sum',
                        'Стоимость_товара_в_заказе': 'sum',
                        'Номер заказа': 'nunique'
                    }).reset_index()
                    city_summary = city_summary.sort_values('Затраты_PLM_на_SKU', ascending=False)
                    city_summary = city_summary.rename(columns={'Номер заказа': 'Кол-во заказов'})
                    
                    total_all_costs = city_summary['Затраты_PLM_на_SKU'].sum()
                    total_all_goods = city_summary['Стоимость_товара_в_заказе'].sum()
                    total_orders_all = city_summary['Кол-во заказов'].sum()
                    
                    st.subheader(f"📊 Сводка по всем городам за {selected_month_display_tab4} {selected_year_tab4}")
                    
                    m1, m2, m3 = st.columns(3)
                    with m1:
                        st.metric("💰 Общие затраты PLM на доставку SKU", f"{format_number(total_all_costs)} ₽")
                    with m2:
                        st.metric("📦 Общая стоимость товара в заказах", f"{format_number(total_all_goods)} ₽")
                    with m3:
                        st.metric("📋 Общее количество заказов", format_number(total_orders_all))
                    
                    if total_all_goods > 0:
                        overall_percent = (total_all_costs / total_all_goods * 100)
                        st.metric("📊 Общая доля затрат PLM в стоимости товара", f"{format_float(overall_percent, 1)}%")
                    
                    st.divider()
                    st.subheader("📋 Детализация по городам")
                    
                    display_cities = city_summary.copy()
                    display_cities['Затраты_PLM_на_SKU'] = display_cities['Затраты_PLM_на_SKU'].apply(lambda x: f"{format_number(x)} ₽")
                    display_cities['Стоимость_товара_в_заказе'] = display_cities['Стоимость_товара_в_заказе'].apply(lambda x: f"{format_number(x)} ₽")
                    display_cities['Кол-во заказов'] = display_cities['Кол-во заказов'].apply(format_number)
                    display_cities['Доля затрат'] = (city_summary['Затраты_PLM_на_SKU'] / city_summary['Стоимость_товара_в_заказе'] * 100).apply(lambda x: f"{format_float(x, 1)}%")
                    
                    st.dataframe(display_cities, use_container_width=True, hide_index=True)
                    
                    if len(city_summary) > 1:
                        st.subheader("📊 Затраты PLM на доставку SKU по городам")
                        fig_cities = px.bar(city_summary, x='Затраты_PLM_на_SKU', y='Город', orientation='h',
                                            title=f'Затраты по городам за {selected_month_display_tab4} {selected_year_tab4}',
                                            color='Затраты_PLM_на_SKU', color_continuous_scale='Blues',
                                            labels={'Затраты_PLM_на_SKU': 'Затраты (₽)', 'Город': ''})
                        st.plotly_chart(fig_cities, use_container_width=True)
                    
                    if len(city_summary) > 1:
                        st.subheader("📊 Стоимость товара в заказах по городам")
                        fig_goods = px.bar(city_summary, x='Стоимость_товара_в_заказе', y='Город', orientation='h',
                                           title=f'Стоимость товара по городам за {selected_month_display_tab4} {selected_year_tab4}',
                                           color='Стоимость_товара_в_заказе', color_continuous_scale='Greens',
                                           labels={'Стоимость_товара_в_заказе': 'Стоимость товара (₽)', 'Город': ''})
                        st.plotly_chart(fig_goods, use_container_width=True)
                    
                    if len(city_summary) > 1:
                        st.subheader("🥧 Доля затрат PLM по городам")
                        fig_pie = px.pie(city_summary, values='Затраты_PLM_на_SKU', names='Город',
                                         title=f'Распределение затрат PLM по городам за {selected_month_display_tab4} {selected_year_tab4}')
                        st.plotly_chart(fig_pie, use_container_width=True)
                    
                    csv_data = city_summary.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
                    st.download_button("📥 Скачать сводку по городам (CSV)", csv_data, f"logistics_cities_summary_{selected_year_tab4}_{selected_month_tab4}.csv", "text/csv")

# ==========================================
# СТРАНИЦА 3: АНАЛИЗ СЕБЕСТОИМОСТИ
# ==========================================
elif page == "📊 Анализ себестоимости":
    st.title("📊 Анализ себестоимости продукции")
    
    st.markdown("""
    ### Динамика себестоимости без НДС по номенклатурам
    Графики сгруппированы по параметру из столбца **M** (Группа/категория).
    """)
    
    # Подготовка данных
    df_cost = sales_df.copy()
    
    # Создаём столбец для группировки (параметр из столбца M)
    # В вашем файле это может быть 'Группа' или 'Подгруппа'
    group_col = None
    for col in df_cost.columns:
        if col in ['Группа', 'Подгруппа', 'Категория']:
            group_col = col
            break
    
    if group_col is None:
        st.warning("⚠️ Не найден столбец для группировки номенклатур (Группа, Подгруппа или Категория)")
        group_col = 'Номенклатура'
    
    # Получаем список уникальных групп
    unique_groups = df_cost[group_col].dropna().unique()
    unique_groups = sorted(unique_groups)
    
    # Создаём временный столбец для сортировки дат
    df_cost['Дата_сортировка'] = pd.to_datetime(df_cost['Дата'], errors='coerce')
    df_cost['Период'] = df_cost['Дата_сортировка'].dt.strftime('%Y-%m')
    
    # Функция для форматирования дат на графиках
    def format_date_axis(fig):
        fig.update_xaxes(
            tickangle=-45,
            tickformat="%b %Y"
        )
        return fig
    
    # Выбор группы через выпадающий список
    selected_group = st.selectbox(
        "📂 Выберите группу номенклатур",
        unique_groups,
        help="Группировка по столбцу 'Группа' (или аналогичному)"
    )
    
    # Фильтруем номенклатуры по выбранной группе
    df_group = df_cost[df_cost[group_col] == selected_group]
    
    if df_group.empty:
        st.warning(f"Нет данных для группы '{selected_group}'")
    else:
        # Получаем список номенклатур в выбранной группе
        nomenclatures = sorted(df_group['Номенклатура'].dropna().unique())
        
        st.subheader(f"📦 Номенклатуры в группе: {selected_group}")
        st.caption(f"Всего номенклатур в группе: {len(nomenclatures)}")
        
        # Для каждой номенклатуры строим отдельный график
        for nomen in nomenclatures:
            df_nomen = df_group[df_group['Номенклатура'] == nomen].copy()
            
            if df_nomen.empty:
                continue
            
            # Группируем по месяцам
            monthly_cost = df_nomen.groupby('Период')['Себестоимость'].mean().reset_index()
            monthly_cost = monthly_cost.sort_values('Период')
            
            if monthly_cost.empty:
                continue
            
            # Создаём график
            fig = go.Figure()
            
            fig.add_trace(go.Scatter(
                x=monthly_cost['Период'],
                y=monthly_cost['Себестоимость'],
                mode='lines+markers',
                name=nomen,
                line=dict(width=2, color='#2E86AB'),
                marker=dict(size=6, color='#1A5276')
            ))
            
            # Добавляем линию среднего значения
            avg_cost = monthly_cost['Себестоимость'].mean()
            fig.add_hline(
                y=avg_cost,
                line_dash="dash",
                line_color="red",
                annotation_text=f"Среднее: {format_number(avg_cost)} ₽",
                annotation_position="top right"
            )
            
            fig.update_layout(
                title=f"📈 {nomen}",
                xaxis_title="Период (месяц/год)",
                yaxis_title="Себестоимость без НДС (₽)",
                hovermode='x unified',
                height=400,
                margin=dict(l=50, r=50, t=60, b=50)
            )
            
            fig.update_xaxes(tickangle=-45)
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Добавляем небольшую статистику под графиком
            min_cost = monthly_cost['Себестоимость'].min()
            max_cost = monthly_cost['Себестоимость'].max()
            last_cost = monthly_cost['Себестоимость'].iloc[-1] if not monthly_cost.empty else 0
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Минимальная себестоимость", f"{format_number(min_cost)} ₽")
            with col2:
                st.metric("Максимальная себестоимость", f"{format_number(max_cost)} ₽")
            with col3:
                st.metric("Последнее значение", f"{format_number(last_cost)} ₽")
            
            st.divider()
    
    # Дополнительный блок: сводная таблица по всем номенклатурам группы
    if not df_group.empty:
        st.subheader(f"📊 Сводная таблица себестоимости по номенклатурам группы '{selected_group}'")
        
        # Сводная таблица: номенклатура → средняя себестоимость
        summary_table = df_group.groupby('Номенклатура').agg({
            'Себестоимость': ['mean', 'min', 'max', 'std']
        }).round(2)
        summary_table.columns = ['Средняя', 'Минимум', 'Максимум', 'Стд. отклонение']
        summary_table = summary_table.sort_values('Средняя', ascending=False)
        
        # Форматируем
        for col in summary_table.columns:
            summary_table[col] = summary_table[col].apply(lambda x: format_number(x))
        
        st.dataframe(summary_table, use_container_width=True)
        
        # Кнопка скачивания
        csv_summary = df_group.groupby('Номенклатура').agg({
            'Себестоимость': ['mean', 'min', 'max', 'std']
        }).round(2)
        csv_summary.columns = ['Средняя', 'Минимум', 'Максимум', 'Стд_отклонение']
        csv_summary = csv_summary.reset_index()
        
        csv_data = csv_summary.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
        st.download_button(
            "📥 Скачать сводную таблицу (CSV)",
            csv_data,
            f"cost_summary_{selected_group}.csv",
            "text/csv"
        )
