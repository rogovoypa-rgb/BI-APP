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

@st.cache_data
def load_production_data():
    try:
        df_raw = pd.read_excel('production_data.xlsx', header=None)
        
        records = []
        
        for idx, row in df_raw.iterrows():
            if row[0] == 'Номенклатура':
                continue
            
            col0 = row[0] if pd.notna(row[0]) else None
            col1 = row[1] if pd.notna(row[1]) else None
            col2 = row[2] if pd.notna(row[2]) else None
            col3 = row[3] if pd.notna(row[3]) else None
            col4 = row[4] if pd.notna(row[4]) else None
            col5 = row[5] if pd.notna(row[5]) else None
            col6 = row[6] if pd.notna(row[6]) else None
            col7 = row[7] if pd.notna(row[7]) else None
            
            if col3 is not None and col4 is not None and col6 is not None and col7 is not None:
                try:
                    batch_num = str(col3).strip()
                    qty = float(str(col4).replace(',', '.'))
                    cost = float(str(col7).replace(',', '.'))
                    
                    if batch_num.isdigit() and len(batch_num) == 6:
                        records.append({
                            'Тип': 'Партия',
                            'Партия': batch_num,
                            'Количество_выпущено': qty,
                            'Себестоимость_единицы': cost,
                            'Сырье': None,
                            'Количество_сырья': None,
                            'Цена_сырья': None,
                            'Сумма_сырья': None,
                            'Себестоимость_на_единицу_продукции': cost
                        })
                except (ValueError, TypeError):
                    pass
            
            if col0 is not None and col3 is not None and col5 is not None and col6 is not None and col7 is not None:
                try:
                    if col0 == 'П/Ф Дрип Гватемала Декаф 1шт.' or col0 == 'Дрип машина (Производство)':
                        continue
                    
                    batch_id = str(col3).strip()
                    if batch_id.isdigit() and len(batch_id) == 6:
                        qty_raw = float(str(col5).replace(',', '.'))
                        sum_raw = float(str(col6).replace(',', '.'))
                        
                        records.append({
                            'Тип': 'Сырье',
                            'Партия': batch_id,
                            'Количество_выпущено': None,
                            'Себестоимость_единицы': None,
                            'Сырье': col0,
                            'Количество_сырья': qty_raw,
                            'Цена_сырья': sum_raw / qty_raw if qty_raw > 0 else 0,
                            'Сумма_сырья': sum_raw,
                            'Себестоимость_на_единицу_продукции': None
                        })
                except (ValueError, TypeError):
                    pass
        
        df_result = pd.DataFrame(records)
        
        if not df_result.empty:
            for idx, row in df_result[df_result['Тип'] == 'Сырье'].iterrows():
                batch = row['Партия']
                batch_row = df_result[(df_result['Тип'] == 'Партия') & (df_result['Партия'] == batch)]
                if not batch_row.empty:
                    qty = batch_row.iloc[0]['Количество_выпущено']
                    if qty and qty > 0:
                        df_result.at[idx, 'Себестоимость_на_единицу_продукции'] = row['Сумма_сырья'] / qty
                    else:
                        df_result.at[idx, 'Себестоимость_на_единицу_продукции'] = row['Сумма_сырья']
                else:
                    df_result.at[idx, 'Себестоимость_на_единицу_продукции'] = row['Сумма_сырья']
        
        return df_result
    except Exception as e:
        st.error(f"Ошибка загрузки производственных данных: {e}")
        return pd.DataFrame()

@st.cache_data
def load_logistics_update_data():
    try:
        df = pd.read_excel('BI logisticks.xlsx', header=0)
        
        # Оставляем только строки, где все столбцы I-Q (индексы 8-16) заполнены
        if len(df.columns) >= 17:
            mask = pd.Series([True] * len(df))
            for i in range(8, 17):
                mask = mask & df.iloc[:, i].notna()
            df = df[mask].copy()
        
        if df.empty:
            return pd.DataFrame()
        
        # Определяем нужные столбцы по индексам
        # A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12, N=13, O=14, P=15, Q=16
        df['Сумма_PLM_до_PЦ'] = pd.to_numeric(df.iloc[:, 0], errors='coerce').fillna(0)   # A
        df['Сумма_КЗ_до_PLM'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)   # D
        df['Кол_во_паллет'] = pd.to_numeric(df.iloc[:, 7], errors='coerce').fillna(0)      # H
        
        # Город и даты (индексы 8, 9, 10)
        if len(df.columns) > 8:
            df['Город'] = df.iloc[:, 8]
        if len(df.columns) > 9:
            df['Дата_отгрузки'] = pd.to_datetime(df.iloc[:, 9], errors='coerce')
        if len(df.columns) > 10:
            df['Дата_заказа'] = pd.to_datetime(df.iloc[:, 10], errors='coerce')
        
        # Используем дату заказа или отгрузки
        if 'Дата_заказа' in df.columns and df['Дата_заказа'].notna().any():
            df['Дата'] = df['Дата_заказа']
        elif 'Дата_отгрузки' in df.columns:
            df['Дата'] = df['Дата_отгрузки']
        else:
            df['Дата'] = pd.Timestamp.now()
        
        df['Год'] = df['Дата'].dt.year
        df['Месяц'] = df['Дата'].dt.month
        
        return df
    except FileNotFoundError:
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Ошибка загрузки логистики update: {e}")
        return pd.DataFrame()

sales_df = load_sales_data()
logistics_df = load_logistics_data()
production_df = load_production_data()
logistics_update_df = load_logistics_update_data()

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
    ["📈 Продажи", "🚚 Логистика", "📊 Анализ себестоимости", "🏭 Формирование себестоимости ПФ", "🚚 Логистика Update"]
)

# ==========================================
# СТРАНИЦА 1: ПРОДАЖИ
# ==========================================
if page == "📈 Продажи":
    st.title("📊 BI Портал аналитики продаж")
    
    available_years = sorted(sales_df['Год'].dropna().unique())
    available_years_int = [int(y) for y in available_years]
    if len(available_years_int) == 0:
        available_years_int = [2024]
    
    st.divider()
    
    col_filter_year = st.columns([1])[0]
    with col_filter_year:
        selected_year = st.selectbox("📅 Выберите год", available_years_int)
    
    df_year = sales_df[sales_df['Год'] == selected_year]
    available_months_num = sorted(df_year['Период.Месяц'].unique())
    available_months_display = [month_names[m] for m in available_months_num]
    
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
    
    total_top5 = cust_rev[cust_rev['Контрагент'].isin(top5)]['Выручка_без_НДС'].sum()
    st.caption(f"📊 Топ-5: {format_number(total_top5)} ₽ ({format_float(total_top5/year_revenue*100,1)}% от общей выручки без НДС)")
    st.divider()
    
    st.subheader("🔽 ВЫБЕРИТЕ ПАРАМЕТРЫ ДЛЯ ДЕТАЛИЗАЦИИ")
    
    col_filter_month, col_filter_customer = st.columns(2)
    
    with col_filter_month:
        selected_month_display = st.selectbox("📅 Выберите месяц", available_months_display)
        selected_month_num = available_months_num[available_months_display.index(selected_month_display)]
    
    with col_filter_customer:
        all_customers = sorted(df_year['Контрагент'].dropna().unique())
        selected_customers = st.multiselect(
            "🏢 Выберите контрагентов",
            all_customers,
            default=all_customers[:5] if len(all_customers) > 5 else all_customers
        )
    
    df_filtered = df_year[(df_year['Период.Месяц'] == selected_month_num) & (df_year['Контрагент'].isin(selected_customers))]
    
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
# (Оставлен без изменений, но для краткости пропущен в этом сообщении)
# Полный код можно восстановить из предыдущих версий

# ==========================================
# СТРАНИЦА 5: ЛОГИСТИКА UPDATE
# ==========================================
elif page == "🚚 Логистика Update":
    st.title("🚚 Аналитика логистики (обновленная)")
    st.markdown("### Данные по доставке от КЗ до PLM и от PLM до РЦ")
    
    if logistics_update_df.empty:
        st.warning("⚠️ Файл 'BI logisticks.xlsx' не найден или не удалось загрузить данные.")
    else:
        # ФИЛЬТРЫ
        st.divider()
        
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        
        with col_filter1:
            available_years = sorted(logistics_update_df['Год'].dropna().unique())
            if len(available_years) == 0:
                available_years = [2025]
            selected_year = st.selectbox("📅 Выберите год", available_years, key="lu_year")
        
        with col_filter2:
            df_year = logistics_update_df[logistics_update_df['Год'] == selected_year]
            available_months = sorted(df_year['Месяц'].dropna().unique())
            available_months_display = [month_names[m] for m in available_months]
            selected_month_display = st.selectbox("📅 Выберите месяц", available_months_display, key="lu_month")
            selected_month = available_months[available_months_display.index(selected_month_display)]
        
        with col_filter3:
            df_month = df_year[df_year['Месяц'] == selected_month]
            if 'Город' in df_month.columns:
                all_cities = sorted(df_month['Город'].dropna().unique())
                selected_cities = st.multiselect(
                    "🏙️ Выберите города",
                    all_cities,
                    default=all_cities[:5] if len(all_cities) > 5 else all_cities,
                    key="lu_cities"
                )
            else:
                selected_cities = []
        
        mask = (logistics_update_df['Год'] == selected_year) & (logistics_update_df['Месяц'] == selected_month)
        if selected_cities:
            mask = mask & (logistics_update_df['Город'].isin(selected_cities))
        df_filtered = logistics_update_df[mask]
        
        # МЕТРИКИ
        st.divider()
        st.subheader(f"📊 ИТОГИ ЗА {selected_month_display} {selected_year}")
        
        total_plm = df_filtered['Сумма_PLM_до_PЦ'].sum()
        total_kz = df_filtered['Сумма_КЗ_до_PLM'].sum()
        total_delivery = total_plm + total_kz
        total_pallets = df_filtered['Кол_во_паллет'].sum()
        total_orders = len(df_filtered)
        
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.metric("🚛 КЗ→PLM", f"{format_number(total_kz)} ₽")
        with c2:
            st.metric("📦 PLM→РЦ", f"{format_number(total_plm)} ₽")
        with c3:
            st.metric("💰 Итого", f"{format_number(total_delivery)} ₽")
        with c4:
            st.metric("📦 Паллет", f"{format_number(total_pallets)}")
        with c5:
            st.metric("📋 Заказов", f"{format_number(total_orders)}")
        
        if total_pallets > 0:
            st.metric("📊 Средняя стоимость паллеты", f"{format_number(total_delivery / total_pallets)} ₽")
        
        # ТАБЛИЦА
        st.divider()
        st.subheader("📋 Детальные данные")
        
        if not df_filtered.empty:
            display_cols = ['Дата', 'Город', 'Кол_во_паллет', 'Сумма_КЗ_до_PLM', 'Сумма_PLM_до_PЦ']
            display_cols = [c for c in display_cols if c in df_filtered.columns]
            
            df_display = df_filtered[display_cols].copy()
            for col in ['Сумма_КЗ_до_PLM', 'Сумма_PLM_до_PЦ']:
                if col in df_display.columns:
                    df_display[col] = df_display[col].apply(lambda x: f"{format_number(x)} ₽")
            
            st.dataframe(df_display.head(100), use_container_width=True)
            
            csv_data = df_filtered[display_cols].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
            st.download_button("📥 Скачать CSV", csv_data, f"logistics_update_{selected_year}_{selected_month}.csv", "text/csv")
        else:
            st.info("Нет данных за выбранный период")
