import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os

# ==========================================
# ФУНКЦИИ ДЛЯ ФОРМАТИРОВАНИЯ ЧИСЕЛ
# ==========================================
def format_number(value):
    try:
        if value is None or pd.isna(value):
            return "0"
        if hasattr(value, 'iloc'):
            value = value.iloc[0] if len(value) > 0 else 0
        num_val = float(value)
        if np.isnan(num_val) or np.isinf(num_val):
            return "0"
        return f"{int(num_val):,}".replace(",", " ")
    except (ValueError, TypeError, OverflowError):
        return "0"

def format_float(value, decimals=1):
    try:
        if value is None or pd.isna(value):
            return "0"
        num_val = float(value)
        if np.isnan(num_val) or np.isinf(num_val):
            return "0"
        formatted = f"{num_val:.{decimals}f}".replace(".", ",")
        return formatted
    except (ValueError, TypeError, OverflowError):
        return "0"

# ==========================================
# 1. ЗАГРУЗКА ДАННЫХ
# ==========================================
@st.cache_data
def load_nomenclature():
    try:
        df = pd.read_excel('nomenclature.xlsx')
        
        df = df.rename(columns={
            'Код': 'Код',
            'Артикул': 'Артикул',
            'Наименование': 'Наименование',
            'Наименование полное': 'Наименование_полное',
            'Категория': 'Категория',
            'Вес': 'Вес_кг',
            'Свободно': 'Остаток',
            'Тип': 'Тип'
        })
        
        if df.empty:
            st.warning("Файл номенклатуры пуст")
            return pd.DataFrame()
        
        return df
    except Exception as e:
        st.error(f"Ошибка загрузки номенклатуры: {e}")
        return pd.DataFrame()

@st.cache_data
def load_sales_data():
    df = pd.read_excel('sales_data.xlsx')
    
    df['Дата'] = pd.to_datetime(df['Документ.Дата'], dayfirst=True, errors='coerce')
    df['Месяц'] = df['Дата'].dt.strftime('%Y-%m')
    df['Год'] = df['Дата'].dt.year
    df['Месяц_цифра'] = df['Дата'].dt.month
    
    df['Себестоимость_без_НДС'] = pd.to_numeric(df['Себестоимость без НДС'], errors='coerce').fillna(0)
    
    df['Валовая_прибыль'] = df['Сумма без НДС'] - df['Себестоимость_без_НДС']
    df['Рентабельность_%'] = (df['Валовая_прибыль'] / df['Сумма без НДС'] * 100).fillna(0)
    
    df = df.rename(columns={
        'Сумма без НДС': 'Выручка_без_НДС',
        'Себестоимость_без_НДС': 'Себестоимость',
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
        
        if len(df.columns) >= 17:
            mask = pd.Series([True] * len(df))
            for i in range(8, 17):
                mask = mask & df.iloc[:, i].notna()
            df = df[mask].copy()
        
        if df.empty:
            return pd.DataFrame()
        
        df['Сумма_PLM_до_PЦ'] = pd.to_numeric(df.iloc[:, 0], errors='coerce').fillna(0)
        df['Сумма_КЗ_до_PLM'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
        df['Кол_во_паллет'] = pd.to_numeric(df.iloc[:, 7], errors='coerce').fillna(0)
        
        if len(df.columns) > 8:
            df['Город'] = df.iloc[:, 8]
        if len(df.columns) > 9:
            df['Дата_отгрузки'] = pd.to_datetime(df.iloc[:, 9], errors='coerce')
        if len(df.columns) > 10:
            df['Дата_заказа'] = pd.to_datetime(df.iloc[:, 10], errors='coerce')
        
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

# Загрузка всех данных
sales_df = load_sales_data()
logistics_df = load_logistics_data()
production_df = load_production_data()
logistics_update_df = load_logistics_update_data()
nomenclature_df = load_nomenclature()

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

if not nomenclature_df.empty:
    st.sidebar.success(f"✅ Номенклатура: {len(nomenclature_df)} позиций")
else:
    st.sidebar.warning("⚠️ Номенклатура не загружена")

page = st.sidebar.radio(
    "Выберите раздел",
    ["📈 Продажи", "🚚 Логистика", "📊 Анализ себестоимости", "🏭 Формирование себестоимости ПФ", 
     "🚚 Логистика Update", "🏭 Аналитика производства", "📋 Справочник номенклатуры"]
)

st.sidebar.divider()
if st.sidebar.button("🔄 Принудительно обновить данные", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

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
        try:
            if hasattr(value, 'iloc'):
                val = value.iloc[0] if len(value) > 0 else 0
            else:
                val = value
            
            if pd.isna(val):
                val = 0
            
            st.markdown(
                f"""
                <div style='
                    background-color: #F0F2F6;
                    border-radius: 10px;
                    padding: 10px;
                    text-align: center;
                '>
                    <div style='font-size: 14px; color: #666; margin-bottom: 5px;'>{label}</div>
                    <div style='font-size: 20px; font-weight: bold; color: #1f1f1f;'>{format_number(val)}{suffix}</div>
                </div>
                """,
                unsafe_allow_html=True
            )
        except Exception:
            st.markdown(
                f"""
                <div style='
                    background-color: #F0F2F6;
                    border-radius: 10px;
                    padding: 10px;
                    text-align: center;
                '>
                    <div style='font-size: 14px; color: #666; margin-bottom: 5px;'>{label}</div>
                    <div style='font-size: 20px; font-weight: bold; color: #1f1f1f;'>0{suffix}</div>
                </div>
                """,
                unsafe_allow_html=True
            )
    
    for _, row in monthly.iterrows():
        cols = st.columns([1.5, 1, 1, 1, 1])
        with cols[0]:
            st.markdown(f"<div style='font-weight: bold; font-size: 16px; padding-top: 12px;'>{row['Название']}</div>", unsafe_allow_html=True)
        with cols[1]:
            render_small_metric("Выручка", row['Выручка_без_НДС'], " ₽")
        with cols[2]:
            render_small_metric("Прибыль", row['Валовая_прибыль'], " ₽")
        with cols[3]:
            render_small_metric("Рентабельность", row['Рентабельность'], "%")
        with cols[4]:
            render_small_metric("Кол-во (шт)", row['Количество'])
    
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
    html += '<tr>'
    
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
elif page == "🚚 Логистика":
    st.title("🚚 Аналитика затрат на логистику")
    
    if logistics_df.empty:
        st.warning("⚠️ Файл 'logistics_data.xlsx' не найден или пуст.")
    else:
        st.dataframe(logistics_df.head(100), use_container_width=True)

# ==========================================
# СТРАНИЦА 3: АНАЛИЗ СЕБЕСТОИМОСТИ
# ==========================================
elif page == "📊 Анализ себестоимости":
    st.title("📊 Анализ себестоимости продукции")
    
    df_cost = sales_df.copy()
    df_cost['Себестоимость_единицы'] = df_cost['Себестоимость'] / df_cost['Количество']
    df_cost['Себестоимость_единицы'] = df_cost['Себестоимость_единицы'].replace([float('inf'), -float('inf')], 0).fillna(0)
    
    st.dataframe(df_cost[['Номенклатура', 'Себестоимость_единицы', 'Дата']].head(100), use_container_width=True)

# ==========================================
# СТРАНИЦА 4: ФОРМИРОВАНИЕ СЕБЕСТОИМОСТИ ПФ
# ==========================================
elif page == "🏭 Формирование себестоимости ПФ":
    st.title("🏭 Формирование себестоимости полуфабриката")
    
    if production_df.empty:
        st.warning("⚠️ Файл 'production_data.xlsx' не найден")
    else:
        batches = production_df[production_df['Тип'] == 'Партия'].copy()
        st.dataframe(batches.head(100), use_container_width=True)

# ==========================================
# СТРАНИЦА 5: ЛОГИСТИКА UPDATE
# ==========================================
elif page == "🚚 Логистика Update":
    st.title("🚚 Аналитика логистики (обновленная)")
    
    if logistics_update_df.empty:
        st.warning("⚠️ Файл 'BI logisticks.xlsx' не найден")
    else:
        st.dataframe(logistics_update_df.head(100), use_container_width=True)

# ==========================================
# СТРАНИЦА 6: АНАЛИТИКА ПРОИЗВОДСТВА
# ==========================================
elif page == "🏭 Аналитика производства":
    st.title("🏭 Аналитика производства")
    st.markdown("### Статистика выполнения производственных планов")
    
    planning_df = pd.DataFrame()
    try:
        if os.path.exists('planning_data.xlsx'):
            planning_df = pd.read_excel('planning_data.xlsx')
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
    
    if planning_df.empty:
        st.info("📌 Нет данных о производственных планах.")
        st.caption("Данные загружаются из приложения 'Планирование и производство' (файл planning_data.xlsx)")
    else:
        if 'Дата' in planning_df.columns:
            planning_df['Дата'] = pd.to_datetime(planning_df['Дата'], errors='coerce')
        if 'Запланировано' in planning_df.columns:
            planning_df['Запланировано'] = planning_df['Запланировано'].astype(bool) if planning_df['Запланировано'].dtype == 'object' else planning_df['Запланировано']
        if 'Выполнено' in planning_df.columns:
            planning_df['Выполнено'] = planning_df['Выполнено'].astype(bool) if planning_df['Выполнено'].dtype == 'object' else planning_df['Выполнено']
        
        total_planned = planning_df[planning_df['Запланировано'] == True].shape[0] if 'Запланировано' in planning_df.columns else 0
        total_completed = planning_df[planning_df['Выполнено'] == True].shape[0] if 'Выполнено' in planning_df.columns else 0
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📦 Всего запланировано", format_number(total_planned))
        with col2:
            st.metric("✅ Выполнено", format_number(total_completed))
        with col3:
            percent = (total_completed / total_planned * 100) if total_planned > 0 else 0
            st.metric("📊 Процент выполнения", f"{format_float(percent, 1)}%")
        with col4:
            roasters = planning_df['Ростер'].nunique() if 'Ростер' in planning_df.columns else 0
            st.metric("🔥 Задействовано ростеров", format_number(roasters))
        
        if 'Ростер' in planning_df.columns and 'Запланировано' in planning_df.columns:
            st.divider()
            st.subheader("📊 Статистика по ростерам")
            
            planned_mask = planning_df['Запланировано'] == True
            roaster_stats = planning_df[planned_mask].groupby('Ростер').agg({
                'Выполнено': lambda x: x.sum() if 'Выполнено' in planning_df.columns else 0,
                'Слот': 'count'
            }).reset_index()
            roaster_stats.columns = ['Ростер', 'Выполнено', 'Запланировано']
            roaster_stats['Процент'] = (roaster_stats['Выполнено'] / roaster_stats['Запланировано'] * 100).fillna(0)
            roaster_stats['Процент'] = roaster_stats['Процент'].apply(lambda x: f"{format_float(x, 1)}%")
            
            st.dataframe(roaster_stats, use_container_width=True, hide_index=True)
        
        if 'Дата' in planning_df.columns and 'Запланировано' in planning_df.columns:
            st.divider()
            st.subheader("📈 Динамика выполнения плана")
            
            planned_mask = planning_df['Запланировано'] == True
            daily_stats = planning_df[planned_mask].groupby('Дата').agg({
                'Выполнено': lambda x: x.sum() if 'Выполнено' in planning_df.columns else 0,
                'Слот': 'count'
            }).reset_index()
            daily_stats.columns = ['Дата', 'Выполнено', 'Запланировано']
            daily_stats = daily_stats.sort_values('Дата')
            
            if not daily_stats.empty:
                fig = go.Figure()
                fig.add_trace(go.Bar(name='Запланировано', x=daily_stats['Дата'], y=daily_stats['Запланировано'], marker_color='#2E86AB'))
                fig.add_trace(go.Bar(name='Выполнено', x=daily_stats['Дата'], y=daily_stats['Выполнено'], marker_color='#52B788'))
                fig.update_layout(title='Ежедневная статистика', barmode='group', xaxis_title='Дата', yaxis_title='Количество батчей', height=450)
                st.plotly_chart(fig, use_container_width=True)

# ==========================================
# СТРАНИЦА 7: СПРАВОЧНИК НОМЕНКЛАТУРЫ
# ==========================================
elif page == "📋 Справочник номенклатуры":
    st.title("📋 Справочник номенклатуры")
    st.markdown("### Полный перечень товаров и материалов")
    
    if nomenclature_df.empty:
        st.warning("⚠️ Файл 'nomenclature.xlsx' не найден или не удалось загрузить данные.")
        st.info("📌 Пожалуйста, добавьте файл с номенклатурой в папку с приложением.")
        
        import os
        with st.expander("🔧 Диагностика"):
            st.write("Файлы в текущей папке:")
            for f in os.listdir('.'):
                if f.endswith('.xlsx'):
                    st.write(f"  - {f}")
    else:
        st.subheader("🔍 Фильтры")
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        
        with col_f1:
            categories = ['Все'] + sorted(nomenclature_df['Категория'].dropna().unique().tolist())
            selected_category = st.selectbox("Категория", categories)
        
        with col_f2:
            if 'Тип' in nomenclature_df.columns:
                types = ['Все'] + sorted(nomenclature_df['Тип'].dropna().unique().tolist())
                selected_type = st.selectbox("Тип", types)
            else:
                selected_type = 'Все'
        
        with col_f3:
            search = st.text_input("🔎 Поиск по наименованию", placeholder="Введите название...")
        
        with col_f4:
            hide_zero_stock = st.checkbox("📦 Скрыть позиции с нулевыми остатками", value=True)
        
        filtered_df = nomenclature_df.copy()
        
        if selected_category != 'Все':
            filtered_df = filtered_df[filtered_df['Категория'] == selected_category]
        if selected_type != 'Все' and 'Тип' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Тип'] == selected_type]
        if search:
            filtered_df = filtered_df[filtered_df['Наименование'].str.contains(search, case=False, na=False)]
        if hide_zero_stock and 'Остаток' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Остаток'] > 0]
        
        st.divider()
        
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        with col_s1:
            st.metric("📦 Всего позиций", format_number(len(filtered_df)))
        with col_s2:
            st.metric("📂 Категорий", format_number(filtered_df['Категория'].nunique()))
        with col_s3:
            if 'Остаток' in filtered_df.columns:
                total_stock = filtered_df['Остаток'].sum()
                st.metric("📊 Общий остаток", format_number(total_stock))
            else:
                st.metric("📊 Общий остаток", "нет данных")
        with col_s4:
            if 'Вес_кг' in filtered_df.columns and 'Остаток' in filtered_df.columns:
                total_weight = (filtered_df['Вес_кг'] * filtered_df['Остаток']).sum()
                st.metric("⚖️ Общий вес (кг)", format_number(total_weight))
            else:
                st.metric("⚖️ Общий вес (кг)", "нет данных")
        
        st.divider()
        
        display_cols = ['Код', 'Артикул', 'Наименование', 'Категория', 'Вес_кг', 'Остаток']
        if 'Тип' in filtered_df.columns:
            display_cols.append('Тип')
        
        display_cols = [c for c in display_cols if c in filtered_df.columns]
        
        df_display = filtered_df[display_cols].copy()
        
        if 'Вес_кг' in df_display.columns:
            df_display['Вес_кг'] = df_display['Вес_кг'].apply(lambda x: format_float(x, 2) if pd.notna(x) else "0")
        if 'Остаток' in df_display.columns:
            df_display['Остаток'] = df_display['Остаток'].apply(lambda x: format_number(x) if pd.notna(x) else "0")
        
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        csv = filtered_df[display_cols].to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
        st.download_button("📥 Скачать справочник (CSV)", csv, "nomenclature_export.csv", "text/csv")
