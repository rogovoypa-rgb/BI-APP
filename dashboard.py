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
        
        # Проверяем наличие необходимых колонок
        required_cols = ['Код', 'Артикул', 'Наименование', 'Категория']
        for col in required_cols:
            if col not in df.columns:
                # Ищем похожие колонки
                for df_col in df.columns:
                    if col.lower() in str(df_col).lower():
                        df = df.rename(columns={df_col: col})
                        break
        
        # Переименовываем колонки, если они есть
        rename_cols = {
            'Наименование полное': 'Наименование_полное',
            'Вес': 'Вес_кг',
            'Свободно': 'Остаток',
            'Тип': 'Тип'
        }
        
        for old_name, new_name in rename_cols.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        if df.empty:
            st.warning("Файл номенклатуры пуст")
            return pd.DataFrame()
        
        return df
    except Exception as e:
        st.error(f"Ошибка загрузки номенклатуры: {e}")
        return pd.DataFrame()

@st.cache_data
def load_sales_data():
    try:
        df = pd.read_excel('sales_data.xlsx')
        
        # ===== АВТОМАТИЧЕСКОЕ ОПРЕДЕЛЕНИЕ КОЛОНКИ С ДАТОЙ =====
        date_column = None
        
        # Список возможных названий колонок с датой
        possible_date_columns = [
            'Документ.Дата', 
            'Дата документа', 
            'Дата', 
            'Date', 
            'Документ Дата',
            'ДатаДок',
            'Дата заказа'
        ]
        
        # Проверяем наличие колонок с датой
        for col in possible_date_columns:
            if col in df.columns:
                date_column = col
                break
        
        # Если не нашли по точному совпадению, ищем колонки, содержащие слово "дата" или "date"
        if date_column is None:
            for col in df.columns:
                col_lower = str(col).lower()
                if 'дата' in col_lower or 'date' in col_lower:
                    date_column = col
                    break
        
        # Если всё ещё не нашли, ищем колонку с типом datetime
        if date_column is None:
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    date_column = col
                    break
        
        # Если нашли колонку с датой, используем её
        if date_column is not None:
            try:
                df['Дата'] = pd.to_datetime(df[date_column], dayfirst=True, errors='coerce')
            except Exception:
                df['Дата'] = pd.to_datetime(df[date_column], errors='coerce')
        else:
            df['Дата'] = pd.Timestamp.now()
        
        # Создаём временные признаки
        df['Месяц'] = df['Дата'].dt.strftime('%Y-%m')
        df['Год'] = df['Дата'].dt.year
        df['Месяц_цифра'] = df['Дата'].dt.month
        
        # Определяем колонку с себестоимостью
        try:
            if len(df.columns) > 19:
                df['Себестоимость_без_НДС'] = pd.to_numeric(df.iloc[:, 19], errors='coerce').fillna(0)
            else:
                cost_column = None
                for col in df.columns:
                    col_lower = str(col).lower()
                    if 'себестоим' in col_lower or 'cost' in col_lower:
                        cost_column = col
                        break
                
                if cost_column is not None:
                    df['Себестоимость_без_НДС'] = pd.to_numeric(df[cost_column], errors='coerce').fillna(0)
                else:
                    df['Себестоимость_без_НДС'] = 0
        except Exception:
            df['Себестоимость_без_НДС'] = 0
        
        # Расчёт прибыли и рентабельности
        if 'Сумма без НДС' in df.columns:
            df['Валовая_прибыль'] = df['Сумма без НДС'] - df['Себестоимость_без_НДС']
            df['Рентабельность_%'] = (df['Валовая_прибыль'] / df['Сумма без НДС'] * 100).fillna(0)
        else:
            df['Валовая_прибыль'] = 0
            df['Рентабельность_%'] = 0
        
        # Переименовываем колонки
        rename_map = {
            'Сумма без НДС': 'Выручка_без_НДС',
            'Себестоимость без НДС': 'Себестоимость',
            'Контрагент': 'Контрагент',
            'Номенклатура': 'Номенклатура'
        }
        
        for old_name, new_name in rename_map.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        # Добавляем колонку Количество, если её нет
        if 'Количество' not in df.columns:
            for col in df.columns:
                col_lower = str(col).lower()
                if 'количеств' in col_lower or 'кол-во' in col_lower or 'qty' in col_lower:
                    df['Количество'] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    break
            
            if 'Количество' not in df.columns:
                df['Количество'] = 0
        
        # Период.Месяц
        if 'Период.Месяц' in df.columns:
            df['Период.Месяц'] = pd.to_numeric(df['Период.Месяц'], errors='coerce')
            df = df.dropna(subset=['Период.Месяц'])
            df['Период.Месяц'] = df['Период.Месяц'].astype(int)
        else:
            # Создаём Период.Месяц из даты
            df['Период.Месяц'] = df['Дата'].dt.month
        
        return df
    except Exception as e:
        st.error(f"Ошибка загрузки данных продаж: {e}")
        return pd.DataFrame()

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
            
            if col0 is not None and col3 is not None and col5 is not None and col6 is not None:
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

month_names_ru = {
    1: 'Янв', 2: 'Фев', 3: 'Мар', 4: 'Апр', 5: 'Май', 6: 'Июн',
    7: 'Июл', 8: 'Авг', 9: 'Сен', 10: 'Окт', 11: 'Ноя', 12: 'Дек'
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
# СТРАНИЦА 1: ПРОДАЖИ (С АНАЛИЗОМ ТОП-5 ПОКУПАТЕЛЕЙ)
# ==========================================
if page == "📈 Продажи":
    st.title("📊 BI Портал аналитики продаж")
    
    if sales_df.empty:
        st.warning("⚠️ Файл 'sales_data.xlsx' не найден или не удалось загрузить данные.")
        st.info("📌 Пожалуйста, добавьте файл с данными о продажах в папку с приложением.")
    else:
        available_years = sorted(sales_df['Год'].dropna().unique())
        available_years_int = [int(y) for y in available_years]
        if len(available_years_int) == 0:
            available_years_int = [2024]
        
        default_year = 2026 if 2026 in available_years_int else available_years_int[0]
        
        st.divider()
        
        col_filter_year = st.columns([1])[0]
        with col_filter_year:
            selected_year = st.selectbox("📅 Выберите год", available_years_int, index=available_years_int.index(default_year))
        
        df_year = sales_df[sales_df['Год'] == selected_year]
        available_months_num = sorted(df_year['Период.Месяц'].unique())
        available_months_display = [month_names[m] for m in available_months_num]
        
        st.divider()
        st.subheader(f"📈 ИТОГИ ЗА {selected_year} ГОД")
        
        year_revenue = df_year['Выручка_без_НДС'].sum() if 'Выручка_без_НДС' in df_year.columns else 0
        year_profit = df_year['Валовая_прибыль'].sum() if 'Валовая_прибыль' in df_year.columns else 0
        year_margin = (year_profit / year_revenue * 100) if year_revenue > 0 else 0
        year_quantity = df_year['Количество'].sum() if 'Количество' in df_year.columns else 0
        
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
        
        has_cost_column = 'Себестоимость' in df_year.columns
        
        monthly = df_year.groupby('Период.Месяц').agg({
            'Выручка_без_НДС': 'sum',
            'Валовая_прибыль': 'sum',
            'Количество': 'sum'
        }).reset_index()
        
        if has_cost_column:
            monthly_cost = df_year.groupby('Период.Месяц')['Себестоимость'].sum().reset_index()
            monthly = monthly.merge(monthly_cost, on='Период.Месяц')
            monthly['Себестоимость'] = monthly['Себестоимость'].fillna(0)
        else:
            monthly['Себестоимость'] = 0
        
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
            cols = st.columns([1.5, 1, 1, 1, 1, 1])
            with cols[0]:
                st.markdown(f"<div style='font-weight: bold; font-size: 16px; padding-top: 12px;'>{row['Название']}</div>", unsafe_allow_html=True)
            with cols[1]:
                render_small_metric("Выручка", row['Выручка_без_НДС'], " ₽")
            with cols[2]:
                render_small_metric("Прибыль", row['Валовая_прибыль'], " ₽")
            with cols[3]:
                render_small_metric("Себестоимость", row['Себестоимость'], " ₽")
            with cols[4]:
                render_small_metric("Рентабельность", row['Рентабельность'], "%")
            with cols[5]:
                render_small_metric("Кол-во (шт)", row['Количество'])
        
        st.markdown("---")
        
        total_cols = st.columns([1.5, 1, 1, 1, 1, 1])
        total_cost = df_year['Себестоимость'].sum() if has_cost_column and 'Себестоимость' in df_year.columns else 0
        
        with total_cols[0]:
            st.markdown("<div style='font-weight: bold; font-size: 16px;'>📊 ИТОГО</div>", unsafe_allow_html=True)
        with total_cols[1]:
            st.markdown(f"<div style='font-size: 16px;'><b>{format_number(year_revenue)} ₽</b></div>", unsafe_allow_html=True)
        with total_cols[2]:
            st.markdown(f"<div style='font-size: 16px;'><b>{format_number(year_profit)} ₽</b></div>", unsafe_allow_html=True)
        with total_cols[3]:
            st.markdown(f"<div style='font-size: 16px;'><b>{format_number(total_cost)} ₽</b></div>", unsafe_allow_html=True)
        with total_cols[4]:
            st.markdown(f"<div style='font-size: 16px;'><b>{format_float(year_margin, 1)}%</b></div>", unsafe_allow_html=True)
        with total_cols[5]:
            st.markdown(f"<div style='font-size: 16px;'><b>{format_number(year_quantity)}</b></div>", unsafe_allow_html=True)
        
        st.divider()
        
        # ==========================================
        # ТОП-5 КОНТРАГЕНТОВ (СВОДНАЯ ТАБЛИЦА)
        # ==========================================
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
        html += '</table>'
        
        st.markdown(html, unsafe_allow_html=True)
        
        total_top5 = cust_rev[cust_rev['Контрагент'].isin(top5)]['Выручка_без_НДС'].sum()
        st.caption(f"📊 Топ-5: {format_number(total_top5)} ₽ ({format_float(total_top5/year_revenue*100,1)}% от общей выручки без НДС)")
        
        st.divider()
        
        # ==========================================
        # ДЕТАЛЬНЫЙ АНАЛИЗ КАЖДОГО ИЗ ТОП-5 ПОКУПАТЕЛЕЙ
        # ==========================================
        st.subheader(f"📊 ДЕТАЛЬНЫЙ АНАЛИЗ ТОП-5 ПОКУПАТЕЛЕЙ ЗА {selected_year}")
        st.markdown("По каждому контрагенту: выручка, прибыль, себестоимость, рентабельность, количество (шт) по месяцам")
        
        # Функция для отображения таблицы одного контрагента
        def display_customer_analysis(customer_name, df_year, available_months_num, has_cost_column):
            """Отображает детальный анализ по одному контрагенту"""
            
            # Фильтруем данные по контрагенту
            df_customer = df_year[df_year['Контрагент'] == customer_name]
            
            if df_customer.empty:
                st.info(f"Нет данных для контрагента {customer_name}")
                return
            
            # Агрегируем по месяцам
            if has_cost_column:
                customer_monthly = df_customer.groupby('Период.Месяц').agg({
                    'Выручка_без_НДС': 'sum',
                    'Валовая_прибыль': 'sum',
                    'Количество': 'sum',
                    'Себестоимость': 'sum'
                }).reset_index()
                customer_monthly['Себестоимость'] = customer_monthly['Себестоимость'].fillna(0)
            else:
                customer_monthly = df_customer.groupby('Период.Месяц').agg({
                    'Выручка_без_НДС': 'sum',
                    'Валовая_прибыль': 'sum',
                    'Количество': 'sum'
                }).reset_index()
                customer_monthly['Себестоимость'] = 0
            
            # Добавляем названия месяцев
            customer_monthly['Название'] = customer_monthly['Период.Месяц'].map(month_names)
            customer_monthly['Рентабельность'] = (customer_monthly['Валовая_прибыль'] / customer_monthly['Выручка_без_НДС'] * 100).fillna(0)
            customer_monthly = customer_monthly.sort_values('Период.Месяц')
            
            # Заполняем пропущенные месяцы нулями
            full_months = pd.DataFrame({'Период.Месяц': available_months_num})
            customer_monthly = pd.merge(full_months, customer_monthly, on='Период.Месяц', how='left')
            customer_monthly['Название'] = customer_monthly['Период.Месяц'].map(month_names)
            customer_monthly = customer_monthly.fillna(0)
            customer_monthly['Рентабельность'] = (customer_monthly['Валовая_прибыль'] / customer_monthly['Выручка_без_НДС'] * 100).fillna(0)
            
            # Итоговые суммы
            total_revenue = customer_monthly['Выручка_без_НДС'].sum()
            total_profit = customer_monthly['Валовая_прибыль'].sum()
            total_cost = customer_monthly['Себестоимость'].sum() if has_cost_column else 0
            total_quantity = customer_monthly['Количество'].sum()
            total_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
            
            # Заголовок контрагента
            st.markdown(f"""
            <div style='
                background-color: #2E86AB;
                border-radius: 10px;
                padding: 12px;
                margin-top: 20px;
                margin-bottom: 10px;
            '>
                <div style='
                    font-size: 20px;
                    font-weight: bold;
                    color: white;
                    text-align: center;
                '>🏢 {customer_name}</div>
            </div>
            """, unsafe_allow_html=True)
            
            # Метрики контрагента
            col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
            with col_m1:
                st.metric("💰 Выручка", f"{format_number(total_revenue)} ₽")
            with col_m2:
                st.metric("📈 Прибыль", f"{format_number(total_profit)} ₽")
            with col_m3:
                st.metric("📦 Себестоимость", f"{format_number(total_cost)} ₽")
            with col_m4:
                st.metric("🎯 Рентабельность", f"{format_float(total_margin, 1)}%")
            with col_m5:
                st.metric("📦 Продано (шт)", f"{format_number(total_quantity)}")
            
            # Таблица по месяцам
            st.markdown("**📋 Помесячная детализация:**")
            
            # Создаём таблицу для отображения - включаем все нужные колонки
            columns_to_include = ['Название', 'Выручка_без_НДС', 'Валовая_прибыль', 'Количество', 'Рентабельность']
            if has_cost_column and 'Себестоимость' in customer_monthly.columns:
                columns_to_include.append('Себестоимость')
            
            display_df = customer_monthly[columns_to_include].copy()
            
            # Переименовываем колонки
            rename_dict = {
                'Название': 'Месяц',
                'Выручка_без_НДС': 'Выручка',
                'Валовая_прибыль': 'Прибыль',
                'Количество': 'Кол-во (шт)',
                'Рентабельность': 'Рентабельность'
            }
            if has_cost_column and 'Себестоимость' in display_df.columns:
                rename_dict['Себестоимость'] = 'Себестоимость'
            
            display_df = display_df.rename(columns=rename_dict)
            
            # Форматирование - проверяем наличие каждой колонки
            if 'Выручка' in display_df.columns:
                display_df['Выручка'] = display_df['Выручка'].apply(lambda x: f"{format_number(x)} ₽")
            if 'Прибыль' in display_df.columns:
                display_df['Прибыль'] = display_df['Прибыль'].apply(lambda x: f"{format_number(x)} ₽")
            if 'Себестоимость' in display_df.columns:
                display_df['Себестоимость'] = display_df['Себестоимость'].apply(lambda x: f"{format_number(x)} ₽")
            if 'Рентабельность' in display_df.columns:
                display_df['Рентабельность'] = display_df['Рентабельность'].apply(lambda x: f"{format_float(x, 1)}%")
            if 'Кол-во (шт)' in display_df.columns:
                display_df['Кол-во (шт)'] = display_df['Кол-во (шт)'].apply(lambda x: format_number(x))
            
            # Подсвечиваем отрицательную рентабельность
            def highlight_negative(val):
                if isinstance(val, str) and '%' in val:
                    try:
                        num_val = float(val.replace('%', '').replace(',', '.'))
                        if num_val < 0:
                            return 'color: #D9534F; font-weight: bold;'
                    except:
                        pass
                return ''
            
            if 'Рентабельность' in display_df.columns:
                styled_df = display_df.style.applymap(highlight_negative, subset=['Рентабельность'])
            else:
                styled_df = display_df.style
            
            st.dataframe(styled_df, use_container_width=True, hide_index=True)
            
            # График динамики
            fig_customer = go.Figure()
            
            # Выручка и прибыль
            fig_customer.add_trace(go.Bar(
                name='Выручка',
                x=customer_monthly['Название'],
                y=customer_monthly['Выручка_без_НДС'],
                marker_color='#2E86AB',
                text=customer_monthly['Выручка_без_НДС'].apply(lambda x: f'{format_number(x)} ₽'),
                textposition='outside',
                yaxis='y'
            ))
            
            fig_customer.add_trace(go.Bar(
                name='Прибыль',
                x=customer_monthly['Название'],
                y=customer_monthly['Валовая_прибыль'],
                marker_color='#52B788',
                text=customer_monthly['Валовая_прибыль'].apply(lambda x: f'{format_number(x)} ₽'),
                textposition='outside',
                yaxis='y'
            ))
            
            # Рентабельность (линия на второй оси)
            fig_customer.add_trace(go.Scatter(
                name='Рентабельность',
                x=customer_monthly['Название'],
                y=customer_monthly['Рентабельность'],
                mode='lines+markers',
                line=dict(color='#D9534F', width=2),
                marker=dict(size=8, color='#D9534F'),
                yaxis='y2',
                text=customer_monthly['Рентабельность'].apply(lambda x: f'{format_float(x, 1)}%'),
                hovertemplate='%{x}<br>Рентабельность: %{text}<extra></extra>'
            ))
            
            fig_customer.update_layout(
                title=f'Динамика продаж и рентабельности - {customer_name}',
                xaxis_title='Месяц',
                yaxis_title='Сумма (₽)',
                yaxis2=dict(
                    title='Рентабельность (%)',
                    overlaying='y',
                    side='right',
                    range=[-50, 100]
                ),
                barmode='group',
                height=400,
                hovermode='x unified'
            )
            
            st.plotly_chart(fig_customer, use_container_width=True)
            
            # Разделитель
            st.markdown("---")
        
        # ===== ВЫВОДИМ АНАЛИЗ ДЛЯ КАЖДОГО ИЗ ТОП-5 =====
        for customer in top5:
            display_customer_analysis(customer, df_year, available_months_num, has_cost_column)
        
        # ===== ТАБЛИЦА СВОДКА ПО ВСЕМ ТОП-5 =====
        st.subheader(f"📊 СВОДНАЯ ТАБЛИЦА ПО ТОП-5 КОНТРАГЕНТАМ ЗА {selected_year}")
        
        # Собираем данные по всем топ-5 контрагентам
        summary_top5 = []
        for customer in top5:
            df_customer = df_year[df_year['Контрагент'] == customer]
            
            if not df_customer.empty:
                total_rev = df_customer['Выручка_без_НДС'].sum()
                total_profit = df_customer['Валовая_прибыль'].sum()
                total_cost = df_customer['Себестоимость'].sum() if has_cost_column else 0
                total_qty = df_customer['Количество'].sum()
                margin = (total_profit / total_rev * 100) if total_rev > 0 else 0
                
                summary_top5.append({
                    'Контрагент': customer,
                    'Выручка': total_rev,
                    'Прибыль': total_profit,
                    'Себестоимость': total_cost,
                    'Рентабельность': margin,
                    'Кол-во (шт)': total_qty
                })
        
        if summary_top5:
            summary_df = pd.DataFrame(summary_top5)
            summary_df = summary_df.sort_values('Выручка', ascending=False)
            
            # Форматирование для отображения
            display_summary = summary_df.copy()
            display_summary['Выручка'] = display_summary['Выручка'].apply(lambda x: f"{format_number(x)} ₽")
            display_summary['Прибыль'] = display_summary['Прибыль'].apply(lambda x: f"{format_number(x)} ₽")
            display_summary['Себестоимость'] = display_summary['Себестоимость'].apply(lambda x: f"{format_number(x)} ₽")
            display_summary['Рентабельность'] = display_summary['Рентабельность'].apply(lambda x: f"{format_float(x, 1)}%")
            display_summary['Кол-во (шт)'] = display_summary['Кол-во (шт)'].apply(lambda x: format_number(x))
            
            # Подсвечиваем отрицательную рентабельность
            styled_summary = display_summary.style.applymap(
                lambda x: 'color: #D9534F; font-weight: bold;' if isinstance(x, str) and '%' in x and float(x.replace('%', '').replace(',', '.')) < 0 else '',
                subset=['Рентабельность']
            )
            
            st.dataframe(styled_summary, use_container_width=True, hide_index=True)
            
            # Экспорт сводной таблицы
            csv_data = summary_df.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
            st.download_button(
                "📥 Скачать сводку по ТОП-5 контрагентам (CSV)",
                csv_data,
                f"top5_customers_{selected_year}.csv",
                "text/csv",
                use_container_width=True
            )
        
        st.caption(f"📅 Анализ по {selected_year} году | ТОП-5 контрагентов по выручке без НДС")

# ==========================================
# СТРАНИЦА 2: ЛОГИСТИКА
# ==========================================
elif page == "🚚 Логистика":
    st.title("🚚 Аналитика затрат на логистику")
    
    if logistics_df.empty:
        st.warning("⚠️ Файл 'logistics_data.xlsx' не найден или пуст.")
        with st.expander("📌 Требуемая структура файла"):
            st.markdown("""
            Файл должен содержать колонки:
            - `Дата отгрузки` / `Дата заказа`
            - `Итого совокупная Стоимость доставки 1 паллета без НДС`
            - `Кол-во паллет в заказе`
            - `Город`, `Контрагент`, `Категория`, `Подкатегория`
            - `Стоимость товара в заказе Без НДС`
            """)
    else:
        # ===== 1. ПРЕДОБРАБОТКА ДАННЫХ =====
        df_log = logistics_df.copy()
        
        # Фильтруем только строки для анализа
        if 'Принимать к анализу?' in df_log.columns:
            df_log = df_log[df_log['Принимать к анализу?'] == 1]
        
        # ===== ДАТА (столбец K - Дата отгрузки) =====
        if 'Дата отгрузки' in df_log.columns:
            df_log['Дата'] = pd.to_datetime(df_log['Дата отгрузки'], errors='coerce')
        elif 'Дата заказа' in df_log.columns:
            df_log['Дата'] = pd.to_datetime(df_log['Дата заказа'], errors='coerce')
        else:
            st.error("❌ Столбец с датами не найден!")
            st.stop()
        
        # Удаляем строки с некорректными датами
        df_log = df_log.dropna(subset=['Дата'])
        
        if df_log.empty:
            st.warning("⚠️ Нет данных с корректными датами")
            st.stop()
        
        # Создаём временные признаки
        df_log['Год'] = df_log['Дата'].dt.year
        df_log['Месяц'] = df_log['Дата'].dt.month
        df_log['Месяц_год'] = df_log['Дата'].dt.strftime('%Y-%m')
        
        # ===== ЗАТРАТЫ НА ЛОГИСТИКУ (столбец AB) =====
        if len(df_log.columns) > 27:
            ab_column = df_log.columns[27]
            df_log['Логистика_затраты'] = pd.to_numeric(df_log[ab_column], errors='coerce').fillna(0)
        else:
            st.error(f"❌ Столбец AB не найден (всего колонок: {len(df_log.columns)})")
            st.stop()
        
        # ===== КОЛИЧЕСТВО ПАЛЛЕТ (столбец I) =====
        if len(df_log.columns) > 8:
            pallet_column = df_log.columns[8]
            df_log['Кол-во_паллет_сырые'] = pd.to_numeric(df_log[pallet_column], errors='coerce').fillna(0)
        else:
            df_log['Кол-во_паллет_сырые'] = 0
        
        # Группируем для уникальных отгрузок (по дате и городу)
        unique_shipments = df_log.groupby(['Дата', 'Город']).agg({
            'Кол-во_паллет_сырые': 'first',
            'Логистика_затраты': 'first',
        }).reset_index()
        
        df_unique = unique_shipments.copy()
        df_unique['Год'] = df_unique['Дата'].dt.year
        df_unique['Месяц'] = df_unique['Дата'].dt.month
        df_unique['Месяц_год'] = df_unique['Дата'].dt.strftime('%Y-%m')
        
        # Количество строк на одну отгрузку (для пропорционального распределения)
        shipment_counts = df_log.groupby(['Дата', 'Город']).size().reset_index(name='Кол-во_строк')
        df_log = df_log.merge(shipment_counts, on=['Дата', 'Город'], how='left')
        df_log['Кол-во_паллет_скорректированное'] = df_log['Кол-во_паллет_сырые'] / df_log['Кол-во_строк']
        df_log['Кол-во_паллет'] = df_log['Кол-во_паллет_скорректированное']
        
        # ===== СТОИМОСТЬ ТОВАРОВ =====
        if 'Стоимость товара в заказе Без НДС' in df_log.columns:
            df_log['Стоимость_товара'] = pd.to_numeric(df_log['Стоимость товара в заказе Без НДС'], errors='coerce').fillna(0)
        else:
            df_log['Стоимость_товара'] = 0
        
        # ===== 2. ОБЩИЕ МЕТРИКИ =====
        st.subheader("📊 Ключевые показатели логистики")
        
        total_pallets = df_unique['Кол-во_паллет_сырые'].sum()
        total_logistics_cost = df_unique['Логистика_затраты'].sum()
        total_goods_cost = df_log['Стоимость_товара'].sum()
        avg_cost_per_pallet = total_logistics_cost / total_pallets if total_pallets > 0 else 0
        logistics_share = (total_logistics_cost / total_goods_cost * 100) if total_goods_cost > 0 else 0
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("📦 Всего паллет", format_number(total_pallets))
        with col2:
            st.metric("💰 Затраты на логистику", f"{format_number(total_logistics_cost)} ₽")
        with col3:
            st.metric("📊 Средняя стоимость паллеты", f"{format_number(avg_cost_per_pallet)} ₽")
        with col4:
            st.metric("🏷️ Стоимость товаров", f"{format_number(total_goods_cost)} ₽")
        with col5:
            st.metric("📈 Доля логистики", f"{format_float(logistics_share, 1)}%")
        
        st.caption(f"📊 Всего уникальных отгрузок: {len(df_unique)} | Всего строк с номенклатурой: {len(df_log)}")
        
        if total_logistics_cost == 0:
            st.warning("⚠️ Все затраты на логистику равны 0.")
            st.stop()
        
        st.divider()
        
        # ===== 3. ФИЛЬТРЫ =====
        st.subheader("🔍 Фильтрация данных")
        
        col_f1, col_f2, col_f3 = st.columns(3)
        
        with col_f1:
            available_years = sorted(df_unique['Год'].dropna().unique())
            if available_years:
                selected_year = st.selectbox("📅 Год", available_years, index=len(available_years)-1)
            else:
                selected_year = None
        
        with col_f2:
            if 'Город' in df_unique.columns:
                cities_list = df_unique['Город'].dropna()
                cities_list = cities_list[cities_list.apply(lambda x: isinstance(x, str) and x not in ['nan', 'None', ''])]
                cities_unique = sorted(cities_list.unique().tolist())
                cities = ['Все'] + cities_unique if cities_unique else ['Все']
                selected_city = st.selectbox("🏙️ Город", cities)
            else:
                selected_city = 'Все'
        
        with col_f3:
            if 'Категория' in df_log.columns:
                categories_list = df_log['Категория'].dropna()
                categories_list = categories_list[categories_list.apply(lambda x: isinstance(x, str) and x not in ['nan', 'None', ''])]
                categories_unique = sorted(categories_list.unique().tolist())
                categories = ['Все'] + categories_unique if categories_unique else ['Все']
                selected_category = st.selectbox("📁 Категория", categories)
            else:
                selected_category = 'Все'
        
        # Применяем фильтры
        filtered_unique = df_unique.copy()
        if selected_year:
            filtered_unique = filtered_unique[filtered_unique['Год'] == selected_year]
        if selected_city != 'Все':
            filtered_unique = filtered_unique[filtered_unique['Город'] == selected_city]
        
        filtered_df = df_log.copy()
        if selected_year:
            filtered_df = filtered_df[filtered_df['Год'] == selected_year]
        if selected_city != 'Все':
            filtered_df = filtered_df[filtered_df['Город'] == selected_city]
        if selected_category != 'Все':
            filtered_df = filtered_df[filtered_df['Категория'] == selected_category]
        
        if filtered_unique.empty:
            st.warning("⚠️ Нет данных для выбранных фильтров")
            st.stop()
        
        # ===== 4. ДИНАМИКА ПО МЕСЯЦАМ =====
        st.subheader("📈 Динамика логистических затрат")
        
        monthly_stats = filtered_unique.groupby('Месяц_год').agg({
            'Логистика_затраты': 'sum',
            'Кол-во_паллет_сырые': 'sum'
        }).reset_index()
        monthly_stats = monthly_stats.sort_values('Месяц_год')
        
        if not monthly_stats.empty and monthly_stats['Логистика_затраты'].sum() > 0:
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=monthly_stats['Месяц_год'],
                y=monthly_stats['Логистика_затраты'],
                marker_color='#2E86AB',
                text=monthly_stats['Логистика_затраты'].apply(lambda x: f'{format_number(x)} ₽'),
                textposition='outside'
            ))
            fig1.update_layout(
                title='Затраты на логистику по месяцам',
                xaxis_title='Месяц',
                yaxis_title='Затраты (₽)',
                height=400,
                showlegend=False
            )
            st.plotly_chart(fig1, use_container_width=True)
            
            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(
                x=monthly_stats['Месяц_год'],
                y=monthly_stats['Кол-во_паллет_сырые'],
                mode='lines+markers',
                marker=dict(size=10, color='#D9534F'),
                line=dict(width=3, color='#D9534F'),
                name='Паллеты'
            ))
            fig2.update_layout(
                title='Динамика количества паллет',
                xaxis_title='Месяц',
                yaxis_title='Кол-во паллет',
                height=400
            )
            st.plotly_chart(fig2, use_container_width=True)
            
            monthly_stats['Средняя_стоимость_паллеты'] = monthly_stats['Логистика_затраты'] / monthly_stats['Кол-во_паллет_сырые'].replace(0, 1)
            fig3 = go.Figure()
            fig3.add_trace(go.Scatter(
                x=monthly_stats['Месяц_год'],
                y=monthly_stats['Средняя_стоимость_паллеты'],
                mode='lines+markers',
                marker=dict(size=8, color='#52B788'),
                line=dict(width=2, color='#52B788'),
                name='Ср. стоимость паллеты'
            ))
            fig3.update_layout(
                title='Средняя стоимость паллеты по месяцам',
                xaxis_title='Месяц',
                yaxis_title='Стоимость (₽)',
                height=400
            )
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("📊 Нет данных для построения графиков")
        
        st.divider()
        
        # ===== 5. АНАЛИЗ ПО ГОРОДАМ =====
        if 'Город' in filtered_unique.columns and filtered_unique['Логистика_затраты'].sum() > 0:
            st.subheader("🏙️ Анализ по городам")
            
            city_stats = filtered_unique.groupby('Город').agg({
                'Логистика_затраты': 'sum',
                'Кол-во_паллет_сырые': 'sum'
            }).reset_index()
            
            shipment_count = filtered_unique.groupby('Город').size().reset_index(name='Кол-во_отгрузок')
            city_stats = city_stats.merge(shipment_count, on='Город', how='left')
            
            city_stats = city_stats[
                (city_stats['Логистика_затраты'] > 0) & 
                (city_stats['Город'].apply(lambda x: isinstance(x, str) and x not in ['nan', 'None', '']))
            ]
            
            if not city_stats.empty:
                city_stats['Средняя_стоимость_паллеты'] = city_stats['Логистика_затраты'] / city_stats['Кол-во_паллет_сырые'].replace(0, 1)
                city_stats = city_stats.sort_values('Логистика_затраты', ascending=False)
                
                top_cities = city_stats.head(10)
                
                fig4 = px.bar(
                    top_cities,
                    x='Город',
                    y='Логистика_затраты',
                    title='Топ-10 городов по затратам на логистику',
                    labels={'Логистика_затраты': 'Затраты (₽)'},
                    color='Средняя_стоимость_паллеты',
                    color_continuous_scale='Blues'
                )
                fig4.update_layout(height=450)
                st.plotly_chart(fig4, use_container_width=True)
                
                st.markdown("**📊 Детализация по городам**")
                display_cities = city_stats.copy()
                display_cities['Затраты'] = display_cities['Логистика_затраты'].apply(lambda x: f"{format_number(x)} ₽")
                display_cities['Паллеты'] = display_cities['Кол-во_паллет_сырые'].apply(lambda x: format_number(x))
                display_cities['Ср. цена паллеты'] = display_cities['Средняя_стоимость_паллеты'].apply(lambda x: f"{format_number(x)} ₽")
                display_cities['Отгрузок'] = display_cities['Кол-во_отгрузок'].apply(lambda x: format_number(x))
                
                st.dataframe(
                    display_cities[['Город', 'Затраты', 'Паллеты', 'Ср. цена паллеты', 'Отгрузок']],
                    use_container_width=True,
                    hide_index=True
                )
        
        st.divider()
        
        # ===== 6. ДЕТАЛЬНАЯ ТАБЛИЦА =====
        st.subheader("📋 Детализация по поставкам (по номенклатурам)")
        
        display_cols = []
        col_map = {
            'Дата': '📅 Дата отгрузки',
            'Город': '🏙️ Город',
            'Категория': '📁 Категория',
            'Подкатегория': '📂 Подкатегория',
            'Номенклатура': '📦 Номенклатура',
            'Кол-во_паллет': '📦 Паллеты (скорректированные)',
            'Логистика_затраты': '💰 Логистика (без НДС)',
            'Стоимость_товара': '🏷️ Товары (без НДС)',
            'Кол-во_строк': '📋 Строк в отгрузке'
        }
        
        for col in col_map.keys():
            if col in filtered_df.columns:
                display_cols.append(col)
        
        if display_cols:
            table_df = filtered_df[display_cols].copy().head(100)
            
            if 'Логистика_затраты' in table_df.columns:
                table_df['Логистика_затраты'] = table_df['Логистика_затраты'].apply(lambda x: f"{format_number(x)} ₽")
            if 'Кол-во_паллет' in table_df.columns:
                table_df['Кол-во_паллет'] = table_df['Кол-во_паллет'].apply(lambda x: format_number(x) if pd.notna(x) else "0")
            if 'Стоимость_товара' in table_df.columns:
                table_df['Стоимость_товара'] = table_df['Стоимость_товара'].apply(lambda x: f"{format_number(x)} ₽")
            if 'Кол-во_строк' in table_df.columns:
                table_df['Кол-во_строк'] = table_df['Кол-во_строк'].apply(lambda x: format_number(x))
            if 'Дата' in table_df.columns:
                table_df['Дата'] = table_df['Дата'].dt.strftime('%d.%m.%Y')
            
            table_df = table_df.rename(columns=col_map)
            st.dataframe(table_df, use_container_width=True)
            st.caption(f"📄 Показано {min(100, len(table_df))} из {len(filtered_df)} записей")
            st.info("💡 **Примечание:** Паллеты распределены пропорционально между номенклатурами в одной отгрузке. Общее количество паллет считается по уникальным отгрузкам.")
        else:
            st.info("Нет данных для отображения")

# ==========================================
# СТРАНИЦА 3: АНАЛИЗ СЕБЕСТОИМОСТИ
# ==========================================
elif page == "📊 Анализ себестоимости":
    st.title("📊 Анализ себестоимости продукции")
    
    if sales_df.empty:
        st.warning("⚠️ Нет данных о продажах для анализа себестоимости.")
        st.info("📌 Пожалуйста, добавьте файл 'sales_data.xlsx' в папку с приложением.")
    else:
        st.markdown("""
        ### Динамика себестоимости без НДС на единицу продукции
        **Формула:** Себестоимость без НДС / Количество
        """)
        
        # Подготовка данных
        df_cost = sales_df.copy()
        
        # Создаём временный столбец для сортировки дат
        df_cost['Дата_сортировка'] = pd.to_datetime(df_cost['Дата'], errors='coerce')
        
        # Создаём столбец с русскими названиями месяцев
        df_cost['Месяц_рус'] = df_cost['Дата_сортировка'].dt.month.map(month_names_ru)
        df_cost['Год'] = df_cost['Дата_сортировка'].dt.year
        df_cost['Период'] = df_cost['Дата_сортировка'].dt.strftime('%Y-%m')
        df_cost['Период_рус'] = df_cost['Дата_сортировка'].dt.strftime('%Y') + ' ' + df_cost['Месяц_рус']
        
        # Убираем дубликаты столбцов
        df_cost = df_cost.loc[:, ~df_cost.columns.duplicated()]
        
        # Рассчитываем себестоимость на единицу продукции
        if 'Себестоимость' in df_cost.columns and 'Количество' in df_cost.columns:
            df_cost['Себестоимость_единицы'] = df_cost['Себестоимость'] / df_cost['Количество']
            df_cost['Себестоимость_единицы'] = df_cost['Себестоимость_единицы'].replace([float('inf'), -float('inf')], 0).fillna(0)
        else:
            st.error("❌ Не найдены столбцы 'Себестоимость' и 'Количество' в данных.")
            st.stop()
        
        # ==========================================
        # ФУНКЦИЯ ДЛЯ IQR-АНАЛИЗА
        # ==========================================
        def detect_outliers_iqr(data, multiplier=1.5):
            """
            Определяет выбросы методом IQR (межквартильного расстояния)
            """
            data_clean = data.dropna()
            if len(data_clean) < 4:
                return {
                    'outliers': [],
                    'q1': None,
                    'q3': None,
                    'lower_bound': None,
                    'upper_bound': None,
                    'outlier_indices': [],
                    'iqr': None
                }
            
            q1 = data_clean.quantile(0.25)
            q3 = data_clean.quantile(0.75)
            iqr = q3 - q1
            
            lower_bound = q1 - multiplier * iqr
            upper_bound = q3 + multiplier * iqr
            
            outlier_mask = (data < lower_bound) | (data > upper_bound)
            outliers = data[outlier_mask]
            outlier_indices = data.index[outlier_mask].tolist()
            
            return {
                'outliers': outliers.tolist(),
                'q1': q1,
                'q3': q3,
                'lower_bound': lower_bound,
                'upper_bound': upper_bound,
                'outlier_indices': outlier_indices,
                'iqr': iqr
            }
        
        # ==========================================
        # ОПРЕДЕЛЯЕМ СТОЛБЕЦ ДЛЯ ГРУППИРОВКИ
        # ==========================================
        group_col = None
        possible_group_cols = ['Группа', 'Подгруппа', 'Категория', 'Группа товаров', 'Категория товара']
        
        for col in possible_group_cols:
            if col in df_cost.columns:
                group_col = col
                break
        
        if group_col is None:
            for col in df_cost.columns:
                col_lower = str(col).lower()
                if 'групп' in col_lower or 'категор' in col_lower:
                    group_col = col
                    break
        
        if group_col is None:
            st.warning("⚠️ Не найден столбец для группировки номенклатур")
            group_col = 'Номенклатура'
        
        # Получаем список уникальных групп
        unique_groups = df_cost[group_col].dropna().unique()
        unique_groups = sorted(unique_groups)
        
        if len(unique_groups) == 0:
            st.warning("Нет данных для отображения")
        else:
            # ===== ВЫБОР ГРУППЫ =====
            col_select1, col_select2 = st.columns([2, 1])
            
            with col_select1:
                selected_group = st.selectbox(
                    "📂 Выберите группу номенклатур",
                    unique_groups,
                    help=f"Группировка по столбцу '{group_col}'"
                )
            
            with col_select2:
                show_iqr_analysis = st.checkbox(
                    "📊 Показать IQR-анализ выбросов",
                    value=True,
                    help="Метод межквартильного расстояния для обнаружения выбросов"
                )
            
            # Фильтруем номенклатуры по выбранной группе
            df_group = df_cost[df_cost[group_col] == selected_group]
            
            if df_group.empty:
                st.warning(f"Нет данных для группы '{selected_group}'")
            else:
                # ===== ФИЛЬТРЫ =====
                st.divider()
                
                col_filter1, col_filter2 = st.columns(2)
                
                with col_filter1:
                    period_options = ["За всё время", "Текущий год", "Прошлый год", "Текущий и прошлый год"]
                    default_index = period_options.index("Текущий и прошлый год")
                    period_filter = st.selectbox(
                        "📅 Период",
                        period_options,
                        index=default_index,
                        help="Ограничить период отображаемых данных"
                    )
                
                with col_filter2:
                    sort_options = [
                        "По убыванию средней себестоимости",
                        "По возрастанию средней себестоимости",
                        "По стандартному отклонению (от наибольшего)",
                        "По стандартному отклонению (от наименьшего)",
                        "По подкатегории",
                        "По категории",
                        "По группе",
                        "По номенклатуре"
                    ]
                    default_sort_index = sort_options.index("По стандартному отклонению (от наибольшего)")
                    sort_by = st.selectbox(
                        "📊 Сортировка графиков",
                        sort_options,
                        index=default_sort_index,
                        help="Порядок вывода графиков сверху вниз"
                    )
                
                # ===== ПРИМЕНЯЕМ ФИЛЬТР ПО ДАТЕ =====
                current_year = pd.Timestamp.now().year
                available_years = sorted(df_cost['Год'].dropna().unique())
                
                df_cost_filtered = df_group.copy()
                period_caption = "за всё время"
                
                if period_filter == "Текущий год":
                    if current_year in available_years:
                        df_cost_filtered = df_group[df_group['Год'] == current_year]
                        period_caption = f"за {current_year} год"
                    else:
                        st.warning(f"Нет данных за {current_year} год. Отображаются все данные.")
                elif period_filter == "Прошлый год":
                    prev_year = current_year - 1
                    if prev_year in available_years:
                        df_cost_filtered = df_group[df_group['Год'] == prev_year]
                        period_caption = f"за {prev_year} год"
                    else:
                        st.warning(f"Нет данных за {prev_year} год. Отображаются все данные.")
                elif period_filter == "Текущий и прошлый год":
                    prev_year = current_year - 1
                    years_to_show = [y for y in [current_year, prev_year] if y in available_years]
                    if years_to_show:
                        df_cost_filtered = df_group[df_group['Год'].isin(years_to_show)]
                        period_caption = f"за {', '.join(map(str, years_to_show))} годы"
                    else:
                        st.warning("Нет данных за текущий или прошлый год. Отображаются все данные.")
                
                # ===== ПОЛУЧАЕМ СПИСОК НОМЕНКЛАТУР =====
                nomenclatures = df_cost_filtered['Номенклатура'].dropna().unique()
                
                # ===== СОБИРАЕМ СТАТИСТИКУ =====
                nomen_stats = []
                for nomen in nomenclatures:
                    df_nomen = df_cost_filtered[df_cost_filtered['Номенклатура'] == nomen]
                    if not df_nomen.empty:
                        monthly_data = df_nomen.groupby('Период')['Себестоимость_единицы'].mean()
                        monthly_data = monthly_data.sort_index()
                        
                        if not monthly_data.empty:
                            periods_sorted = monthly_data.index.tolist()
                            periods_ru = []
                            for p in periods_sorted:
                                try:
                                    year, month = p.split('-')
                                    month_ru = month_names_ru[int(month)]
                                    periods_ru.append(f"{year} {month_ru}")
                                except (ValueError, KeyError):
                                    periods_ru.append(p)
                            
                            avg = monthly_data.mean()
                            std = monthly_data.std() if len(monthly_data) > 1 else 0
                            
                            subcat = df_nomen['Подкатегория'].iloc[0] if 'Подкатегория' in df_nomen.columns else ''
                            category = df_nomen['Категория'].iloc[0] if 'Категория' in df_nomen.columns else ''
                            nomen_group = df_nomen[group_col].iloc[0] if group_col in df_nomen.columns else ''
                            
                            if len(monthly_data) >= 2:
                                first_val = monthly_data.iloc[0]
                                last_val = monthly_data.iloc[-1]
                                trend = (last_val - first_val) / first_val * 100 if first_val > 0 else 0
                            else:
                                trend = 0
                            
                            iqr_result = detect_outliers_iqr(monthly_data, multiplier=1.5)
                            
                            nomen_stats.append({
                                'Номенклатура': nomen,
                                'Средняя': avg,
                                'Стд_отклонение': std,
                                'Подкатегория': subcat,
                                'Категория': category,
                                'Группа': nomen_group,
                                'Максимум': monthly_data.max(),
                                'Минимум': monthly_data.min(),
                                'Последнее': monthly_data.iloc[-1] if len(monthly_data) > 0 else 0,
                                'Первое': monthly_data.iloc[0] if len(monthly_data) > 0 else 0,
                                'Тренд_%': trend,
                                'Количество_периодов': len(monthly_data),
                                'Данные': monthly_data.values.tolist(),
                                'Периоды': periods_ru,
                                'Периоды_сорт': periods_sorted,
                                'IQR_анализ': iqr_result,
                                'Количество_выбросов': len(iqr_result['outliers']) if iqr_result['outliers'] else 0
                            })
                
                # ===== СОРТИРУЕМ =====
                if sort_by == "По убыванию средней себестоимости":
                    nomen_stats.sort(key=lambda x: x['Средняя'], reverse=True)
                    sort_caption = "от наибольшей средней себестоимости"
                elif sort_by == "По возрастанию средней себестоимости":
                    nomen_stats.sort(key=lambda x: x['Средняя'])
                    sort_caption = "от наименьшей средней себестоимости"
                elif sort_by == "По стандартному отклонению (от наибольшего)":
                    nomen_stats.sort(key=lambda x: x['Стд_отклонение'], reverse=True)
                    sort_caption = "от наибольшей волатильности"
                elif sort_by == "По стандартному отклонению (от наименьшего)":
                    nomen_stats.sort(key=lambda x: x['Стд_отклонение'])
                    sort_caption = "от наименьшей волатильности"
                elif sort_by == "По подкатегории":
                    nomen_stats.sort(key=lambda x: x['Подкатегория'])
                    sort_caption = "по подкатегории"
                elif sort_by == "По категории":
                    nomen_stats.sort(key=lambda x: x['Категория'])
                    sort_caption = "по категории"
                elif sort_by == "По группе":
                    nomen_stats.sort(key=lambda x: x['Группа'])
                    sort_caption = "по группе"
                else:
                    nomen_stats.sort(key=lambda x: x['Номенклатура'])
                    sort_caption = "по наименованию"
                
                # ===== ОТОБРАЖАЕМ ЗАГОЛОВОК =====
                st.divider()
                st.subheader(f"📦 Номенклатуры в группе: {selected_group}")
                st.caption(f"📅 {period_caption} | 📊 Сортировка: {sort_caption} | Всего номенклатур: {len(nomen_stats)}")
                
                # ===== IQR-СВОДКА ПО ВСЕЙ ГРУППЕ =====
                if show_iqr_analysis and len(nomen_stats) > 0:
                    all_costs = []
                    for stat in nomen_stats:
                        all_costs.extend(stat['Данные'])
                    
                    if all_costs:
                        all_costs_series = pd.Series(all_costs)
                        iqr_all = detect_outliers_iqr(all_costs_series, multiplier=1.5)
                        
                        if iqr_all['q1'] is not None:
                            st.subheader("📊 IQR-АНАЛИЗ ВСЕЙ ГРУППЫ")
                            
                            col_iqr1, col_iqr2, col_iqr3, col_iqr4 = st.columns(4)
                            with col_iqr1:
                                st.metric("Q1 (25-й перцентиль)", f"{format_number(iqr_all['q1'])} ₽/ед.")
                            with col_iqr2:
                                st.metric("Q3 (75-й перцентиль)", f"{format_number(iqr_all['q3'])} ₽/ед.")
                            with col_iqr3:
                                st.metric("IQR (Q3 - Q1)", f"{format_number(iqr_all['iqr'])} ₽/ед.")
                            with col_iqr4:
                                st.metric("🔴 Количество выбросов", len(iqr_all['outliers']))
                            
                            col_bound1, col_bound2 = st.columns(2)
                            with col_bound1:
                                st.metric("⬇️ Нижняя граница", f"{format_number(iqr_all['lower_bound'])} ₽/ед.")
                            with col_bound2:
                                st.metric("⬆️ Верхняя граница", f"{format_number(iqr_all['upper_bound'])} ₽/ед.")
                            
                            st.subheader("📦 Boxplot всей группы")
                            fig_box = go.Figure()
                            fig_box.add_trace(go.Box(
                                y=all_costs,
                                name=selected_group,
                                boxmean='sd',
                                marker_color='#2E86AB',
                                line_color='#1A5276'
                            ))
                            fig_box.update_layout(
                                title=f'Распределение себестоимости в группе "{selected_group}"',
                                yaxis_title='Себестоимость (₽/ед.)',
                                height=400,
                                showlegend=False
                            )
                            st.plotly_chart(fig_box, use_container_width=True)
                            
                            if len(iqr_all['outliers']) > 0:
                                st.warning(f"⚠️ Обнаружено {len(iqr_all['outliers'])} выбросов в группе")
                                
                                outlier_nomenclatures = []
                                for stat in nomen_stats:
                                    if stat['Количество_выбросов'] > 0:
                                        outlier_nomenclatures.append({
                                            'Номенклатура': stat['Номенклатура'],
                                            'Количество_выбросов': stat['Количество_выбросов'],
                                            'Максимум': stat['Максимум'],
                                            'Минимум': stat['Минимум'],
                                            'Средняя': stat['Средняя']
                                        })
                                
                                if outlier_nomenclatures:
                                    outlier_df = pd.DataFrame(outlier_nomenclatures)
                                    outlier_df = outlier_df.sort_values('Количество_выбросов', ascending=False)
                                    st.markdown("**🔴 Номенклатуры с выбросами:**")
                                    display_outliers = outlier_df.copy()
                                    for col in ['Максимум', 'Минимум', 'Средняя']:
                                        if col in display_outliers.columns:
                                            display_outliers[col] = display_outliers[col].apply(lambda x: f"{format_number(x)} ₽/ед.")
                                    st.dataframe(display_outliers, use_container_width=True, hide_index=True)
                            else:
                                st.success("✅ Выбросы в группе не обнаружены")
                            
                            st.divider()
                
                # ===== ДЛЯ КАЖДОЙ НОМЕНКЛАТУРЫ СТРОИМ ГРАФИК =====
                for idx, stat in enumerate(nomen_stats):
                    nomen = stat['Номенклатура']
                    periods = stat['Периоды']
                    costs = stat['Данные']
                    
                    col_left, col_right = st.columns([0.3, 0.7])
                    
                    with col_left:
                        trend_color = "🟢" if stat['Тренд_%'] <= 0 else "🔴"
                        trend_text = f"{trend_color} {format_float(stat['Тренд_%'], 1)}%"
                        outlier_indicator = "🔴" if stat['Количество_выбросов'] > 0 else "🟢"
                        outlier_text = f"{stat['Количество_выбросов']}" if stat['Количество_выбросов'] > 0 else "нет"
                        
                        st.markdown(f"""
                        <div style='
                            background-color: #F0F2F6;
                            border-radius: 10px;
                            padding: 12px;
                            margin-top: 30px;
                        '>
                            <div style='font-weight: bold; font-size: 16px; margin-bottom: 10px;'>{nomen}</div>
                            <div style='font-size: 13px; color: #666;'>Среднее:</div>
                            <div style='font-size: 18px; font-weight: bold; color: #2E86AB;'>{format_number(stat['Средняя'])} ₽/ед.</div>
                            <div style='font-size: 13px; color: #666; margin-top: 8px;'>Максимум:</div>
                            <div style='font-size: 16px; font-weight: bold; color: #D9534F;'>{format_number(stat['Максимум'])} ₽/ед.</div>
                            <div style='font-size: 13px; color: #666; margin-top: 8px;'>Минимум:</div>
                            <div style='font-size: 16px; font-weight: bold; color: #5CB85C;'>{format_number(stat['Минимум'])} ₽/ед.</div>
                            <div style='font-size: 13px; color: #666; margin-top: 8px;'>Тренд:</div>
                            <div style='font-size: 16px; font-weight: bold;'>{trend_text}</div>
                            <div style='font-size: 13px; color: #666; margin-top: 8px;'>Выбросы (IQR):</div>
                            <div style='font-size: 16px; font-weight: bold;'>{outlier_indicator} {outlier_text}</div>
                            <div style='font-size: 13px; color: #666; margin-top: 8px;'>Стд. отклонение:</div>
                            <div style='font-size: 14px;'>{format_number(stat['Стд_отклонение'])}</div>
                            <div style='font-size: 13px; color: #666; margin-top: 8px;'>Периодов:</div>
                            <div style='font-size: 14px;'>{stat['Количество_периодов']}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col_right:
                        fig = go.Figure()
                        
                        fig.add_trace(go.Scatter(
                            x=periods,
                            y=costs,
                            mode='lines+markers',
                            line=dict(width=2, color='#2E86AB'),
                            marker=dict(size=8, color='#1A5276'),
                            name='Себестоимость',
                            hovertemplate='%{x}<br>Себестоимость: %{y:.2f} ₽/ед.<extra></extra>'
                        ))
                        
                        if show_iqr_analysis and stat['IQR_анализ']['q1'] is not None:
                            iqr_res = stat['IQR_анализ']
                            
                            if iqr_res['lower_bound'] is not None:
                                fig.add_hline(
                                    y=iqr_res['lower_bound'],
                                    line_dash="dash",
                                    line_color="orange",
                                    annotation_text=f"Нижняя граница: {format_number(iqr_res['lower_bound'])}",
                                    annotation_position="bottom left"
                                )
                            if iqr_res['upper_bound'] is not None:
                                fig.add_hline(
                                    y=iqr_res['upper_bound'],
                                    line_dash="dash",
                                    line_color="orange",
                                    annotation_text=f"Верхняя граница: {format_number(iqr_res['upper_bound'])}",
                                    annotation_position="top left"
                                )
                            
                            if iqr_res['outliers']:
                                outlier_periods = []
                                outlier_y = []
                                
                                for i, val in enumerate(iqr_res['outliers']):
                                    if i < len(periods):
                                        outlier_periods.append(periods[i])
                                        outlier_y.append(val)
                                
                                if not outlier_periods and iqr_res['outliers']:
                                    max_points = min(len(periods), len(iqr_res['outliers']))
                                    outlier_periods = periods[:max_points]
                                    outlier_y = iqr_res['outliers'][:max_points]
                                
                                if outlier_periods and outlier_y:
                                    fig.add_trace(go.Scatter(
                                        x=outlier_periods,
                                        y=outlier_y,
                                        mode='markers',
                                        marker=dict(size=15, color='red', symbol='x'),
                                        name='Выбросы (IQR)',
                                        hovertemplate='%{x}<br>⚠️ ВЫБРОС: %{y:.2f} ₽/ед.<extra></extra>'
                                    ))
                        
                        fig.add_hline(
                            y=stat['Средняя'],
                            line_dash="dash",
                            line_color="red",
                            annotation_text=f"Среднее: {format_number(stat['Средняя'])} ₽/ед.",
                            annotation_position="top right"
                        )
                        
                        if stat['Стд_отклонение'] > 0:
                            fig.add_hrect(
                                y0=stat['Средняя'] - stat['Стд_отклонение'],
                                y1=stat['Средняя'] + stat['Стд_отклонение'],
                                fillcolor="rgba(46, 134, 171, 0.2)",
                                line_width=0,
                                annotation_text="±1σ",
                                annotation_position="bottom right"
                            )
                        
                        fig.update_layout(
                            title="",
                            xaxis_title="Период (месяц/год)",
                            yaxis_title="Себестоимость без НДС на единицу (₽/ед.)",
                            hovermode='x unified',
                            height=350,
                            margin=dict(l=40, r=40, t=30, b=40)
                        )
                        
                        fig.update_xaxes(tickangle=-45)
                        
                        st.plotly_chart(fig, use_container_width=True, key=f"cost_chart_{idx}_{nomen[:30]}")
                    
                    if show_iqr_analysis and stat['Количество_периодов'] >= 4 and stat['IQR_анализ']['q1'] is not None:
                        with st.expander(f"📊 Boxplot для {nomen} (IQR-анализ)"):
                            fig_box_nomen = go.Figure()
                            fig_box_nomen.add_trace(go.Box(
                                y=costs,
                                name=nomen,
                                boxmean='sd',
                                marker_color='#2E86AB',
                                line_color='#1A5276'
                            ))
                            
                            if stat['IQR_анализ']['outliers']:
                                iqr_res = stat['IQR_анализ']
                                fig_box_nomen.add_trace(go.Scatter(
                                    y=iqr_res['outliers'],
                                    mode='markers',
                                    marker=dict(size=10, color='red', symbol='x'),
                                    name='Выбросы'
                                ))
                            
                            fig_box_nomen.update_layout(
                                title=f'Распределение себестоимости для "{nomen}"',
                                yaxis_title='Себестоимость (₽/ед.)',
                                height=300,
                                showlegend=True
                            )
                            st.plotly_chart(fig_box_nomen, use_container_width=True)
                            
                            iqr_res = stat['IQR_анализ']
                            col_iqr_d1, col_iqr_d2, col_iqr_d3, col_iqr_d4 = st.columns(4)
                            with col_iqr_d1:
                                st.metric("Q1", f"{format_number(iqr_res['q1'])} ₽/ед.")
                            with col_iqr_d2:
                                st.metric("Q3", f"{format_number(iqr_res['q3'])} ₽/ед.")
                            with col_iqr_d3:
                                st.metric("IQR", f"{format_number(iqr_res['iqr'])} ₽/ед.")
                            with col_iqr_d4:
                                st.metric("Выбросов", len(iqr_res['outliers']))
                    
                    st.divider()
                
                # ===== СВОДНАЯ ТАБЛИЦА =====
                if len(nomen_stats) > 0:
                    st.subheader(f"📊 Сводная таблица себестоимости единицы по номенклатурам группы '{selected_group}'")
                    st.caption(f"{period_caption}")
                    
                    summary_data = []
                    for stat in nomen_stats:
                        summary_data.append({
                            'Номенклатура': stat['Номенклатура'],
                            'Средняя себестоимость ед.': stat['Средняя'],
                            'Стд. отклонение': stat['Стд_отклонение'],
                            'Максимум': stat['Максимум'],
                            'Минимум': stat['Минимум'],
                            'Последнее': stat['Последнее'],
                            'Тренд, %': stat['Тренд_%'],
                            'Периодов': stat['Количество_периодов'],
                            'Выбросов (IQR)': stat['Количество_выбросов'],
                            'Подкатегория': stat['Подкатегория'],
                            'Категория': stat['Категория']
                        })
                    
                    summary_df = pd.DataFrame(summary_data)
                    if not summary_df.empty:
                        display_summary = summary_df.copy()
                        for col in ['Средняя себестоимость ед.', 'Стд. отклонение', 'Максимум', 'Минимум', 'Последнее']:
                            if col in display_summary.columns:
                                display_summary[col] = display_summary[col].apply(lambda x: format_number(x) if pd.notna(x) else "0")
                        
                        if 'Тренд, %' in display_summary.columns:
                            display_summary['Тренд, %'] = display_summary['Тренд, %'].apply(lambda x: f"{format_float(x, 1)}%")
                        
                        st.dataframe(
                            display_summary,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Выбросов (IQR)": st.column_config.NumberColumn(
                                    "Выбросов (IQR)",
                                    help="Количество выбросов по методу IQR",
                                )
                            }
                        )
                        
                        csv_summary = summary_df.copy()
                        for col in ['Средняя себестоимость ед.', 'Стд. отклонение', 'Максимум', 'Минимум', 'Последнее', 'Тренд, %']:
                            if col in csv_summary.columns:
                                csv_summary[col] = csv_summary[col].apply(lambda x: float(x) if pd.notna(x) else 0)
                        
                        csv_data = csv_summary.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
                        st.download_button(
                            "📥 Скачать сводную таблицу (CSV)",
                            csv_data,
                            f"cost_per_unit_summary_{selected_group}_{pd.Timestamp.now().strftime('%Y%m%d')}.csv",
                            "text/csv"
                        )
                
                # ===== ДОПОЛНИТЕЛЬНЫЙ АНАЛИЗ =====
                if len(nomen_stats) > 1 and 'Подкатегория' in df_cost_filtered.columns:
                    st.divider()
                    st.subheader("📊 Сравнительный анализ волатильности себестоимости")
                    
                    category_analysis = df_cost_filtered.groupby('Подкатегория')['Себестоимость_единицы'].agg(['mean', 'std', 'min', 'max', 'count']).reset_index()
                    category_analysis = category_analysis.rename(columns={
                        'Подкатегория': 'Группа',
                        'mean': 'Средняя',
                        'std': 'Стд_отклонение',
                        'min': 'Минимум',
                        'max': 'Максимум',
                        'count': 'Количество_записей'
                    })
                    
                    if not category_analysis.empty:
                        st.markdown("**Анализ по подкатегориям:**")
                        display_cat = category_analysis.copy()
                        for col in ['Средняя', 'Стд_отклонение', 'Минимум', 'Максимум']:
                            display_cat[col] = display_cat[col].apply(lambda x: format_number(x) if pd.notna(x) else "0")
                        st.dataframe(display_cat, use_container_width=True, hide_index=True)
                        
                        fig_compare = go.Figure()
                        fig_compare.add_trace(go.Bar(
                            x=category_analysis['Группа'],
                            y=category_analysis['Средняя'],
                            marker_color='#2E86AB',
                            error_y=dict(
                                type='data',
                                array=category_analysis['Стд_отклонение'],
                                visible=True
                            ),
                            name='Средняя ± отклонение'
                        ))
                        fig_compare.update_layout(
                            title='Средняя себестоимость по подкатегориям',
                            xaxis_title='Подкатегория',
                            yaxis_title='Себестоимость (₽/ед.)',
                            height=400
                        )
                        st.plotly_chart(fig_compare, use_container_width=True)

# ==========================================
# СТРАНИЦА 4: ФОРМИРОВАНИЕ СЕБЕСТОИМОСТИ ПФ
# ==========================================
elif page == "🏭 Формирование себестоимости ПФ":
    st.title("🏭 Формирование себестоимости полуфабриката")
    
    if production_df.empty:
        st.warning("⚠️ Файл 'production_data.xlsx' не найден или не удалось загрузить данные.")
        st.info("📌 Пожалуйста, добавьте файл с данными о производстве в папку с приложением.")
    else:
        st.markdown("### Анализ себестоимости продукта **П/Ф Дрип Гватемала Декаф 1шт.**")
        
        batches = production_df[production_df['Тип'] == 'Партия'].copy()
        materials = production_df[production_df['Тип'] == 'Сырье'].copy()
        
        if batches.empty:
            st.warning("Не удалось распознать партии в файле.")
            with st.expander("🔧 Показать загруженные данные"):
                st.write(production_df.head(20))
        else:
            st.success(f"✅ Загружено {len(batches)} партий и {len(materials)} записей о сырье")
            
            col_filter1, col_filter2 = st.columns(2)
            
            with col_filter1:
                batch_options = ['Все партии'] + sorted(batches['Партия'].astype(str).unique().tolist())
                selected_batch = st.selectbox("📦 Выберите партию для детального анализа", batch_options)
            
            with col_filter2:
                sort_by = st.selectbox(
                    "📊 Сортировка партий",
                    ["По номеру (старые сверху)", "По номеру (новые сверху)", "По себестоимости (от низкой)", "По себестоимости (от высокой)"]
                )
            
            st.divider()
            st.subheader("📊 ОБЩАЯ СТАТИСТИКА")
            
            avg_cost = batches['Себестоимость_единицы'].mean()
            min_cost = batches['Себестоимость_единицы'].min()
            max_cost = batches['Себестоимость_единицы'].max()
            total_quantity = batches['Количество_выпущено'].sum()
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("📦 Всего выпущено", f"{format_number(total_quantity)} шт.")
            with c2:
                st.metric("📋 Количество партий", len(batches))
            with c3:
                st.metric("💰 Средняя себестоимость", f"{format_number(avg_cost)} ₽/шт.")
            with c4:
                st.metric("📊 Разброс", f"{format_number(min_cost)} - {format_number(max_cost)} ₽/шт.")
            
            st.divider()
            
            if len(batches) >= 2:
                st.subheader("💰 СРАВНЕНИЕ САМОЙ ДЕШЕВОЙ И САМОЙ ДОРОГОЙ ПАРТИИ")
                
                cheapest_batch = batches.loc[batches['Себестоимость_единицы'].idxmin()]
                most_expensive_batch = batches.loc[batches['Себестоимость_единицы'].idxmax()]
                
                cheapest_materials = materials[materials['Партия'].astype(str) == str(cheapest_batch['Партия'])]
                expensive_materials = materials[materials['Партия'].astype(str) == str(most_expensive_batch['Партия'])]
                
                col_left, col_right = st.columns(2)
                
                with col_left:
                    st.markdown(f"""
                    <div style='
                        background-color: #D4EDDA;
                        border-radius: 10px;
                        padding: 15px;
                        text-align: center;
                        border: 1px solid #28A745;
                    '>
                        <div style='font-size: 18px; font-weight: bold; color: #155724;'>🟢 САМАЯ ДЕШЕВАЯ ПАРТИЯ</div>
                        <div style='font-size: 24px; font-weight: bold; margin-top: 10px;'>{cheapest_batch['Партия']}</div>
                        <div style='font-size: 20px; font-weight: bold; color: #28A745; margin-top: 10px;'>{format_number(cheapest_batch['Себестоимость_единицы'])} ₽/шт.</div>
                        <div style='font-size: 14px; color: #666; margin-top: 5px;'>Выпущено: {format_number(cheapest_batch['Количество_выпущено'])} шт.</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if not cheapest_materials.empty:
                        st.markdown("**Состав сырья:**")
                        cheapest_display = cheapest_materials[['Сырье', 'Себестоимость_на_единицу_продукции']].copy()
                        cheapest_display['Доля_в_себестоимости'] = cheapest_display['Себестоимость_на_единицу_продукции'] / cheapest_batch['Себестоимость_единицы'] * 100
                        cheapest_display = cheapest_display.sort_values('Себестоимость_на_единицу_продукции', ascending=False)
                        for _, row in cheapest_display.iterrows():
                            st.write(f"• {row['Сырье']}: {format_number(row['Себестоимость_на_единицу_продукции'])} ₽/шт. ({format_float(row['Доля_в_себестоимости'], 1)}%)")
                
                with col_right:
                    st.markdown(f"""
                    <div style='
                        background-color: #F8D7DA;
                        border-radius: 10px;
                        padding: 15px;
                        text-align: center;
                        border: 1px solid #DC3545;
                    '>
                        <div style='font-size: 18px; font-weight: bold; color: #721C24;'>🔴 САМАЯ ДОРОГАЯ ПАРТИЯ</div>
                        <div style='font-size: 24px; font-weight: bold; margin-top: 10px;'>{most_expensive_batch['Партия']}</div>
                        <div style='font-size: 20px; font-weight: bold; color: #DC3545; margin-top: 10px;'>{format_number(most_expensive_batch['Себестоимость_единицы'])} ₽/шт.</div>
                        <div style='font-size: 14px; color: #666; margin-top: 5px;'>Выпущено: {format_number(most_expensive_batch['Количество_выпущено'])} шт.</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if not expensive_materials.empty:
                        st.markdown("**Состав сырья:**")
                        expensive_display = expensive_materials[['Сырье', 'Себестоимость_на_единицу_продукции']].copy()
                        expensive_display['Доля_в_себестоимости'] = expensive_display['Себестоимость_на_единицу_продукции'] / most_expensive_batch['Себестоимость_единицы'] * 100
                        expensive_display = expensive_display.sort_values('Себестоимость_на_единицу_продукции', ascending=False)
                        for _, row in expensive_display.iterrows():
                            st.write(f"• {row['Сырье']}: {format_number(row['Себестоимость_на_единицу_продукции'])} ₽/шт. ({format_float(row['Доля_в_себестоимости'], 1)}%)")
                
                diff_cost = most_expensive_batch['Себестоимость_единицы'] - cheapest_batch['Себестоимость_единицы']
                diff_percent = (diff_cost / cheapest_batch['Себестоимость_единицы'] * 100)
                
                st.markdown(f"""
                <div style='
                    background-color: #E2E3E5;
                    border-radius: 10px;
                    padding: 10px;
                    text-align: center;
                    margin-top: 10px;
                '>
                    <span style='font-size: 16px;'>📊 Разница между самой дорогой и самой дешевой партией:</span>
                    <span style='font-size: 20px; font-weight: bold; color: #DC3545;'>{format_number(diff_cost)} ₽/шт.</span>
                    <span style='font-size: 16px;'>({format_float(diff_percent, 1)}% дороже)</span>
                </div>
                """, unsafe_allow_html=True)
                
                st.divider()
            
            st.subheader("📈 ДИНАМИКА СЕБЕСТОИМОСТИ ПО ПАРТИЯМ")
            
            batches_sorted = batches.copy()
            batches_sorted['Партия_стр'] = batches_sorted['Партия'].astype(str)
            
            if sort_by == "По номеру (старые сверху)":
                batches_sorted = batches_sorted.sort_values('Партия_стр', ascending=True)
            elif sort_by == "По номеру (новые сверху)":
                batches_sorted = batches_sorted.sort_values('Партия_стр', ascending=False)
            elif sort_by == "По себестоимости (от низкой)":
                batches_sorted = batches_sorted.sort_values('Себестоимость_единицы', ascending=True)
            elif sort_by == "По себестоимости (от высокой)":
                batches_sorted = batches_sorted.sort_values('Себестоимость_единицы', ascending=False)
            
            fig_cost = go.Figure()
            fig_cost.add_trace(go.Scatter(
                x=batches_sorted['Партия_стр'],
                y=batches_sorted['Себестоимость_единицы'],
                mode='lines+markers',
                name='Себестоимость единицы',
                line=dict(color='#2E86AB', width=2),
                marker=dict(size=8, color='#1A5276')
            ))
            
            fig_cost.add_hline(
                y=avg_cost,
                line_dash="dash",
                line_color="red",
                annotation_text=f"Среднее: {format_number(avg_cost)} ₽/шт.",
                annotation_position="top right"
            )
            
            fig_cost.update_layout(
                title="Динамика себестоимости единицы продукции по партиям",
                xaxis_title="Партия",
                yaxis_title="Себестоимость (₽/шт.)",
                hovermode='x unified',
                height=450,
                xaxis_tickangle=-45
            )
            
            st.plotly_chart(fig_cost, use_container_width=True)
            
            st.divider()
            
            st.subheader("🔍 ДЕТАЛЬНЫЙ СОСТАВ СЕБЕСТОИМОСТИ")
            
            if selected_batch == 'Все партии':
                st.write("### 📋 Все партии")
                display_batches = batches[['Партия', 'Количество_выпущено', 'Себестоимость_единицы']].copy()
                display_batches['Партия'] = display_batches['Партия'].astype(str)
                display_batches['Себестоимость_единицы'] = display_batches['Себестоимость_единицы'].apply(lambda x: f"{format_number(x)} ₽")
                display_batches['Количество_выпущено'] = display_batches['Количество_выпущено'].apply(format_number)
                st.dataframe(display_batches, use_container_width=True, hide_index=True)
                
                st.subheader("📊 Распределение себестоимости по партиям")
                fig_hist = px.histogram(batches, x='Себестоимость_единицы', nbins=20,
                                         title='Гистограмма распределения себестоимости',
                                         labels={'Себестоимость_единицы': 'Себестоимость (₽/шт.)'})
                st.plotly_chart(fig_hist, use_container_width=True)
                
            else:
                batch_info = batches[batches['Партия'].astype(str) == selected_batch]
                if batch_info.empty:
                    st.warning(f"Партия {selected_batch} не найдена")
                else:
                    batch_info = batch_info.iloc[0]
                    batch_materials = materials[materials['Партия'].astype(str) == selected_batch]
                    
                    st.write(f"### 📦 Партия: {selected_batch}")
                    
                    col_info1, col_info2, col_info3 = st.columns(3)
                    with col_info1:
                        st.metric("📦 Выпущено", f"{format_number(batch_info['Количество_выпущено'])} шт.")
                    with col_info2:
                        st.metric("💰 Себестоимость единицы", f"{format_number(batch_info['Себестоимость_единицы'])} ₽/шт.")
                    with col_info3:
                        total_cost = batch_info['Себестоимость_единицы'] * batch_info['Количество_выпущено']
                        st.metric("💵 Общая себестоимость партии", f"{format_number(total_cost)} ₽")
                    
                    st.divider()
                    
                    if not batch_materials.empty:
                        st.write("### 📋 Состав себестоимости партии")
                        
                        batch_materials['Доля_в_себестоимости'] = batch_materials['Себестоимость_на_единицу_продукции'] / batch_info['Себестоимость_единицы'] * 100
                        
                        display_materials = batch_materials[['Сырье', 'Количество_сырья', 'Цена_сырья', 'Себестоимость_на_единицу_продукции', 'Доля_в_себестоимости']].copy()
                        display_materials['Цена_сырья'] = display_materials['Цена_сырья'].apply(lambda x: f"{format_number(x)} ₽")
                        display_materials['Количество_сырья'] = display_materials['Количество_сырья'].apply(format_number)
                        display_materials['Себестоимость_на_единицу_продукции'] = display_materials['Себестоимость_на_единицу_продукции'].apply(lambda x: f"{format_number(x)} ₽")
                        display_materials['Доля_в_себестоимости'] = display_materials['Доля_в_себестоимости'].apply(lambda x: f"{format_float(x, 1)}%")
                        
                        st.dataframe(display_materials, use_container_width=True, hide_index=True)
                        
                        st.subheader("🥧 Структура себестоимости")
                        fig_pie = px.pie(batch_materials, values='Себестоимость_на_единицу_продукции', names='Сырье',
                                         title=f'Распределение затрат в партии {selected_batch}',
                                         labels={'Себестоимость_на_единицу_продукции': 'Затраты (₽/шт.)'})
                        st.plotly_chart(fig_pie, use_container_width=True)
                        
                        st.subheader("📊 Анализ сырья")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            fig_bar = px.bar(batch_materials, x='Сырье', y='Цена_сырья',
                                             title='Цена сырья',
                                             labels={'Цена_сырья': 'Цена (₽)', 'Сырье': ''})
                            fig_bar.update_layout(xaxis_tickangle=-45)
                            st.plotly_chart(fig_bar, use_container_width=True)
                        
                        with col2:
                            fig_qty = px.bar(batch_materials, x='Сырье', y='Количество_сырья',
                                             title='Количество использованного сырья',
                                             labels={'Количество_сырья': 'Количество', 'Сырье': ''})
                            fig_qty.update_layout(xaxis_tickangle=-45)
                            st.plotly_chart(fig_qty, use_container_width=True)
                        
                    else:
                        st.info("Нет данных о составе сырья для этой партии")
        
        st.divider()
        
        if 'batches' in locals() and not batches.empty and len(batches) >= 2:
            st.subheader("🔄 СРАВНЕНИЕ ПАРТИЙ")
            
            col_comp1, col_comp2 = st.columns(2)
            
            batch_list = sorted(batches['Партия'].astype(str).unique().tolist())
            
            with col_comp1:
                batch1 = st.selectbox("Выберите первую партию", batch_list, key="comp_batch1")
            
            with col_comp2:
                batch2 = st.selectbox("Выберите вторую партию", batch_list, key="comp_batch2")
            
            if batch1 and batch2 and batch1 != batch2:
                batch1_info = batches[batches['Партия'].astype(str) == batch1].iloc[0]
                batch2_info = batches[batches['Партия'].astype(str) == batch2].iloc[0]
                
                materials1 = materials[materials['Партия'].astype(str) == batch1] if 'materials' in locals() else pd.DataFrame()
                materials2 = materials[materials['Партия'].astype(str) == batch2] if 'materials' in locals() else pd.DataFrame()
                
                st.write(f"### Сравнение партий: {batch1} vs {batch2}")
                
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    diff_cost = batch2_info['Себестоимость_единицы'] - batch1_info['Себестоимость_единицы']
                    st.metric("Разница в себестоимости", f"{format_number(diff_cost)} ₽/шт.",
                             delta=f"{format_float(diff_cost / batch1_info['Себестоимость_единицы'] * 100, 1)}%")
                with c2:
                    diff_qty = batch2_info['Количество_выпущено'] - batch1_info['Количество_выпущено']
                    st.metric("Разница в выпуске", f"{format_number(diff_qty)} шт.")
                
                if not materials1.empty and not materials2.empty:
                    st.write("### Сравнение состава сырья")
                    
                    comparison = materials1[['Сырье', 'Себестоимость_на_единицу_продукции']].merge(
                        materials2[['Сырье', 'Себестоимость_на_единицу_продукции']],
                        on='Сырье',
                        how='outer',
                        suffixes=('_1', '_2')
                    ).fillna(0)
                    
                    comparison['Разница'] = comparison['Себестоимость_на_единицу_продукции_2'] - comparison['Себестоимость_на_единицу_продукции_1']
                    
                    display_comp = comparison.copy()
                    for col in ['Себестоимость_на_единицу_продукции_1', 'Себестоимость_на_единицу_продукции_2', 'Разница']:
                        display_comp[col] = display_comp[col].apply(lambda x: f"{format_number(x)} ₽")
                    
                    st.dataframe(display_comp, use_container_width=True, hide_index=True)
                    
                    fig_comp = go.Figure()
                    fig_comp.add_trace(go.Bar(name=batch1, x=comparison['Сырье'], y=comparison['Себестоимость_на_единицу_продукции_1'], marker_color='#2E86AB'))
                    fig_comp.add_trace(go.Bar(name=batch2, x=comparison['Сырье'], y=comparison['Себестоимость_на_единицу_продукции_2'], marker_color='#D9534F'))
                    fig_comp.update_layout(title='Сравнение затрат на сырьё по партиям',
                                           xaxis_title='Сырьё',
                                           yaxis_title='Затраты (₽/шт.)',
                                           barmode='group',
                                           xaxis_tickangle=-45)
                    st.plotly_chart(fig_comp, use_container_width=True)

# ==========================================
# СТРАНИЦА 5: ЛОГИСТИКА UPDATE
# ==========================================
elif page == "🚚 Логистика Update":
    st.title("🚚 Аналитика логистики (обновленная)")
    
    if logistics_df.empty:
        st.warning("⚠️ Файл 'logistics_data.xlsx' не найден или не удалось загрузить данные.")
        st.info("📌 Пожалуйста, добавьте файл с данными по логистике в папку с приложением.")
    else:
        st.markdown("### Анализ логистических затрат по заказам")
        
        # ===== 1. ПОДГОТОВКА ДАННЫХ =====
        df_log_upd = logistics_df.copy()
        
        # Фильтруем только строки для анализа
        if 'Принимать к анализу?' in df_log_upd.columns:
            df_log_upd = df_log_upd[df_log_upd['Принимать к анализу?'] == 1]
        
        # Определяем дату (столбец K - Дата отгрузки)
        if 'Дата отгрузки' in df_log_upd.columns:
            df_log_upd['Дата'] = pd.to_datetime(df_log_upd['Дата отгрузки'], errors='coerce')
        else:
            st.error("❌ Столбец 'Дата отгрузки' не найден!")
            st.stop()
        
        # Удаляем строки с некорректными датами
        df_log_upd = df_log_upd.dropna(subset=['Дата'])
        
        if df_log_upd.empty:
            st.warning("⚠️ Нет данных с корректными датами")
            st.stop()
        
        # ===== 2. КЛЮЧЕВЫЕ КОЛОНКИ =====
        # Столбец I (индекс 8) - количество паллет
        if len(df_log_upd.columns) > 8:
            pallet_col = df_log_upd.columns[8]
            df_log_upd['Кол-во паллет_сырое'] = pd.to_numeric(df_log_upd[pallet_col], errors='coerce').fillna(0)
        else:
            df_log_upd['Кол-во паллет_сырое'] = 0
        
        # Столбец J (индекс 9) - Город
        if len(df_log_upd.columns) > 9:
            city_col = df_log_upd.columns[9]
            df_log_upd['Город'] = df_log_upd[city_col]
        else:
            df_log_upd['Город'] = 'Не указан'
        
        # Столбец AB (индекс 27) - Затраты на логистику
        if len(df_log_upd.columns) > 27:
            ab_col = df_log_upd.columns[27]
            df_log_upd['Логистика_затраты_сырые'] = pd.to_numeric(df_log_upd[ab_col], errors='coerce').fillna(0)
        else:
            st.error(f"❌ Столбец AB не найден (всего колонок: {len(df_log_upd.columns)})")
            st.stop()
        
        # Столбец Y - Стоимость товара в заказе Без НДС
        if 'Стоимость товара в заказе Без НДС' in df_log_upd.columns:
            df_log_upd['Стоимость_товара_сырая'] = pd.to_numeric(df_log_upd['Стоимость товара в заказе Без НДС'], errors='coerce').fillna(0)
        else:
            df_log_upd['Стоимость_товара_сырая'] = 0
        
        # Номенклатура для детализации
        if 'Номенклатура' in df_log_upd.columns:
            df_log_upd['Номенклатура'] = df_log_upd['Номенклатура']
        
        # ===== 3. ПЕРВЫЙ УРОВЕНЬ АГРЕГАЦИИ: (ДАТА + ГОРОД) =====
        df_city_order = df_log_upd.groupby(['Дата', 'Город']).agg({
            'Кол-во паллет_сырое': 'first',
            'Логистика_затраты_сырые': 'first',
            'Стоимость_товара_сырая': 'sum',
            'Номенклатура': lambda x: list(x)
        }).reset_index()
        
        df_city_order = df_city_order.rename(columns={
            'Кол-во паллет_сырое': 'Кол-во паллет_по_городу',
            'Логистика_затраты_сырые': 'Логистика_затраты_по_городу',
            'Стоимость_товара_сырая': 'Стоимость_товара_по_городу',
            'Номенклатура': 'Список_номенклатур'
        })
        
        # ===== 4. ВТОРОЙ УРОВЕНЬ АГРЕГАЦИИ: ПО ЗАКАЗАМ (ДАТА) =====
        df_orders = df_city_order.groupby('Дата').agg({
            'Город': lambda x: list(x),
            'Кол-во паллет_по_городу': 'sum',
            'Логистика_затраты_по_городу': 'sum',
            'Стоимость_товара_по_городу': 'sum',
            'Список_номенклатур': lambda x: [item for sublist in x for item in sublist]
        }).reset_index()
        
        df_orders = df_orders.rename(columns={
            'Кол-во паллет_по_городу': 'Кол-во паллет',
            'Логистика_затраты_по_городу': 'Логистика_затраты',
            'Стоимость_товара_по_городу': 'Стоимость_товара'
        })
        
        # Добавляем количество позиций в заказе
        df_orders['Кол-во_позиций'] = df_orders['Список_номенклатур'].apply(len)
        
        # Добавляем количество УНИКАЛЬНЫХ городов в заказе
        df_orders['Кол-во_городов_в_заказе'] = df_orders['Город'].apply(lambda x: len(set(x)))
        
        # Создаём временные признаки
        df_orders['Год'] = df_orders['Дата'].dt.year
        df_orders['Месяц'] = df_orders['Дата'].dt.month
        df_orders['Месяц_название'] = df_orders['Месяц'].map(month_names)
        df_orders['Месяц_год'] = df_orders['Дата'].dt.strftime('%Y-%m')
        df_orders['День'] = df_orders['Дата'].dt.day
        
        # Для удобства, создаём колонку с первым городом (для фильтрации)
        df_orders['Город_первый'] = df_orders['Город'].apply(lambda x: x[0] if x else 'Не указан')
        
        # Также создаём детальную таблицу (Дата + Город) для анализа по городам внутри заказов
        df_city_detail = df_city_order.copy()
        df_city_detail['Год'] = df_city_detail['Дата'].dt.year
        df_city_detail['Месяц'] = df_city_detail['Дата'].dt.month
        df_city_detail['Месяц_название'] = df_city_detail['Месяц'].map(month_names)
        
        # ===== 5. УРОВЕНЬ 1: АНАЛИЗ ПО МЕСЯЦАМ =====
        st.header("📊 УРОВЕНЬ 1: АНАЛИЗ ПО МЕСЯЦАМ")
        st.markdown("Агрегированные данные по всем заказам за месяц")
        
        # Фильтры для уровня 1
        col_f1, col_f2 = st.columns(2)
        
        with col_f1:
            available_years = sorted(df_orders['Год'].dropna().unique())
            if len(available_years) == 0:
                available_years = [2024]
            selected_year_l1 = st.selectbox("📅 Выберите год", available_years, key="l1_year")
        
        with col_f2:
            all_cities_l1 = sorted(df_city_detail['Город'].dropna().unique().tolist())
            selected_city_l1 = st.selectbox("🏙️ Город (опционально)", ['Все города'] + all_cities_l1, key="l1_city")
        
        # Применяем фильтры к заказам
        df_l1_orders = df_orders[df_orders['Год'] == selected_year_l1]
        
        if selected_city_l1 != 'Все города':
            df_l1_orders = df_l1_orders[df_l1_orders['Город'].apply(lambda x: selected_city_l1 in x)]
        
        if df_l1_orders.empty:
            st.warning("⚠️ Нет данных для выбранных фильтров")
        else:
            # Агрегация по месяцам
            monthly_l1 = df_l1_orders.groupby('Месяц_название').agg({
                'Дата': 'count',  # количество заказов
                'Стоимость_товара': 'sum',
                'Логистика_затраты': 'sum',
                'Кол-во паллет': 'sum',
                'Кол-во_позиций': 'sum',
                'Город': lambda x: list(set([item for sublist in x for item in sublist]))  # УНИКАЛЬНЫЕ города за месяц
            }).reset_index()
            monthly_l1 = monthly_l1.rename(columns={
                'Дата': 'Количество_заказов',
                'Город': 'Уникальные_города'
            })
            
            # Добавляем количество уникальных городов
            monthly_l1['Кол-во_уникальных_городов'] = monthly_l1['Уникальные_города'].apply(len)
            
            # Сортируем по порядку месяцев
            month_order = [month_names[m] for m in sorted(df_l1_orders['Месяц'].unique())]
            monthly_l1['Месяц_порядок'] = pd.Categorical(monthly_l1['Месяц_название'], categories=month_order, ordered=True)
            monthly_l1 = monthly_l1.sort_values('Месяц_порядок')
            
            # Рассчитываем долю логистики
            monthly_l1['Доля_логистики_%'] = (monthly_l1['Логистика_затраты'] / monthly_l1['Стоимость_товара'].replace(0, 1) * 100)
            monthly_l1['Средняя_стоимость_паллеты'] = monthly_l1['Логистика_затраты'] / monthly_l1['Кол-во паллет'].replace(0, 1)
            monthly_l1['Средняя_логистика_на_заказ'] = monthly_l1['Логистика_затраты'] / monthly_l1['Количество_заказов'].replace(0, 1)
            
            # ===== МЕТРИКИ ПО МЕСЯЦАМ =====
            st.subheader(f"📊 СВОДКА ЗА {selected_year_l1} ГОД")
            
            total_orders_l1 = monthly_l1['Количество_заказов'].sum()
            total_goods_l1 = monthly_l1['Стоимость_товара'].sum()
            total_logistics_l1 = monthly_l1['Логистика_затраты'].sum()
            total_pallets_l1 = monthly_l1['Кол-во паллет'].sum()
            total_items_l1 = monthly_l1['Кол-во_позиций'].sum()
            
            # УНИКАЛЬНЫЕ города за весь год
            all_cities_year = []
            for cities_list in monthly_l1['Уникальные_города']:
                all_cities_year.extend(cities_list)
            unique_cities_year = len(set(all_cities_year))
            
            avg_logistics_share = (total_logistics_l1 / total_goods_l1 * 100) if total_goods_l1 > 0 else 0
            avg_pallet_cost = total_logistics_l1 / total_pallets_l1 if total_pallets_l1 > 0 else 0
            avg_logistics_per_order = total_logistics_l1 / total_orders_l1 if total_orders_l1 > 0 else 0
            
            col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
            with col_m1:
                st.metric("📋 Всего заказов", format_number(total_orders_l1))
            with col_m2:
                st.metric("🏷️ Стоимость товаров", f"{format_number(total_goods_l1)} ₽")
            with col_m3:
                st.metric("💰 Затраты на логистику", f"{format_number(total_logistics_l1)} ₽")
            with col_m4:
                st.metric("📊 Доля логистики", f"{format_float(avg_logistics_share, 1)}%")
            with col_m5:
                st.metric("📦 Средняя цена паллеты", f"{format_number(avg_pallet_cost)} ₽")
            
            col_m6, col_m7, col_m8 = st.columns(3)
            with col_m6:
                st.metric("📦 Всего паллет", format_number(total_pallets_l1))
            with col_m7:
                st.metric("🏙️ Уникальных городов", format_number(unique_cities_year))
            with col_m8:
                st.metric("💰 Средние затраты на заказ", f"{format_number(avg_logistics_per_order)} ₽")
            
            st.divider()
            
            # ===== ТАБЛИЦА ПО МЕСЯЦАМ =====
            st.subheader("📅 ПОМЕСЯЧНАЯ РАЗБИВКА")
            
            display_l1 = monthly_l1.copy()
            display_l1['Стоимость_товара'] = display_l1['Стоимость_товара'].apply(lambda x: f"{format_number(x)} ₽")
            display_l1['Логистика_затраты'] = display_l1['Логистика_затраты'].apply(lambda x: f"{format_number(x)} ₽")
            display_l1['Доля_логистики_%'] = display_l1['Доля_логистики_%'].apply(lambda x: f"{format_float(x, 1)}%")
            display_l1['Количество_заказов'] = display_l1['Количество_заказов'].apply(lambda x: format_number(x))
            display_l1['Кол-во паллет'] = display_l1['Кол-во паллет'].apply(lambda x: format_number(x))
            display_l1['Кол-во_уникальных_городов'] = display_l1['Кол-во_уникальных_городов'].apply(lambda x: format_number(x))
            display_l1['Средняя_стоимость_паллеты'] = display_l1['Средняя_стоимость_паллеты'].apply(lambda x: f"{format_number(x)} ₽")
            display_l1['Средняя_логистика_на_заказ'] = display_l1['Средняя_логистика_на_заказ'].apply(lambda x: f"{format_number(x)} ₽")
            display_l1['Уникальные_города'] = display_l1['Уникальные_города'].apply(lambda x: ', '.join(x))
            
            st.dataframe(
                display_l1[['Месяц_название', 'Количество_заказов', 'Кол-во паллет', 'Кол-во_уникальных_городов',
                           'Стоимость_товара', 'Логистика_затраты', 'Доля_логистики_%', 
                           'Средняя_стоимость_паллеты', 'Средняя_логистика_на_заказ', 'Уникальные_города']],
                use_container_width=True,
                hide_index=True
            )
            
            # ===== ГРАФИКИ ПО МЕСЯЦАМ =====
            col_g1, col_g2 = st.columns(2)
            
            with col_g1:
                fig1 = go.Figure()
                fig1.add_trace(go.Bar(
                    x=monthly_l1['Месяц_название'],
                    y=monthly_l1['Стоимость_товара'],
                    name='Стоимость товаров',
                    marker_color='#2E86AB',
                    text=monthly_l1['Стоимость_товара'].apply(lambda x: f'{format_number(x)} ₽'),
                    textposition='outside'
                ))
                fig1.add_trace(go.Bar(
                    x=monthly_l1['Месяц_название'],
                    y=monthly_l1['Логистика_затраты'],
                    name='Затраты на логистику',
                    marker_color='#D9534F',
                    text=monthly_l1['Логистика_затраты'].apply(lambda x: f'{format_number(x)} ₽'),
                    textposition='outside'
                ))
                fig1.update_layout(
                    title='Стоимость товаров vs Затраты на логистику по месяцам',
                    xaxis_title='Месяц',
                    yaxis_title='Сумма (₽)',
                    barmode='group',
                    height=400
                )
                st.plotly_chart(fig1, use_container_width=True)
            
            with col_g2:
                fig2 = go.Figure()
                fig2.add_trace(go.Scatter(
                    x=monthly_l1['Месяц_название'],
                    y=monthly_l1['Кол-во_уникальных_городов'],
                    mode='lines+markers',
                    name='Уникальные города',
                    line=dict(color='#52B788', width=3),
                    marker=dict(size=10, color='#2D6A4F')
                ))
                fig2.add_trace(go.Scatter(
                    x=monthly_l1['Месяц_название'],
                    y=monthly_l1['Количество_заказов'],
                    mode='lines+markers',
                    name='Количество заказов',
                    line=dict(color='#FFB74D', width=2, dash='dash'),
                    marker=dict(size=8, color='#F57C00'),
                    yaxis='y2'
                ))
                fig2.update_layout(
                    title='Динамика уникальных городов и количества заказов',
                    xaxis_title='Месяц',
                    yaxis_title='Количество уникальных городов',
                    yaxis2=dict(
                        title='Количество заказов',
                        overlaying='y',
                        side='right'
                    ),
                    height=400,
                    hovermode='x unified'
                )
                st.plotly_chart(fig2, use_container_width=True)
            
            # ===== ДОПОЛНИТЕЛЬНЫЙ ГРАФИК: ДОЛЯ ЛОГИСТИКИ =====
            fig3 = go.Figure()
            fig3.add_trace(go.Bar(
                x=monthly_l1['Месяц_название'],
                y=monthly_l1['Доля_логистики_%'],
                name='Доля логистики',
                marker_color='#7B68EE',
                text=monthly_l1['Доля_логистики_%'].apply(lambda x: f'{format_float(x, 1)}%'),
                textposition='outside'
            ))
            fig3.update_layout(
                title='Доля логистики в стоимости товаров по месяцам',
                xaxis_title='Месяц',
                yaxis_title='Доля (%)',
                height=400
            )
            st.plotly_chart(fig3, use_container_width=True)
            
            st.divider()
        
        # ===== 6. УРОВЕНЬ 2: АНАЛИЗ ПО ЗАКАЗАМ =====
        st.header("📋 УРОВЕНЬ 2: АНАЛИЗ ПО ЗАКАЗАМ")
        st.markdown("Детальный анализ по каждому заказу (дате отгрузки)")
        
        # Фильтры для уровня 2
        col_f3, col_f4, col_f5 = st.columns(3)
        
        with col_f3:
            available_years_l2 = sorted(df_orders['Год'].dropna().unique())
            if len(available_years_l2) == 0:
                available_years_l2 = [2024]
            selected_year_l2 = st.selectbox("📅 Год", available_years_l2, key="l2_year")
        
        with col_f4:
            df_year_l2 = df_orders[df_orders['Год'] == selected_year_l2]
            available_months_l2 = sorted(df_year_l2['Месяц'].dropna().unique())
            available_months_display_l2 = [month_names[m] for m in available_months_l2]
            if available_months_display_l2:
                selected_month_display_l2 = st.selectbox("📅 Месяц", available_months_display_l2, key="l2_month")
                selected_month_l2 = available_months_l2[available_months_display_l2.index(selected_month_display_l2)]
            else:
                st.warning("Нет данных за выбранный год")
                st.stop()
        
        with col_f5:
            all_cities_l2 = sorted(df_city_detail['Город'].dropna().unique().tolist())
            selected_city_l2 = st.selectbox("🏙️ Город", ['Все города'] + all_cities_l2, key="l2_city")
        
        # Применяем фильтры к заказам
        df_l2 = df_orders.copy()
        df_l2 = df_l2[df_l2['Год'] == selected_year_l2]
        df_l2 = df_l2[df_l2['Месяц'] == selected_month_l2]
        
        if selected_city_l2 != 'Все города':
            df_l2 = df_l2[df_l2['Город'].apply(lambda x: selected_city_l2 in x)]
        
        if df_l2.empty:
            st.warning("⚠️ Нет данных для выбранных фильтров")
        else:
            # Сортируем по дате
            df_l2 = df_l2.sort_values('Дата', ascending=False)
            
            # ===== МЕТРИКИ ПО ЗАКАЗАМ =====
            total_orders_l2 = len(df_l2)
            total_goods_l2 = df_l2['Стоимость_товара'].sum()
            total_logistics_l2 = df_l2['Логистика_затраты'].sum()
            total_pallets_l2 = df_l2['Кол-во паллет'].sum()
            total_items_l2 = df_l2['Кол-во_позиций'].sum()
            
            # Уникальные города в выбранных заказах
            all_cities_selected = []
            for cities in df_l2['Город']:
                all_cities_selected.extend(cities)
            unique_cities_selected = len(set(all_cities_selected))
            
            avg_logistics_share_l2 = (total_logistics_l2 / total_goods_l2 * 100) if total_goods_l2 > 0 else 0
            
            st.subheader(f"📊 СВОДКА ЗА {selected_month_display_l2} {selected_year_l2}")
            
            col_m6, col_m7, col_m8, col_m9, col_m10 = st.columns(5)
            with col_m6:
                st.metric("📋 Заказов", format_number(total_orders_l2))
            with col_m7:
                st.metric("🏷️ Товары", f"{format_number(total_goods_l2)} ₽")
            with col_m8:
                st.metric("💰 Логистика", f"{format_number(total_logistics_l2)} ₽")
            with col_m9:
                st.metric("📊 Доля логистики", f"{format_float(avg_logistics_share_l2, 1)}%")
            with col_m10:
                st.metric("📦 Паллет", format_number(total_pallets_l2))
            
            col_m11, col_m12 = st.columns(2)
            with col_m11:
                st.metric("🏙️ Уникальных городов", format_number(unique_cities_selected))
            with col_m12:
                st.metric("💰 Средние затраты на заказ", f"{format_number(total_logistics_l2 / total_orders_l2 if total_orders_l2 > 0 else 0)} ₽")
            
            st.divider()
            
            # ===== ДЕТАЛЬНАЯ ТАБЛИЦА ПО ЗАКАЗАМ =====
            st.subheader("📋 ДЕТАЛИЗАЦИЯ ПО ЗАКАЗАМ")
            
            display_l2 = df_l2.copy()
            display_l2['Дата_отгрузки'] = display_l2['Дата'].dt.strftime('%d.%m.%Y')
            display_l2['Города'] = display_l2['Город'].apply(lambda x: ', '.join(set(x)))
            display_l2['Кол-во_городов_в_заказе'] = display_l2['Кол-во_городов_в_заказе'].apply(lambda x: format_number(x))
            display_l2['Стоимость_товара'] = display_l2['Стоимость_товара'].apply(lambda x: f"{format_number(x)} ₽")
            display_l2['Логистика_затраты'] = display_l2['Логистика_затраты'].apply(lambda x: f"{format_number(x)} ₽")
            display_l2['Доля_логистики_%'] = (df_l2['Логистика_затраты'] / df_l2['Стоимость_товара'].replace(0, 1) * 100)
            display_l2['Доля_логистики_%'] = display_l2['Доля_логистики_%'].apply(lambda x: f"{format_float(x, 1)}%")
            display_l2['Кол-во паллет'] = display_l2['Кол-во паллет'].apply(lambda x: format_number(x))
            display_l2['Кол-во_позиций'] = display_l2['Кол-во_позиций'].apply(lambda x: format_number(x))
            
            display_l2['Номенклатуры'] = display_l2['Список_номенклатур'].apply(
                lambda x: ', '.join(x[:3]) + ('...' if len(x) > 3 else '')
            )
            
            st.dataframe(
                display_l2[['Дата_отгрузки', 'Города', 'Кол-во паллет', 'Кол-во_позиций', 'Кол-во_городов_в_заказе',
                           'Стоимость_товара', 'Логистика_затраты', 'Доля_логистики_%', 'Номенклатуры']],
                use_container_width=True,
                hide_index=True
            )
            
            # ===== ДЕТАЛЬНАЯ ТАБЛИЦА ПО ГОРОДАМ В ЗАКАЗЕ =====
            with st.expander("🏙️ Детализация по городам в заказах"):
                order_dates = df_l2['Дата'].tolist()
                df_city_detail_filtered = df_city_detail[df_city_detail['Дата'].isin(order_dates)]
                
                if selected_city_l2 != 'Все города':
                    df_city_detail_filtered = df_city_detail_filtered[df_city_detail_filtered['Город'] == selected_city_l2]
                
                display_cities = df_city_detail_filtered.copy()
                display_cities['Дата_отгрузки'] = display_cities['Дата'].dt.strftime('%d.%m.%Y')
                display_cities['Стоимость_товара_по_городу'] = display_cities['Стоимость_товара_по_городу'].apply(lambda x: f"{format_number(x)} ₽")
                display_cities['Логистика_затраты_по_городу'] = display_cities['Логистика_затраты_по_городу'].apply(lambda x: f"{format_number(x)} ₽")
                display_cities['Кол-во паллет_по_городу'] = display_cities['Кол-во паллет_по_городу'].apply(lambda x: format_number(x))
                
                st.dataframe(
                    display_cities[['Дата_отгрузки', 'Город', 'Кол-во паллет_по_городу', 
                                   'Стоимость_товара_по_городу', 'Логистика_затраты_по_городу']],
                    use_container_width=True,
                    hide_index=True
                )
            
            # ===== ЭКСПОРТ =====
            st.subheader("📥 ЭКСПОРТ ДАННЫХ")
            
            col_export1, col_export2 = st.columns(2)
            
            with col_export1:
                export_orders = df_l2.copy()
                export_orders['Дата'] = export_orders['Дата'].dt.strftime('%Y-%m-%d')
                export_orders['Города'] = export_orders['Город'].apply(lambda x: ', '.join(set(x)))
                export_orders['Список_номенклатур'] = export_orders['Список_номенклатур'].apply(lambda x: ', '.join(x))
                
                csv_orders = export_orders[['Дата', 'Города', 'Кол-во паллет', 'Кол-во_позиций', 
                                           'Стоимость_товара', 'Логистика_затраты']].to_csv(
                    index=False, sep=';', decimal=','
                ).encode('utf-8-sig')
                
                st.download_button(
                    "📥 Скачать детализацию по заказам (CSV)",
                    csv_orders,
                    f"logistics_orders_{selected_year_l2}_{selected_month_l2}.csv",
                    "text/csv",
                    use_container_width=True
                )
            
            with col_export2:
                order_dates = df_l2['Дата'].tolist()
                df_city_export = df_city_detail[df_city_detail['Дата'].isin(order_dates)]
                
                if selected_city_l2 != 'Все города':
                    df_city_export = df_city_export[df_city_export['Город'] == selected_city_l2]
                
                df_city_export['Дата'] = df_city_export['Дата'].dt.strftime('%Y-%m-%d')
                csv_cities = df_city_export[['Дата', 'Город', 'Кол-во паллет_по_городу', 
                                            'Стоимость_товара_по_городу', 'Логистика_затраты_по_городу']].to_csv(
                    index=False, sep=';', decimal=','
                ).encode('utf-8-sig')
                
                st.download_button(
                    "📥 Скачать детализацию по городам (CSV)",
                    csv_cities,
                    f"logistics_cities_{selected_year_l2}_{selected_month_l2}.csv",
                    "text/csv",
                    use_container_width=True
                )
            
            # ===== ГРАФИКИ ПО ЗАКАЗАМ =====
            st.subheader("📊 ВИЗУАЛИЗАЦИЯ ПО ЗАКАЗАМ")
            
            col_g3, col_g4 = st.columns(2)
            
            with col_g3:
                top_orders = df_l2.nlargest(10, 'Логистика_затраты')
                fig4 = px.bar(
                    top_orders,
                    x='Дата',
                    y='Логистика_затраты',
                    title='Топ-10 заказов по затратам на логистику',
                    labels={'Дата': 'Дата отгрузки', 'Логистика_затраты': 'Затраты (₽)'},
                    color='Кол-во_городов_в_заказе',
                    color_continuous_scale='Viridis',
                    text=top_orders['Кол-во_городов_в_заказе'].apply(lambda x: f'{x} городов')
                )
                fig4.update_layout(height=400, xaxis_tickangle=-45)
                st.plotly_chart(fig4, use_container_width=True)
            
            with col_g4:
                city_stats = df_city_detail_filtered.groupby('Город').agg({
                    'Кол-во паллет_по_городу': 'sum',
                    'Логистика_затраты_по_городу': 'sum',
                    'Стоимость_товара_по_городу': 'sum'
                }).reset_index()
                city_stats = city_stats.sort_values('Логистика_затраты_по_городу', ascending=False)
                
                fig5 = px.bar(
                    city_stats,
                    x='Город',
                    y='Логистика_затраты_по_городу',
                    title='Затраты на логистику по городам',
                    labels={'Логистика_затраты_по_городу': 'Затраты (₽)', 'Город': ''},
                    color='Кол-во паллет_по_городу',
                    color_continuous_scale='Blues',
                    text=city_stats['Логистика_затраты_по_городу'].apply(lambda x: format_number(x))
                )
                fig5.update_layout(height=400)
                st.plotly_chart(fig5, use_container_width=True)
            
            # ===== ДОПОЛНИТЕЛЬНЫЙ ГРАФИК =====
            fig6 = go.Figure()
            fig6.add_trace(go.Scatter(
                x=df_l2['Стоимость_товара'],
                y=df_l2['Логистика_затраты'],
                mode='markers',
                marker=dict(
                    size=df_l2['Кол-во_городов_в_заказе'] * 8 + 5,
                    color=df_l2['Кол-во_городов_в_заказе'],
                    colorscale='Viridis',
                    showscale=True,
                    colorbar=dict(title='Кол-во городов'),
                    sizemode='diameter'
                ),
                text=df_l2['Дата'].dt.strftime('%d.%m.%Y') + ' - ' + df_l2['Город'].apply(lambda x: ', '.join(set(x))),
                hoverinfo='text+x+y'
            ))
            fig6.update_layout(
                title='Стоимость товаров vs Затраты на логистику по заказам',
                xaxis_title='Стоимость товаров (₽)',
                yaxis_title='Затраты на логистику (₽)',
                height=450
            )
            st.plotly_chart(fig6, use_container_width=True)

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
