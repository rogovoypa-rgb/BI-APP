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
    try:
        df = pd.read_excel('sales_data.xlsx')
        st.success("✅ sales_data.xlsx загружен")
        return df
    except Exception as e:
        st.error(f"Ошибка загрузки sales_data.xlsx: {e}")
        return pd.DataFrame()

@st.cache_data
def load_logistics_data():
    try:
        df = pd.read_excel('logistics_data.xlsx', header=0)
        st.success("✅ logistics_data.xlsx загружен")
        return df
    except Exception as e:
        st.error(f"Ошибка загрузки logistics_data.xlsx: {e}")
        return pd.DataFrame()

@st.cache_data
def load_production_data():
    try:
        df_raw = pd.read_excel('production_data.xlsx', header=None)
        st.success("✅ production_data.xlsx загружен")
        return df_raw
    except Exception as e:
        st.error(f"Ошибка загрузки production_data.xlsx: {e}")
        return pd.DataFrame()

@st.cache_data
def load_logistics_update_data():
    try:
        df = pd.read_excel('BI logisticks.xlsx', header=0)
        
        # Оставляем только строки, где все столбцы I-Q (индексы 8-16) заполнены
        mask = pd.Series([True] * len(df))
        for i in range(8, 17):
            if i < len(df.columns):
                mask = mask & df.iloc[:, i].notna()
        
        df_filtered = df[mask].copy()
        st.success(f"✅ BI logisticks.xlsx загружен, строк: {len(df_filtered)}")
        return df_filtered
    except Exception as e:
        st.error(f"Ошибка загрузки BI logisticks.xlsx: {e}")
        return pd.DataFrame()

# ==========================================
# 2. НАЗВАНИЯ МЕСЯЦЕВ
# ==========================================
month_names = {
    1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
    5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
    9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
}

# ==========================================
# 3. ЗАГРУЗКА
# ==========================================
st.set_page_config(page_title="BI Портал", layout="wide")

st.title("🔧 ДИАГНОСТИКА ЗАГРУЗКИ ФАЙЛОВ")

# Загружаем данные
sales_df = load_sales_data()
logistics_df = load_logistics_data()
production_df = load_production_data()
logistics_update_df = load_logistics_update_data()

# Показываем информацию о загруженных данных
st.divider()
st.subheader("📊 СТАТУС ЗАГРУЗКИ")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("sales_data.xlsx", f"{len(sales_df)} строк" if not sales_df.empty else "❌")
with col2:
    st.metric("logistics_data.xlsx", f"{len(logistics_df)} строк" if not logistics_df.empty else "❌")
with col3:
    st.metric("production_data.xlsx", f"{len(production_df)} строк" if not production_df.empty else "❌")
with col4:
    st.metric("BI logisticks.xlsx", f"{len(logistics_update_df)} строк" if not logistics_update_df.empty else "❌")

# Если все файлы загружены, показываем навигацию
if not sales_df.empty and not logistics_df.empty and not production_df.empty:
    st.divider()
    st.success("✅ Все файлы успешно загружены!")
    
    st.sidebar.title("📊 Навигация")
    page = st.sidebar.radio(
        "Выберите раздел",
        ["📈 Продажи", "🚚 Логистика", "📊 Анализ себестоимости", "🏭 Формирование себестоимости ПФ", "🚚 Логистика Update"]
    )
    
    if page == "📈 Продажи":
        st.write("Страница продаж (временно упрощена)")
        st.dataframe(sales_df.head(10))
    elif page == "🚚 Логистика":
        st.write("Страница логистики (временно упрощена)")
        st.dataframe(logistics_df.head(10))
    elif page == "📊 Анализ себестоимости":
        st.write("Страница анализа себестоимости (временно упрощена)")
        st.dataframe(sales_df.head(10))
    elif page == "🏭 Формирование себестоимости ПФ":
        st.write("Страница формирования себестоимости ПФ (временно упрощена)")
        if not production_df.empty:
            st.dataframe(production_df.head(10))
    elif page == "🚚 Логистика Update":
        st.write("Страница логистики Update (временно упрощена)")
        if not logistics_update_df.empty:
            st.dataframe(logistics_update_df.head(10))
else:
    st.warning("⚠️ Не все файлы загружены. Проверьте наличие файлов в папке.")
    import os
    st.write("Файлы в текущей папке:")
    for f in os.listdir('.'):
        if f.endswith('.xlsx'):
            st.write(f"  - {f}")
