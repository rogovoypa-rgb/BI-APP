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
# 3. БОКОВАЯ ПАНЕЛЬ С ФИЛЬТРАМИ
# ==========================================
st.sidebar.header("🔍 Фильтры")

# Выбор месяца
months = sorted(df['Месяц'].dropna().unique())
selected_month = st.sidebar.selectbox("Выберите месяц", months)

# Выбор контрагентов
all_customers = sorted(df['Контрагент'].dropna().unique())
selected_customers = st.sidebar.multiselect(
    "Выберите контрагентов",
    all_customers,
    default=all_customers[:5] if len(all_customers) > 5 else all_customers
)

# Применяем фильтры
df_filtered = df[(df['Месяц'] == selected_month) & (df['Контрагент'].isin(selected_customers))]

# ==========================================
# 4. ОСНОВНЫЕ МЕТРИКИ (KPI)
# ==========================================
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
# 5. ГРАФИКИ
# ==========================================
st.divider()
st.subheader(f"📊 Аналитика за {selected_month}")

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
margin_by_month = df[df['Контрагент'].isin(selected_customers)].groupby('Месяц')['Рентабельность_%'].mean().reset_index()
if not margin_by_month.empty:
    fig3 = px.line(margin_by_month, x='Месяц', y='Рентабельность_%', 
                   markers=True, line_shape='linear',
                   title=f'Средняя рентабельность выбранных контрагентов',
                   labels={'Рентабельность_%': 'Рентабельность (%)', 'Месяц': ''})
    fig3.update_layout(hovermode='x unified', height=450)
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.info("Нет данных для отображения динамики")

# ==========================================
# 6. ТАБЛИЦА С ДАННЫМИ
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
        file_name=f'sales_data_{selected_month}.csv',
        mime='text/csv',
    )
else:
    st.warning("Нет данных для отображения. Попробуйте изменить фильтры.")

# ==========================================
# 7. ИТОГОВАЯ СТАТИСТИКА
# ==========================================
st.divider()
st.caption(f"📅 Данные за {selected_month} | Всего записей: {len(df_filtered)} | Обновлено: автоматически при запуске")