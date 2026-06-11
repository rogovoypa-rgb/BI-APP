import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

st.set_page_config(page_title="Планирование производства", layout="wide")

st.title("📋 Планирование и производство")
st.markdown("### Управление производственным планом обжарок")

# ==========================================
# ФУНКЦИИ ДЛЯ РАБОТЫ С ДАННЫМИ
# ==========================================
def load_planning_data():
    """Загрузка данных планирования из Excel"""
    try:
        if os.path.exists('planning_data.xlsx'):
            df = pd.read_excel('planning_data.xlsx')
            return df.to_dict('records')
        else:
            return []
    except Exception as e:
        st.error(f"Ошибка загрузки: {e}")
        return []

def save_planning_data(records):
    """Сохранение данных планирования в Excel"""
    try:
        df = pd.DataFrame(records)
        df.to_excel('planning_data.xlsx', index=False)
        return True
    except Exception as e:
        st.error(f"Ошибка сохранения: {e}")
        return False

# ==========================================
# ИНИЦИАЛИЗАЦИЯ СОСТОЯНИЯ
# ==========================================
if 'planning_data' not in st.session_state:
    st.session_state.planning_data = load_planning_data()

if 'roasters_list' not in st.session_state:
    # Загружаем список ростеров из сохраненных данных или используем стандартный
    roasters = set()
    for record in st.session_state.planning_data:
        roasters.add(record.get('Ростер', ''))
    if roasters:
        st.session_state.roasters_list = sorted(list(roasters))
    else:
        st.session_state.roasters_list = ["Ростер №1 (15кг)", "Ростер №2 (30кг)", "Ростер №3 (60кг)"]

# ==========================================
# ФУНКЦИЯ ДЛЯ ПОЛУЧЕНИЯ ПЛАНА ПО КЛЮЧУ
# ==========================================
def get_plan_dict(roaster, start_date, days):
    """Получает план в виде словаря для отображения"""
    plan_dict = {}
    dates = [start_date + timedelta(days=i) for i in range(days)]
    
    for date in dates:
        date_str = date.strftime('%Y-%m-%d')
        plan_dict[date_str] = {}
        for slot in range(1, 31):
            # Ищем сохраненные данные
            found = False
            for record in st.session_state.planning_data:
                if (record.get('Ростер') == roaster and 
                    record.get('Дата') == date_str and 
                    record.get('Слот') == slot):
                    plan_dict[date_str][slot] = {
                        'запланировано': record.get('Запланировано', False),
                        'кофе': record.get('Кофе', ''),
                        'вес': record.get('Вес_кг', 0),
                        'примечание': record.get('Примечание', ''),
                        'выполнено': record.get('Выполнено', False)
                    }
                    found = True
                    break
            
            if not found:
                plan_dict[date_str][slot] = {
                    'запланировано': False,
                    'кофе': '',
                    'вес': 0,
                    'примечание': '',
                    'выполнено': False
                }
    
    return plan_dict

def save_slot(roaster, date_str, slot_num, slot_data):
    """Сохраняет данные одного слота"""
    # Удаляем старую запись если есть
    st.session_state.planning_data = [
        r for r in st.session_state.planning_data 
        if not (r.get('Ростер') == roaster and r.get('Дата') == date_str and r.get('Слот') == slot_num)
    ]
    
    # Добавляем новую
    st.session_state.planning_data.append({
        'Ростер': roaster,
        'Дата': date_str,
        'Слот': slot_num,
        'Время_начала': f"{(slot_num-1)*15} мин",
        'Запланировано': slot_data['запланировано'],
        'Кофе': slot_data['кофе'],
        'Вес_кг': slot_data['вес'],
        'Примечание': slot_data['примечание'],
        'Выполнено': slot_data.get('выполнено', False)
    })
    
    # Сохраняем в файл
    save_planning_data(st.session_state.planning_data)

# ==========================================
# ВКЛАДКИ
# ==========================================
tab1, tab2, tab3 = st.tabs(["📅 План обжарок", "📊 Производственная статистика", "⚙️ Настройки"])

# ==========================================
# ВКЛАДКА 1: ПЛАН ОБЖАРОК
# ==========================================
with tab1:
    st.subheader("📅 План обжарок по ростерам")
    
    col1, col2 = st.columns(2)
    with col1:
        selected_roaster = st.selectbox("🔥 Выберите ростер", st.session_state.roasters_list)
    with col2:
        start_date = st.date_input("📅 Начало периода", value=datetime.now())
    
    days_to_show = st.slider("📆 Количество дней", min_value=1, max_value=14, value=7)
    
    dates = [start_date + timedelta(days=i) for i in range(days_to_show)]
    current_plan = get_plan_dict(selected_roaster, start_date, days_to_show)
    
    day_names_ru = {
        'Monday': 'Понедельник', 'Tuesday': 'Вторник', 'Wednesday': 'Среда',
        'Thursday': 'Четверг', 'Friday': 'Пятница', 'Saturday': 'Суббота', 'Sunday': 'Воскресенье'
    }
    
    for date in dates:
        date_str = date.strftime('%Y-%m-%d')
        day_name_ru = day_names_ru.get(date.strftime('%A'), date.strftime('%A'))
        
        planned_count = sum(1 for s in current_plan[date_str].values() if s['запланировано'])
        completed_count = sum(1 for s in current_plan[date_str].values() if s.get('выполнено', False))
        
        with st.expander(f"📅 {date_str} - {day_name_ru}", expanded=(date == dates[0])):
            st.caption(f"📊 Слотов: 30 | Запланировано: {planned_count} | Выполнено: {completed_count}")
            
            slots_per_row = 6
            slot_num = 1
            
            while slot_num <= 30:
                cols = st.columns(slots_per_row)
                for i, col in enumerate(cols):
                    if slot_num <= 30:
                        slot_data = current_plan[date_str][slot_num]
                        slot_status = "✅" if slot_data.get('выполнено', False) else ("🟢" if slot_data['запланировано'] else "⚪")
                        slot_label = f"{slot_status} Слот {slot_num}"
                        
                        with col:
                            with st.popover(slot_label, use_container_width=True):
                                st.write(f"**Слот {slot_num}**")
                                st.write(f"⏰ Время: {(slot_num-1)*15} мин")
                                
                                new_status = st.checkbox("📌 Запланировано", value=slot_data['запланировано'], 
                                                        key=f"{date_str}_{selected_roaster}_slot_{slot_num}_status")
                                coffee_name = st.text_input("☕ Название кофе", value=slot_data['кофе'], 
                                                          key=f"{date_str}_{selected_roaster}_slot_{slot_num}_coffee")
                                weight = st.number_input("⚖️ Вес (кг)", min_value=0, max_value=100, value=slot_data['вес'], 
                                                        key=f"{date_str}_{selected_roaster}_slot_{slot_num}_weight")
                                completed = st.checkbox("✅ Выполнено", value=slot_data.get('выполнено', False),
                                                      key=f"{date_str}_{selected_roaster}_slot_{slot_num}_completed")
                                note = st.text_area("📝 Примечание", value=slot_data['примечание'], 
                                                   key=f"{date_str}_{selected_roaster}_slot_{slot_num}_note", height=68)
                                
                                if st.button("💾 Сохранить", key=f"{date_str}_{selected_roaster}_slot_{slot_num}_save"):
                                    save_slot(selected_roaster, date_str, slot_num, {
                                        'запланировано': new_status,
                                        'кофе': coffee_name,
                                        'вес': weight,
                                        'примечание': note,
                                        'выполнено': completed
                                    })
                                    st.rerun()
                        
                        slot_num += 1
                st.markdown("---")
    
    # Экспорт плана
    st.divider()
    if st.button("📥 Экспортировать текущий план в Excel", use_container_width=True):
        rows = []
        for date_str, slots in current_plan.items():
            for slot_num, slot_data in slots.items():
                rows.append({
                    'Дата': date_str,
                    'Ростер': selected_roaster,
                    'Слот': slot_num,
                    'Время_начала': f"{(slot_num-1)*15} мин",
                    'Запланировано': slot_data['запланировано'],
                    'Кофе': slot_data['кофе'],
                    'Вес_кг': slot_data['вес'],
                    'Примечание': slot_data['примечание'],
                    'Выполнено': slot_data.get('выполнено', False)
                })
        df_export = pd.DataFrame(rows)
        csv = df_export.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
        st.download_button("📥 Скачать план (CSV)", csv, f"plan_{selected_roaster}_{start_date}.csv", "text/csv")

# ==========================================
# ВКЛАДКА 2: ПРОИЗВОДСТВЕННАЯ СТАТИСТИКА
# ==========================================
with tab2:
    st.subheader("📊 Производственная статистика")
    
    total_planned = 0
    total_completed = 0
    roaster_stats = {}
    
    for record in st.session_state.planning_data:
        roaster = record.get('Ростер', 'Неизвестно')
        if roaster not in roaster_stats:
            roaster_stats[roaster] = {'planned': 0, 'completed': 0}
        
        if record.get('Запланировано', False):
            roaster_stats[roaster]['planned'] += 1
            total_planned += 1
            if record.get('Выполнено', False):
                roaster_stats[roaster]['completed'] += 1
                total_completed += 1
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Всего запланировано", format_number(total_planned))
    with col2:
        st.metric("✅ Выполнено", format_number(total_completed))
    with col3:
        percent = (total_completed / total_planned * 100) if total_planned > 0 else 0
        st.metric("📊 Процент выполнения", f"{percent:.1f}%")
    with col4:
        roaster_count = len(roaster_stats)
        st.metric("🔥 Активных ростеров", roaster_count)
    
    if roaster_stats:
        st.divider()
        st.subheader("📊 Статистика по ростерам")
        
        stats_data = []
        for roaster, stats in roaster_stats.items():
            percent = (stats['completed'] / stats['planned'] * 100) if stats['planned'] > 0 else 0
            stats_data.append({
                'Ростер': roaster,
                'Запланировано': stats['planned'],
                'Выполнено': stats['completed'],
                'Процент': f"{percent:.1f}%"
            })
        st.dataframe(pd.DataFrame(stats_data), use_container_width=True, hide_index=True)

# ==========================================
# ВКЛАДКА 3: НАСТРОЙКИ
# ==========================================
with tab3:
    st.subheader("⚙️ Настройки")
    
    col1, col2 = st.columns(2)
    with col1:
        new_roaster = st.text_input("➕ Добавить новый ростер")
        if st.button("Добавить ростер", use_container_width=True):
            if new_roaster and new_roaster not in st.session_state.roasters_list:
                st.session_state.roasters_list.append(new_roaster)
                st.success(f"Ростер '{new_roaster}' добавлен!")
                st.rerun()
        
        st.write("**Текущие ростера:**")
        for r in st.session_state.roasters_list:
            st.write(f"- {r}")
    
    with col2:
        if st.button("🗑️ Очистить все данные", use_container_width=True):
            st.session_state.planning_data = []
            save_planning_data([])
            st.success("Все данные очищены!")
            st.rerun()
    
    st.divider()
    st.subheader("📊 Импорт/Экспорт всех данных")
    
    if st.button("📥 Экспортировать все данные в Excel"):
        if st.session_state.planning_data:
            df_all = pd.DataFrame(st.session_state.planning_data)
            csv = df_all.to_csv(index=False, sep=';', decimal=',').encode('utf-8-sig')
            st.download_button("📥 Скачать все данные (CSV)", csv, "all_planning_data.csv", "text/csv")
        else:
            st.warning("Нет данных для экспорта")
    
    uploaded_file = st.file_uploader("📂 Импортировать данные из Excel/CSV", type=["xlsx", "csv"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.xlsx'):
                df_upload = pd.read_excel(uploaded_file)
            else:
                df_upload = pd.read_csv(uploaded_file, sep=';')
            
            records = df_upload.to_dict('records')
            st.session_state.planning_data = records
            save_planning_data(records)
            st.success(f"Импортировано {len(records)} записей!")
            st.rerun()
        except Exception as e:
            st.error(f"Ошибка импорта: {e}")
