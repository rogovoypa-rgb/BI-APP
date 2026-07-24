import streamlit as st
import bcrypt
import os
from dotenv import load_dotenv

load_dotenv()  # загружаем переменные из .env (не коммитим!)

# Хранилище пользователей: логин -> {password_hash, role}
# Пароли хешированы bcrypt. Для генерации хеша используй: bcrypt.hashpw("plain".encode(), bcrypt.gensalt())
USERS = {
    "admin": {
        "hash": os.getenv("ADMIN_HASH"),  # в .env: ADMIN_HASH=$2b$12$...
        "role": "admin"
    },
    "partner1": {
        "hash": os.getenv("PARTNER1_HASH"),
        "role": "partner"
    },
    "partner2": {
        "hash": os.getenv("PARTNER2_HASH"),
        "role": "partner"
    },
    "manager1": {
        "hash": os.getenv("MANAGER1_HASH"),
        "role": "manager"
    },
    # ... добавь всех 15 сотрудников, но лучше потом перенести в отдельный файл
}

def check_password(username, password):
    if username not in USERS:
        return False, None
    stored_hash = USERS[username]["hash"]
    if bcrypt.checkpw(password.encode(), stored_hash.encode()):
        return True, USERS[username]["role"]
    return False, None

def login():
    st.sidebar.markdown("## 🔑 Вход")
    username = st.sidebar.text_input("Логин")
    password = st.sidebar.text_input("Пароль", type="password")
    if st.sidebar.button("Войти"):
        ok, role = check_password(username, password)
        if ok:
            st.session_state["logged_in"] = True
            st.session_state["username"] = username
            st.session_state["role"] = role
            st.rerun()
        else:
            st.sidebar.error("Неверный логин или пароль")

def logout():
    if st.sidebar.button("Выйти"):
        for key in ["logged_in", "username", "role"]:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

def require_auth():
    if "logged_in" not in st.session_state or not st.session_state["logged_in"]:
        login()
        st.stop()  # останавливаем рендеринг основного контента
    else:
        st.sidebar.markdown(f"👤 {st.session_state['username']} ({st.session_state['role']})")
        logout()
