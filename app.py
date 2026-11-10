import sqlite3
from urllib.parse import urlparse
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.chart import PieChart, Reference

DB_NAME = "links_database.db"

# --- Инициализация БД ---
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS links (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            url TEXT UNIQUE,
            root_domain TEXT,
            title TEXT,
            category TEXT,
            subcategory TEXT,
            tags TEXT,
            color TEXT,
            date_added TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_root_domain(url):
    try:
        parsed = urlparse(url)
        netloc = parsed.netloc or parsed.path.split('/')[0]
        parts = netloc.split('.')
        if len(parts) >= 2:
            return f"{parts[-2]}.{parts[-1]}"
        return netloc
    except:
        return "unknown"

st.set_page_config(page_title="Менеджер Ссылок", layout="wide")
st.title("📂 Умный менеджер ссылок и баз данных")

menu = st.sidebar.selectbox("Меню", ["Добавить пачку ссылок", "Просмотр базы и Экспорт"])

if menu == "Добавить пачку ссылок":
    st.subheader("Шаг 1: Вставьте новые ссылки (до 100 шт.)")
    raw_text = st.text_area("Введите ссылки (каждая с новой строки):", height=150)
    
    if st.button("Проверить и обработать"):
        if not raw_text.strip():
            st.warning("Пожалуйста, введите хотя бы одну ссылку.")
        else:
            urls = [line.strip() for line in raw_text.split('\n') if line.strip()]
            
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            
            new_items = []
            duplicates = []
            
            for url in urls:
                # Проверка на дубликат
                cursor.execute("SELECT category, title FROM links WHERE url = ?", (url,))
                existing = cursor.fetchone()
                if existing:
                    duplicates.append({"url": url, "existing_category": existing[0], "existing_title": existing[1]})
                else:
                    root_dom = get_root_domain(url)
                    # Базовая авто-категоризация по домену
                    new_items.append({
                        "url": url,
                        "root_domain": root_dom,
                        "title": root_dom.capitalize(),
                        "category": "Общее",
                        "subcategory": "Разное",
                        "tags": "веб",
                        "color": "#4B9CD3"
                    })
            conn.close()
            
            st.session_state['new_items'] = new_items
            st.session_state['duplicates'] = duplicates
            st.success(f"Обработано! Новых уникальных: {len(new_items)}, Дубликатов найдено: {len(duplicates)}")

    if 'new_items' in st.session_state and st.session_state['new_items']:
        st.subheader("Шаг 2: Редактирование и подтверждение новых ссылок")
        st.info("Вы можете изменить название, категорию, субкатегорию и цвет перед сохранением.")
        
        df_edit = pd.DataFrame(st.session_state['new_items'])
        edited_df = st.data_editor(df_edit, num_rows="dynamic", key="link_editor")
        
        if st.button("Подтвердить и сохранить в базу"):
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            
            for index, row in edited_df.iterrows():
                try:
                    cursor.execute('''
                        INSERT OR IGNORE INTO links (url, root_domain, title, category, subcategory, tags, color)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (row['url'], row['root_domain'], row['title'], row['category'], row['subcategory'], row['tags'], row['color']))
                except Exception as e:
                    st.error(e)
            conn.commit()
            conn.close()
            st.success("Ссылки успешно сохранены! База и файлы обновлены.")
            del st.session_state['new_items']

elif menu == "Просмотр базы и Экспорт":
    st.subheader("Управление сохраненными данными")
    conn = sqlite3.connect(DB_NAME)
    df_all = pd.read_sql("SELECT * FROM links", conn)
    conn.close()
    
    st.write(f"Всего ссылок в базе: **{len(df_all)}**")
    
    if not df_all.empty:
        st.dataframe(df_all)
        
        # Генерация Excel с листами
        if st.button("Сгенерировать и скачать файлы (Excel, Markdown, HTML)"):
            excel_path = "links_dashboard.xlsx"
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df_all.to_excel(writer, sheet_name='Обзор', index=False)
                
                for cat in df_all['category'].unique():
                    df_cat = df_all[df_all['category'] == cat]
                    safe_cat_name = str(cat)[:31] # Ограничение Excel на длину имени листа
                    df_cat.to_excel(writer, sheet_name=safe_cat_name, index=False)
            
            with open(excel_path, "rb") as f:
                st.download_button("📥 Скачать Excel-дашборд (.xlsx)", f, file_name="links_dashboard.xlsx")
            
            # Генерация Markdown
            md_content = "# Моя база ссылок\n\n"
            for cat, group in df_all.groupby('category'):
                md_content += f"## Категория: {cat}\n"
                for _, row in group.iterrows():
                    md_content += f"* [{row['title']}]({row['url']}) — *{row['subcategory']}* [Теги: {row['tags']}]\n"
                md_content += "\n"
            
            st.download_button("📥 Скачать Markdown (.md)", md_content, file_name="links_archive.md", mime="text/markdown")
            
            # Генерация HTML закладок для браузера (Kiwi / Chrome)
            html_bookmarks = """<!DOCTYPE NETSCAPE-Bookmark-file-1>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>
<DL><p>
"""
            for cat, group in df_all.groupby('category'):
                html_bookmarks += f"    <DT><H3>{cat}</H3>\n    <DL><p>\n"
                for _, row in group.iterrows():
                    html_bookmarks += f'        <DT><A HREF="{row["url"]}">{row["title"]}</A>\n'
                html_bookmarks += "    </DL><p>\n"
            html_bookmarks += "</DL><p>"
            
            st.download_button("📥 Скачать закладки для браузера (bookmarks.html)", html_bookmarks, file_name="bookmarks.html", mime="text/html")
