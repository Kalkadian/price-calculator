import streamlit as st
import pandas as pd

def clean_numeric(value):
    """Преобразует строку вида '1 260,00' или '15 шт' в число 1260.0 или 15.0"""
    if pd.isna(value):
        return 0.0
    
    text = str(value)
    # Убираем пробелы и неразрывные пробелы
    text = text.replace(' ', '').replace('\xa0', '').replace('\u2009', '')
    
    # Убираем распространенные единицы измерения и валюты
    for symbol in ['шт', 'ед', 'упак', 'руб', '₽', 'р.', '$', '€', 'руб.', 'м', 'см']:
        text = text.replace(symbol, '')
    
    # Заменяем запятую на точку
    text = text.replace(',', '.')
    
    # Оставляем только цифры и точку
    text = ''.join(c for c in text if c.isdigit() or c == '.')
    
    try:
        return float(text)
    except ValueError:
        return 0.0

# Настройка страницы
st.set_page_config(page_title="Калькулятор цен", layout="wide")

st.title("🌿 Калькулятор стоимости растений")

# --- Блок загрузки файла ---
st.sidebar.header("📁 Загрузка файла")
uploaded_file = st.sidebar.file_uploader(
    "Загрузите Excel или CSV файл с позициями", 
    type=['xlsx', 'xls', 'csv']
)

import pandas as pd

def find_table_in_excel(file_path):
    """
    Автоматически находит таблицу в Excel-файле (первый лист), 
    корректно обрабатывая объединенные ячейки.
    """
    # Читаем ТОЛЬКО ПЕРВЫЙ ЛИСТ (индекс 0) без заголовков
    raw_df = pd.read_excel(file_path, sheet_name=0, header=None)
    
    if raw_df.empty:
        return None, None, None
    
    # Функция проверки на пустоту (игнорируем пробелы и 'nan' как текст)
    def is_not_empty(cell):
        if pd.isna(cell):
            return False
        cell_str = str(cell).strip().lower()
        return cell_str != '' and cell_str != 'nan'
    
    # Создаем маску заполненных ячеек
    mask = raw_df.applymap(is_not_empty)
    
    # Ищем строку с максимальным числом заполненных ячеек (заголовок)
    row_fill_counts = mask.sum(axis=1)
    if row_fill_counts.max() == 0:
        return None, None, None
        
    header_row_idx = row_fill_counts.idxmax()
    
    # Находим последнюю строку с данными ниже заголовка
    data_below = mask.iloc[header_row_idx + 1:]
    if data_below.empty:
        return None, None, None
        
    rows_with_data = data_below.any(axis=1)
    if not rows_with_data.any():
        return None, None, None
        
    last_data_row = rows_with_data[rows_with_data].index[-1]
    
    # Находим границы столбцов с данными
    data_region = mask.iloc[header_row_idx:last_data_row + 1]
    cols_with_data = data_region.any(axis=0)
    data_cols = cols_with_data[cols_with_data].index.tolist()
    
    if not data_cols:
        return None, None, None
    
    first_col = data_cols[0]
    last_col = data_cols[-1]
    
    # Обрезаем DataFrame до найденной области
    trimmed_df = raw_df.iloc[header_row_idx:last_data_row + 1, first_col:last_col + 1].copy()
    
    # =========================================================
    # 🛠️ ГЛАВНОЕ ИСПРАВЛЕНИЕ ДЛЯ ОБЪЕДИНЕННЫХ ЯЧЕЕК И JSON ОШИБКИ
    # =========================================================
    
    # 1. Берем первую строку как заголовки
    header = trimmed_df.iloc[0]
    
    # 2. МАГИЯ: заполняем пропуски (NaN) значением СЛЕВА. 
    # Если было ["Название", NaN], станет ["Название", "Название"]
    header = header.ffill()
    
    # 3. Заменяем оставшиеся NaN (если самый первый столбец был пуст) и пустые строки
    header = header.fillna('Столбец').astype(str).str.strip()
    header = header.replace('', 'Столбец').replace('nan', 'Столбец')
    
    # 4. Делаем имена столбцов уникальными (на случай дубликатов после ffill)
    clean_cols = []
    seen = {}
    for col in header:
        if col in seen:
            seen[col] += 1
            clean_cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            clean_cols.append(col)
            
    trimmed_df.columns = clean_cols
    
    # 5. Убираем строку, которую использовали как заголовок
    trimmed_df = trimmed_df.iloc[1:].reset_index(drop=True)
    
    # 6. УДАЛЯЕМ полностью пустые столбцы (это уберет мусор из выпадающих списков!)
    trimmed_df = trimmed_df.dropna(axis=1, how='all')
    
    # 7. ФИНАЛЬНАЯ ЗАЩИТА ОТ JSON ОШИБКИ: гарантируем, что в именах столбцов нет объектов NaN
    trimmed_df.columns = [str(c).strip() for c in trimmed_df.columns]
    
    # Получаем имя листа
    xl_file = pd.ExcelFile(file_path)
    sheet_name = xl_file.sheet_names[0]
    
    return trimmed_df, sheet_name, header_row_idx

if uploaded_file is not None:
   if uploaded_file is not None:
    try:
        # 🆕 Автоматически ищем таблицу
        with st.spinner("🔍 Ищу таблицу в файле..."):
            df, sheet_name, score = find_table_in_excel(uploaded_file)
        
        if df is None or df.empty:
            st.error("❌ Не удалось найти таблицу в файле. Проверьте формат.")
            st.stop()
        
        # Показываем, что нашли
        st.sidebar.success(f"✅ Таблица найдена!")
        st.sidebar.info(f"📄 Лист: **{sheet_name}**\n📊 Строк: **{len(df)}**\n📋 Столбцов: **{len(df.columns)}**")
        
        # Превью найденной таблицы
        with st.sidebar.expander("👀 Превью найденной таблицы"):
            st.dataframe(df.head(10))
        
        # 🆕 Возможность ручной корректировки (если автопоиск ошибся)
       # 🆕 Возможность ручной корректировки (если автопоиск ошибся)
        with st.sidebar.expander("⚙️ Ручная корректировка (если нужно)"):
            st.caption("Если таблица определилась неверно, можно выбрать строки вручную")
            
            # 🆕 Читаем ТОЛЬКО ПЕРВЫЙ ЛИСТ
            raw_df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            
            use_manual = st.checkbox("Использовать ручной режим")
            
            if use_manual:
                header_row = st.number_input("Строка заголовка (с 0)", 
                                            min_value=0, max_value=len(raw_df)-1, value=header_row)
                
                # Берем заголовки из указанной строки
                df = raw_df.iloc[header_row:].copy()
                df.columns = df.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)
                
                # Убираем полностью пустые столбцы
                df = df.dropna(axis=1, how='all')
                
                st.info(f"📋 Столбцов после очистки: {len(df.columns)}")    

        
        # Выбор колонок
              # Выбор колонок
        st.sidebar.subheader("🎯 Настройка колонок")
        name_col = st.sidebar.selectbox("Колонка с названием", df.columns)
        price_col = st.sidebar.selectbox("Колонка с ценой", df.columns)
        
        # 🆕 Добавляем выбор колонки "В наличии"
        stock_col = st.sidebar.selectbox(
            "Колонка 'В наличии' (опционально)", 
            ['Нет'] + list(df.columns)
        )
        
        # Очистка данных
        df[price_col] = df[price_col].apply(clean_numeric)
        
        # 🆕 Если выбрана колонка с остатками, очищаем и её
        if stock_col != 'Нет':
            df[stock_col] = df[stock_col].apply(clean_numeric)
        # --- Основная часть приложения ---
        st.header("🛒 Выберите позиции")
        
        # Выпадающий список с множественным выбором
        all_positions = df[name_col].tolist()
        selected_positions = st.multiselect(
            "Выберите растения (можно несколько):",
            all_positions,
            default=[]
        )
        
        if selected_positions:
            st.subheader("💰 Укажите количество")
            
            total_sum = 0
            results = []
            
            # Создаем таблицу с полями ввода количества
            cols = st.columns([3, 2, 2, 2])
            cols[0].write("**Название**")
            cols[1].write("**Цена за шт.**")
            cols[2].write("**Количество**")
            cols[3].write("**Сумма**")
            
            for position in selected_positions:
                # Получаем строку с данными по этой позиции
                row_data = df[df[name_col] == position].iloc[0]
                price = float(row_data[price_col])
                
                # 🆕 Определяем наличие
                if stock_col != 'Нет' and stock_col in df.columns:
                    stock = int(row_data[stock_col])
                    stock_display = f"📦 В наличии: **{stock}**"
                else:
                    stock = float('inf')  # Бесконечность, если колонка не выбрана
                    stock_display = "📦 В наличии: не указано"
                
                # Рисуем строку таблицы
                cols = st.columns([3, 2, 2, 2])
                cols[0].write(position)
                cols[1].write(f"{price:.2f} ₽")
                
                # Поле ввода количества
                quantity = cols[2].number_input(
                    "qty", 
                    min_value=0, 
                    value=1, 
                    key=f"qty_{position}",
                    label_visibility="collapsed"
                )
                
                # 🆕 ПРОВЕРКА НА ПРЕВЫШЕНИЕ ОСТАТКОВ
                if stock != float('inf') and quantity > stock:
                    cols[2].error(f"⚠️ Превышено на {quantity - stock}!")
                elif stock != float('inf') and quantity == stock:
                    cols[2].warning("⚠️ Под ноль!")
                
                position_sum = price * quantity
                cols[3].write(f"**{position_sum:.2f} ₽**")
                
                total_sum += position_sum
                
                # Добавляем в итоговый список только если количество > 0
                if quantity > 0:
                    results.append({
                        'Название': position,
                        'Цена': price,
                        'В наличии': stock if stock != float('inf') else 'Н/Д',
                        'Заказано': quantity,
                        'Сумма': position_sum
                    })
            
            # Итоговая сумма
                        # Итоговая сумма
            st.markdown("---")
            col1, col2 = st.columns([3, 1])
            col1.subheader("📊 Итого к оплате:")
            col2.subheader(f"💵 {total_sum:,.2f} ₽") # Добавлено форматирование с пробелами
            
            # Таблица с результатами
            if results:
                st.subheader("📋 Итоговая спецификация")
                result_df = pd.DataFrame(results)
                
                # 🆕 Подсветка строк, где заказано больше, чем в наличии
                def highlight_overstock(row):
                    if isinstance(row['В наличии'], (int, float)) and row['Заказано'] > row['В наличии']:
                        return ['background-color: #ffcccc'] * len(row) # Красный фон
                    return [''] * len(row)
                
                st.dataframe(
                    result_df.style.apply(highlight_overstock, axis=1),
                    hide_index=True,
                    use_container_width=True
                )
                
                # Кнопка скачивания
                csv = result_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button(
                    label="📥 Скачать спецификацию (CSV)",
                    data=csv,
                    file_name="спецификация.csv",
                    mime="text/csv"
                )
        else:
            st.info("👆 Выберите хотя бы одну позицию из списка выше")
    
    except Exception as e:
        st.error(f"❌ Ошибка при чтении файла: {e}")
        st.write("Убедитесь, что файл содержит корректные данные")

else:
    st.info("👈 Загрузите файл в боковой панели слева")
    
    # Пример формата файла
    st.subheader("📝 Пример формата файла")
    example_df = pd.DataFrame({
        'Название': ['Драцена маргината 2 ств 17/75', 'Фикус бенджамина 17/75', 'Монстера 50/50/50'],
        'Цена': [1500.00, 2300.00, 3500.00]
    })
    st.dataframe(example_df, hide_index=True)
