import os
import tempfile
from pathlib import Path

import streamlit as st

from libs.spec_pdf_to_xlsx import DEFAULT_TABLE_SETTINGS, pdf_spec_to_row_list, row_list_to_xlsx_bytes

# Настройки страницы
st.set_page_config(
    page_title="Спецификация: PDF → XLSX",
    page_icon="📄",
    layout="centered",
)



st.title("📄 Спецификация: PDF → XLSX")
#st.markdown("Загрузите PDF-файл спецификации для её перевода в формат XLSX.")


# Инициализация настроек pdfplumber (при первом запуске — значения по умолчанию)
if "table_settings" not in st.session_state:
    st.session_state["table_settings"] = dict(DEFAULT_TABLE_SETTINGS)


@st.dialog("Настройки pdfplumber")
def pdfplumber_settings_dialog():
    """Модальное окно редактирования параметров извлечения таблиц pdfplumber."""
    current = st.session_state["table_settings"]

    values = {}
    for key in DEFAULT_TABLE_SETTINGS:
        values[key] = st.number_input(
            f"`{key}`",
            min_value=0,
            value=int(current[key]),
            step=1,
            key=f"pdfplumber_{key}",
        )

    changed = any(values[key] != current[key] for key in DEFAULT_TABLE_SETTINGS)
    default = all(values[key] == DEFAULT_TABLE_SETTINGS[key] for key in DEFAULT_TABLE_SETTINGS)

    # 2. Три колонки для распределения пространства
    col_left, col_center, col_right = st.columns([1, 1, 1])

    with col_left:
        apply_clicked = st.button("Применить", type="primary", disabled=not changed, key="btn_apply")

    with col_center:
        defaults_clicked = st.button("Значения по умолчанию", disabled=default, key="btn_defaults")

    with col_right:
        cancel_clicked = st.button("Отмена", key="btn_cancel")

    # 3. Логика
    if defaults_clicked:
        st.session_state["table_settings"] = dict(DEFAULT_TABLE_SETTINGS)
        for key in DEFAULT_TABLE_SETTINGS:
            st.session_state.pop(f"pdfplumber_{key}", None)
        st.rerun()

    if apply_clicked:
        for key in DEFAULT_TABLE_SETTINGS:
            st.session_state["table_settings"][key] = int(values[key])
        st.rerun()

    if cancel_clicked:
        st.rerun()





# 1. Загрузка файла
pdf_file = st.file_uploader(
    "Выберите PDF файл, нажав Upload, или перетащите его сюда мышкой из Проводника", type=["pdf"]
)

# При загрузке нового файла сбрасываем настройки pdfplumber в значения по умолчанию
uploaded_file_id = (pdf_file.name, pdf_file.size) if pdf_file is not None else None
if st.session_state.get("uploaded_file_id") != uploaded_file_id:
    st.session_state["uploaded_file_id"] = uploaded_file_id
    st.session_state["table_settings"] = dict(DEFAULT_TABLE_SETTINGS)
    # Сбрасываем ключи виджетов диалога настроек, чтобы он открывался с дефолтами
    for key in DEFAULT_TABLE_SETTINGS:
        st.session_state.pop(f"pdfplumber_{key}", None)

if pdf_file is not None:
    #st.info(f"Загружен файл: **{pdf_file.name}** ({pdf_file.size / 1024:.1f} КБ)")

    # 2. Кнопка настроек pdfplumber
    if st.button("⚙️ Настройки pdfplumber"):
        # Сбрасываем ключи виджетов диалога, чтобы поля открывались
        # со значениями текущих настроек
        for key in DEFAULT_TABLE_SETTINGS:
            st.session_state.pop(f"pdfplumber_{key}", None)
        pdfplumber_settings_dialog()

    # 3. Кнопка запуска обработки
    if st.button("🚀 Извлечь данные из PDF и сформировать XLSX", type="primary"):
        with st.spinner("Идёт анализ PDF ... Пожалуйста, подождите."):
            pdf_path = None
            try:
                # Сохраняем PDF во временный файл (нужен pdfplumber)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(pdf_file.getvalue())
                    pdf_path = tmp_pdf.name

                spec, template_cols_count,restored_rows = pdf_spec_to_row_list(pdf_path, st.session_state["table_settings"])
                if template_cols_count == 9:
                    # Путь к шаблону спецификации
                    TEMPLATE_PATH = Path("templates") / "Шаблон_спецификации_РД.xlsx"
                elif template_cols_count == 10:
                    TEMPLATE_PATH = Path("templates") / "Шаблон_спецификации_РД_KKS.xlsx"
                else:
                    raise ValueError("В спецификации должно быть 9 или 10 столбцов.")
                xlsx_bytes = row_list_to_xlsx_bytes(spec, template_cols_count, restored_rows, Path(pdf_file.name).stem, TEMPLATE_PATH)
                xlsx_name = Path(pdf_file.name).with_suffix(".xlsx").name

                st.success("✅ Файл успешно обработан!")

                # 3. Кнопка скачивания
                st.download_button(
                    label="📥 Скачать XLSX файл",
                    data=xlsx_bytes,
                    file_name=xlsx_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

            except Exception as e:
                st.error(f"❌ Ошибка: {e}")
                #st.exception(e)

            finally:
                # 4. Очистка временного PDF-файла
                if pdf_path:
                    try:
                        os.remove(pdf_path)
                    except OSError:
                        pass

else:
    #st.warning("Загрузите PDF файл спецификации для начала работы.")
    pass

st.markdown(
    """
    <hr>
    <p style="text-align: left; color: gray;">
    <small>
    2026, С.В. Медведев, engpython@yandex.ru
    </small>
    </p>
    """,
    unsafe_allow_html=True,
)
