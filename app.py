import os
import tempfile
from pathlib import Path

import streamlit as st

from libs.spec_pdf_to_xlsx import pdf_spec_to_row_list, row_list_to_xlsx_bytes

# Настройки страницы
st.set_page_config(
    page_title="Спецификация: PDF → XLSX",
    page_icon="📄",
    layout="centered",
)



st.title("📄 Спецификация: PDF → XLSX")
#st.markdown("Загрузите PDF-файл спецификации для её перевода в формат XLSX.")


# 1. Загрузка файла
pdf_file = st.file_uploader(
    "Выберите PDF файл, нажав Upload, или перетащите его сюда мышкой из Проводника", type=["pdf"]
)

if pdf_file is not None:
    #st.info(f"Загружен файл: **{pdf_file.name}** ({pdf_file.size / 1024:.1f} КБ)")

    # 2. Кнопка запуска обработки
    if st.button("🚀 Извлечь данные из PDF и сформировать XLSX", type="primary"):
        with st.spinner("Идёт анализ PDF ... Пожалуйста, подождите."):
            pdf_path = None
            try:
                # Сохраняем PDF во временный файл (нужен pdfplumber)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(pdf_file.getvalue())
                    pdf_path = tmp_pdf.name

                spec, template_cols_count,restored_rows = pdf_spec_to_row_list(pdf_path)
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
                st.error(f"❌ Произошла ошибка при обработке файла: {e}")
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
