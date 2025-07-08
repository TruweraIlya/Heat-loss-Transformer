import streamlit as st
from transform import transform_excel

st.set_page_config(page_title="Transform App", page_icon="⚙️")

st.title("⚙️ Автоматическая обработка Excel")

uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx"])

language = st.radio("Выберите язык выходного файла:", ("Русский", "Английский"))

if uploaded_file:
    st.success("Файл загружен успешно.")

    if st.button("Выполнить преобразование"):
        lang_code = "ru" if language == "Русский" else "en"
        result_path = transform_excel(uploaded_file, language=lang_code)

        if result_path.endswith(".xlsx"):
            with open(result_path, "rb") as f:
                st.download_button(
                    label="📥 Скачать готовый файл",
                    data=f,
                    file_name=result_path.split("\\")[-1],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.success("Файл успешно обработан!")
        else:
            st.error(result_path)
else:
    st.info("Пожалуйста, загрузите файл для обработки.")
