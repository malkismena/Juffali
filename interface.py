import streamlit as st
import os
from translator import main
from io import BytesIO
import shutil

original_versions = {
    "تقرير شهر مايو.xlsx": "تقرير شهر مايو - أصلي.xlsx",
    "تصنيف مستفيدي السكن الداخلي في الأنشطة الترفيهية والاجتماعية.xlsx": "تصنيف مستفيدي السكن الداخلي في الأنشطة الترفيهية والاجتماعية - أصلي.xlsx"
}

if "initialized" not in st.session_state:
    for temp_file, original_file in original_versions.items():
        if os.path.exists(original_file):
            if os.path.exists(temp_file):
                os.remove(temp_file)
            shutil.copy(original_file, temp_file)
    st.session_state.initialized = True


available_files = [
    ("تقرير شهر مايو.xlsx", "C", 7),
    ("تصنيف مستفيدي السكن الداخلي في الأنشطة الترفيهية والاجتماعية.xlsx", "F", 11)
]

def load_file_bytes(file_path):
    with open(file_path, "rb") as f:
        return f.read()

st.set_page_config(page_title="Munazzim Chatbot", layout="wide")

st.markdown("<h1 style='text-align: center; color: #2c3e50;'>Munazzim Chatbot</h1>", unsafe_allow_html=True)

if "conversation_history" not in st.session_state:
    st.session_state.conversation_history = []


for message in st.session_state.conversation_history:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

user_input = st.chat_input("💬 أدخل الجملة (بالعربية، منظمة أو غير منظمة):")

if user_input:
    with st.chat_message("user"):
        st.markdown(user_input)

    st.session_state.conversation_history.append({
        "role": "user",
        "content": user_input,
    })

    ai_response = main(user_input)

    with st.chat_message("assistant"):
        st.markdown(ai_response)

    st.session_state.conversation_history.append({
        "role": "assistant",
        "content": ai_response,
    })
    
    with st.sidebar:
        for file_name, _, _ in available_files:
            if os.path.exists(file_name):
                file_bytes = load_file_bytes(file_name)
                
with st.sidebar:
    st.markdown("### 📝 مساعد إدخال البيانات")
    st.markdown("الرجاء إدخال المعلومات المطلوبة بدقة.\nسيتم حفظها تلقائيًا في ملف Excel.")
    st.markdown("---")

    st.markdown("### 📂 تحميل الملفات المحدثة:")

    for file_name, _, _ in available_files:
        if os.path.exists(file_name):
            file_bytes = load_file_bytes(file_name)
            st.download_button(
                label=f"📥 تحميل {file_name}",
                data=file_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"sidebar-{file_name}"
            )
        else:
            st.warning(f"⚠️ الملف غير موجود: {file_name}")
