import os
import streamlit as st
from Template.base import template_sidebar, confidential
from View.base import process_files
from openai import OpenAI

st.set_page_config(page_title="Question PDF", page_icon="🛠️")

confidential()

template_sidebar()

st.markdown("# Upload form files")
st.sidebar.header("Main")
st.write(
    """Give it pdfs or docx files"""
)

def on_upload(uploaded_files, info_file):
    main_directory = os.path.join(os.path.dirname(__file__), 'tmp')

    # Create the directory if it doesn't exist
    os.makedirs(main_directory, exist_ok=True)

    new_paths = uploaded_files
    new_paths.append(info_file)

    file_paths = []

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        file_path = os.path.join(main_directory, file_name)

        # Handle case where file with the same name already exists
        if os.path.exists(file_path):
            st.warning(f"A file with the name '{file_name}' already exists. Renaming...")
            file_name, file_extension = os.path.splitext(file_name)
            i = 1
            while os.path.exists(os.path.join(main_directory, f"{file_name}_{i}{file_extension}")):
                i += 1
            file_name = f"{file_name}_{i}{file_extension}"

        # Save the file
        with open(os.path.join(main_directory, file_name), "wb") as file:
            file.write(uploaded_file.getvalue())

        file_paths.append(os.path.join(main_directory, file_name))

    qas = process_files(file_paths[:-1], file_paths[-1])
    st.json(qas)

uploaded_main_file = st.file_uploader("Upload your file", type=['pdf', 'docx', 'doc', 'xls', 'xlsx'], accept_multiple_files=False)
uploaded_info_file = st.file_uploader("Upload your file", type=['csv'], accept_multiple_files=False)

if uploaded_main_file and uploaded_info_file:
    uploaded_files = [uploaded_main_file]
    progress_bar = st.sidebar.progress(20)
    on_upload(uploaded_files, uploaded_info_file)
    progress_bar.progress(100)
