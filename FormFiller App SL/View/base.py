from View.analyzer import WordAnalyzer, PdfAnalyzer, ExcelAnalyzer, Analyzer
from View.asker import Asker
from openai import OpenAI
import pandas as pd

def is_safe():
    import streamlit as st
    if "authentication_status" in st.session_state:
        if st.session_state["authentication_status"] == True:
            return True
    return False

def process_files(file_path_list, person_info_path = "./result/dummy_data.csv"):
    import os

    df = pd.read_csv(person_info_path)
    data_dict = df.to_dict(orient='records')

    personal_info_list = [', '.join([f'"{key}" : {value}' for key, value in data.items()]) for data in data_dict]
    # data_dict = df.to_dict(orient='records')[0]
    # '. '.join([f'"{key}": {value}' for key, value in data_dict.items()])

    for file_path in file_path_list:
        split_up = os.path.splitext(file_path)
        file_extension = split_up[1]

        analyzer:Analyzer = Analyzer()
        asker:Asker = Asker()
        if file_extension == '.pdf':
            analyzer = PdfAnalyzer()
        elif file_extension in ['.doc', '.docx']:
            analyzer = WordAnalyzer()
        elif file_extension in ['.xlsx', '.xls']:
            analyzer = ExcelAnalyzer()

        questions = analyzer.analyze(file_path)
        qas = asker.ask(questions, personal_info_list)

        return qas
