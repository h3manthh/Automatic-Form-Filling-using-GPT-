
import streamlit as st
import json
import time
from openai import OpenAI

OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]
GPT_MODEL = st.secrets["GPT_MODEL"]

class Asker:
    def __init__(self) -> None:
        pass

    def change_info_for_person(self, first_answer, person_description):
        messages = [
            {
                "role" : "system",
                "content" : f"I will provide you a json structure of first human information and change it for second human's information\
                    Remember! Follow 3 rules.\
                        1) Do not answer until you think the result is accurate\
                        2) First human description starts with <<< and ends with >>>\
                        3) Second human description starts with ((( and ends with )))\
                    "
            },
            {
                "role" : "user",
                "content" : f""" <<< {first_answer} >>>\
                      ((({person_description})))"""
            }
        ]

        tools = [
            {
                "type" : "function",
                "function" : {
                    "name" : "change_person_info",
                    "description" : "Provide answers for the form fields for second person",
                    "parameters" : {
                        "type" : "object",
                        "properties" : {
                            "items" : {
                                "type" : "object",
                                "properties" : {
                                    "request_form" : {
                                        "type" : "string", 
                                        "description" : "Form field (eg. <Forename 1>, <Have you been a student at this University before? (tick)>, <Address Line 2>, <Lecturer>, <Assessment Title>)"
                                    },
                                    "answer" : {
                                        "type" : "string",
                                        "description" : "Example answer for corresponding form field (eg. <David>, <Yes>, <911 Hillside Dr, Kodiak, Alaska 99615, USA>, <Math>, <Midterm Exam: Introduction to Python Programming>)"
                                    }
                                },
                                "required" : ["request_form", "answer"]
                            }
                        },
                        "required" : ["items"]
                    }
                }
            }
        ]

        print(messages, tools)

        client = OpenAI(api_key=OPENAI_API_KEY)

        response = client.chat.completions.create(
            model=GPT_MODEL,
            messages=messages,
            tools=tools,
            tool_choice="auto",  # auto is default, but we'll be explicit
        )

        response_message = response.choices[0].message
        tool_calls = response_message.tool_calls
        if tool_calls:
            for tool_call in tool_calls:
                function_name = tool_call.function.name
                if function_name == 'change_person_info':
                    function_args = json.loads(tool_call.function.arguments)
                    return function_args.get('items', [])
        return []

    def ask_one_person(self, questions, person_description):
        messages = [
            {
                "role" : "system",
                "content" : f"I will give you a list of strings which are extracted from form input file. \
                    Remember! Follow 4 rules.\
                        1) For each The given string, it starts with <<< and ends with >>>\
                        2) Analyze the whole list and extract fit input form fields and generate any answer for the form fields\
                        3) Do not answer until you think the result is accurate\
                    Reference person's information here {person_description}\
                        "
            },
            {
                "role" : "user",
                "content" : " , ".join(f"<<<{question}>>>"for question in questions)
            }
        ]
        tools = [
            {
                "type" : "function",
                "function" : {
                    "name" : "extract_forms_and_fill_answers",
                    "description" : "Identify form fields, for example,\
                          ['Forename 1', 'Address Line 1', 'If yes, please give your Student Identification number, if known', 'Student Name', 'Unit Title']\
                         Then, provide answers for the form fields",
                    "parameters" : {
                        "type" : "object",
                        "properties" : {
                            "items" : {
                                "type" : "object",
                                "properties" : {
                                    "request_form" : {
                                        "type" : "string", 
                                        "description" : "Form field (eg. <Forename 1>, <Have you been a student at this University before? (tick)>, <Address Line 2>, <Lecturer>, <Assessment Title>)"
                                    },
                                    "answer" : {
                                        "type" : "string",
                                        "description" : "Example answer for corresponding form field (eg. <David>, <Yes>, <911 Hillside Dr, Kodiak, Alaska 99615, USA>, <Math>, <Midterm Exam: Introduction to Python Programming>)"
                                    }
                                },
                                "required" : ["request_form", "answer"]
                            }
                        },
                        "required" : ["items"]
                    }
                }
            }
        ]

        client = OpenAI(api_key=OPENAI_API_KEY)

        response = client.chat.completions.create(
            model=GPT_MODEL,
            messages=messages,
            tools=tools,
            tool_choice="auto",  # auto is default, but we'll be explicit
        )

        response_message = response.choices[0].message
        tool_calls = response_message.tool_calls
        if tool_calls:
            for tool_call in tool_calls:
                function_name = tool_call.function.name
                if function_name == 'extract_forms_and_fill_answers':
                    function_args = json.loads(tool_call.function.arguments)
                    return function_args.get('items', [])
        return []
    
    def ask(self, questions, person_infos):
        answers = []
        for index, person_info in enumerate(person_infos):
            for i in range(3):
                try:
                    if index == 0:
                        one_person_answer = self.ask_one_person(questions, person_info)

                        answers.append(one_person_answer)
                        time.sleep(2)
                    else:
                        one_person_answer = self.change_info_for_person(answers[0], person_info)
                        answers.append(one_person_answer)
                        time.sleep(2)
                    break
                except Exception as e:
                    print(e)
                    return []
        return answers