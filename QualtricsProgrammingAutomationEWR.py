import streamlit as st
import aspose.words as aw
import re
import string
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from time import sleep
import os
import subprocess

# Functions to process text
def read_text_file(file_path):
    unwanted_start_line = "Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License https://products.aspose.com/words/temporary-license/"
    unwanted_end_line = "Evaluation Only. Created with Aspose.Words. Copyright 2003-2024 Aspose Pty Ltd."

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    # Remove the unwanted lines
    lines = [line for line in lines if unwanted_start_line not in line and unwanted_end_line not in line]
    
    return ''.join(lines)

def identify_question_type(question_text):
    if "matrix" in question_text.lower():
        return "Matrix"
    elif "open-end" in question_text.lower() or "text" in question_text.lower():
        return "TextEntry"
    elif "rank order" in question_text.lower():
        return "RankOrder"
    elif "constant sum" in question_text.lower():
        return "ConstantSum"
    elif "descriptive block" in question_text.lower():
        return "DB"
    else:
        return "MC"

def remove_square_bracket_content(text):
    return re.sub(r'\[.*?\]', '', text).strip()

def process_choices(choices, question_type):
    processed_choices = []
    number_sequence_pattern = re.compile(r'(\d+\s*)+(-99\s*)?$')
    trailing_number_pattern = re.compile(r'^\s*\d+\s*$')

    for choice in choices:
        if question_type == "Matrix":
            # Remove number sequences
            cleaned_choice = number_sequence_pattern.sub('', choice).strip()
            processed_choices.append(cleaned_choice)
        else:
            # Remove trailing numbers
            if trailing_number_pattern.match(choice.split()[-1]):
                processed_choices.append(' '.join(choice.split()[:-1]).strip())
            else:
                processed_choices.append(choice)
    
    return processed_choices

def convert_to_qualtrics_format(input_text):
    lines = input_text.strip().split('\n')
    qualtrics_lines = ['[[AdvancedFormat]]']
    question_counter = 0
    question_text = ""
    choices = []
    matrix_statements = []
    matrix_scale_points = []
    question_type = None
    matrix_answer_mode = False

    for i, line in enumerate(lines):
        line = line.strip()

        if not line:
            continue

        # Identify if the current line is a new question
        if re.match(r'\d+\.', line):
            if question_counter > 0:
                if choices:
                    qualtrics_lines.append("[[Choices]]")
                    qualtrics_lines.extend(process_choices(choices, question_type))
                    choices = []
                if matrix_statements and matrix_scale_points:
                    qualtrics_lines.append("[[Answers]]")
                    qualtrics_lines.extend(matrix_statements)
                    qualtrics_lines.append("[[Choices]]")
                    qualtrics_lines.extend(process_choices(matrix_scale_points, question_type))
                    matrix_statements = []
                    matrix_scale_points = []
                qualtrics_lines.append('[[PageBreak]]')

            question_text = line
            question_type = identify_question_type(question_text)
            cleaned_question_text = remove_square_bracket_content(question_text)
            if question_type == "DB":
                qualtrics_lines.append('[[DB]]')
                qualtrics_lines.append(cleaned_question_text.split(" ", 1)[1])
            else:
                qualtrics_lines.append(f'[[Question:{question_type}]]')
                qualtrics_lines.append(cleaned_question_text.split(" ", 1)[1])
            question_counter += 1
            matrix_answer_mode = False

        elif re.match(r'\[IF Q\d', line):
            qualtrics_lines.append(f'[[{line}]]')
        else:
            if question_type == "Matrix":
                if re.match(r'\[.*?\]', line):
                    matrix_answer_mode = True
                if matrix_answer_mode:
                    matrix_scale_points.append(line)
                else:
                    matrix_statements.append(line)
            else:
                choices.append(line)
    
    # Append any remaining choices or matrix statements for the last question
    if choices:
        qualtrics_lines.append("[[Choices]]")
        qualtrics_lines.extend(process_choices(choices, question_type))
    if matrix_statements and matrix_scale_points:
        qualtrics_lines.append("[[Answers]]")
        qualtrics_lines.extend(matrix_statements)
        qualtrics_lines.append("[[Choices]]")
        qualtrics_lines.extend(process_choices(matrix_scale_points, question_type))
    qualtrics_lines.append('[[PageBreak]]')
    
    return '\n'.join(qualtrics_lines)

def remove_blank_lines(text):
    lines = text.split('\n')
    cleaned_lines = [line for line in lines if line.strip()]
    return '\n'.join(cleaned_lines)

# Selenium functions to interact with Qualtrics
def login_to_qualtrics(driver, username, password):
    driver.get("https://nea.co1.qualtrics.com/login")
    sleep(2)
    driver.find_element(By.ID, 'UserName').send_keys(username)
    driver.find_element(By.ID, 'UserPassword').send_keys(password)
    driver.find_element(By.ID, 'loginButton').click()
    sleep(5)
    print("Logged in to Qualtrics successfully")

def create_survey(driver, survey_name):
    driver.find_element(By.CSS_SELECTOR, 'button[data-testid="profile-data-create-new-button"]').click()
    sleep(2)
    driver.find_element(By.CSS_SELECTOR, 'button[data-testid="Catalog.Offering.Survey"]').click()
    sleep(2)
    driver.find_element(By.CSS_SELECTOR, 'button[data-testid="Catalog.DetailsPane.CallToAction"]').click()
    sleep(2)
    driver.find_element(By.CSS_SELECTOR, 'input[data-testid="Catalog.GetStartedFlow.Name"]').send_keys(survey_name)
    driver.find_element(By.CSS_SELECTOR, 'button[data-testid="Catalog.GetStartedFlow.Create"]').click()
    sleep(5)
    print(f"Survey '{survey_name}' created successfully")

def import_survey(driver, survey_text_file):
    driver.find_element(By.ID, 'builder-tools-menu').click()
    sleep(2)
    driver.find_element(By.ID, 'import-export-menu').click()
    sleep(2)
    driver.find_element(By.ID, 'import-survey-tool').click()
    sleep(2)
    file_input = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
    file_input.send_keys(survey_text_file)
    sleep(2)
    driver.find_element(By.ID, 'importButton').click()
    sleep(5)
    print("Survey imported successfully")

def install_chrome():
    # Install Chrome
    os.system('wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb')
    os.system('dpkg -i google-chrome-stable_current_amd64.deb')
    os.system('apt-get -f install -y')

def install_chromedriver():
    # Install ChromeDriver
    os.system('wget https://chromedriver.storage.googleapis.com/$(curl -sS chromedriver.storage.googleapis.com/LATEST_RELEASE)/chromedriver_linux64.zip')
    os.system('unzip chromedriver_linux64.zip -d /usr/local/bin/')

def automate_survey_creation(doc_path, username, password, survey_name):
    install_chrome()
    install_chromedriver()

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=ChromeService("/usr/local/bin/chromedriver"), options=chrome_options)
    try:
        login_to_qualtrics(driver, username, password)
        create_survey(driver, survey_name)
        doc = aw.Document(doc_path)
        doc.save("converted_text.txt")
        input_text = read_text_file("converted_text.txt")
        converted_text = convert_to_qualtrics_format(input_text)
        cleaned_text = remove_blank_lines(converted_text)
        cleaned_text = ''.join(filter(lambda x: x in string.printable, cleaned_text))
        survey_text_file = os.path.abspath('tags_text_conv.txt')
        with open(survey_text_file, 'w', encoding='utf-8') as output_file:
            output_file.write(cleaned_text)
        import_survey(driver, survey_text_file)
    finally:
        driver.quit()

# Streamlit app
st.title("Qualtrics Survey Automation")
username = st.text_input("Qualtrics Username")
password = st.text_input("Qualtrics Password", type="password")
survey_name = st.text_input("Survey Name")
doc_file = st.file_uploader("Upload DOCX File", type=["docx"])

if st.button("Create Survey"):
    if username and password and survey_name and doc_file:
        with open("uploaded_file.docx", "wb") as f:
            f.write(doc_file.getbuffer())
        automate_survey_creation("uploaded_file.docx", username, password, survey_name)
        st.success("Survey created and imported successfully!")
    else:
        st.error("Please fill all the fields and upload the DOCX file.")
