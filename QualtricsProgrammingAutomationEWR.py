import streamlit as st
import re
import string

def read_text_file(file_path):
    unwanted_start_line = "Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License https://products.aspose.com/words/temporary-license/"
    unwanted_line = "This document was truncated here because it was created in the Evaluation Mode."
    unwanted_end_line = "Evaluation Only. Created with Aspose.Words. Copyright 2003-2024 Aspose Pty Ltd."

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    # Remove the unwanted lines
    lines = [line for line in lines if unwanted_start_line not in line and unwanted_end_line not in line and unwanted_line not in line]
    
    return ''.join(lines)

def identify_question_type(question_text):
    text = question_text.lower()
    if "matrix" in text:
        return "Matrix"
    elif "open-end" in text:
        return "TextEntry"
    elif "rank order" in text:
        return "RankOrder"
    elif "constant sum" in text:
        return "ConstantSum"
    elif "[db]" in text:
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
            cleaned_choice = number_sequence_pattern.sub('', choice).strip()
            processed_choices.append(cleaned_choice)
        else:
            if trailing_number_pattern.match(choice.split()[-1]):
                processed_choices.append(' '.join(choice.split()[:-1]).strip())
            else:
                processed_choices.append(choice)
    
    return processed_choices

def convert_to_qualtrics_format(input_text):
    lines = input_text.strip().split('\n')
    qualtrics_lines = ['[[AdvancedFormat]]']
    question_counter = 0
    intro_counter = 0
    choices = []
    matrix_statements = []
    matrix_scale_points = []
    question_type = None
    matrix_answer_mode = False
    previous_line_was_db = False

    for i, line in enumerate(lines):
        line = line.strip()

        if not line:
            continue

        # Identify if the current line is a new question
        if re.match(r'^\d+\.', line):
            # If there are pending choices or matrix items, append them before processing new question
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
            if question_type != "DB":
                qualtrics_lines.append(f'[[Question:{question_type}]]')
                question_counter += 1
                qualtrics_lines.append(f'[[ID:Q{question_counter}]]')
                qualtrics_lines.append(cleaned_question_text)
            matrix_answer_mode = False
            previous_line_was_db = False

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

        # Process DB block separately to ensure it doesn't interrupt current question
        if "[db]" in line.lower():
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
            if not previous_line_was_db:
                intro_counter += 1
                qualtrics_lines.append('[[Question:DB]]')
                qualtrics_lines.append(f'[[ID:Intro{intro_counter}]]')
                qualtrics_lines.append(remove_square_bracket_content(line))
                previous_line_was_db = True
    
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
    
    # Remove lines ending with '[DB]'
    cleaned_qualtrics_lines = [line for line in qualtrics_lines if not line.strip().endswith('[DB]')]

    return '\n'.join(cleaned_qualtrics_lines)

def remove_blank_lines(text):
    lines = text.split('\n')
    cleaned_lines = [line for line in lines if line.strip()]
    return '\n'.join(cleaned_lines)

def remove_initial_content(text):
    # Find the position of the first question tag
    first_question_match = re.search(r'\[\[Question:(MC|DB|Matrix)\]\]', text)
    if first_question_match:
        first_question_index = first_question_match.start()
        return '[[AdvancedFormat]]\n' + text[first_question_index:]
    else:
        return text

def main():
    st.title('Qualtrics Programming Automation!')

    uploaded_file = st.file_uploader("Upload a Text file", type="txt")
    output_file_name = st.text_input("Enter the output file name:", "output_file.txt")

    if uploaded_file is not None and output_file_name:
        # # Save the uploaded TXT file to a temporary location
        # with open("uploaded_input.txt", "wb") as f:
        #     f.write(uploaded_file.getbuffer())

        # Read and process the text file
        input_text = read_text_file("uploaded_file.txt")
        converted_text = convert_to_qualtrics_format(input_text)
        cleaned_text = remove_blank_lines(converted_text)
        final_text = remove_initial_content(cleaned_text)
        final_text = ''.join(filter(lambda x: x in string.printable, final_text))
        
        if st.button("Convert into desired text file"):
            # Save the final text to the specified output file
            with open(output_file_name, "w", encoding='utf-8') as output_file:
                output_file.write(final_text)

            st.success(f"File '{output_file_name}' has been created successfully.")
            st.download_button("Download the file", final_text, file_name=output_file_name)

if __name__ == "__main__":
    main()
