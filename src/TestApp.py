import streamlit as st
import openai
import os
import pandas as pd
from io import BytesIO
import xlsxwriter

api_key = os.getenv("OPENAI_API_KEY")
client = openai.OpenAI(api_key=api_key)

st.title("Test Suite")

# First area: Document uploader
st.header("1. Upload a Document")
uploaded_file = st.file_uploader("Choose a document", type=["txt", "pdf", "docx"])

document_text = ""
if uploaded_file is not None:
    if uploaded_file.type == "text/plain":
        document_text = uploaded_file.read().decode("utf-8")
    elif uploaded_file.type == "application/pdf":
        from PyPDF2 import PdfReader

        reader = PdfReader(uploaded_file)
        document_text = ""
        for page in range(len(reader.pages)):
            document_text += reader.pages[page].extract_text()
    elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        import docx

        doc = docx.Document(uploaded_file)
        for para in doc.paragraphs:
            document_text += para.text

option = st.sidebar.selectbox(
    "Choose an action",
    ("Generate test cases", "Generate test scenarios", "Generate test script", "Business Analyst", "Test Document")
)


# Function to handle programming language selection
def handle_language_selection(language, script):
    if language == "Python":
        st.write("You selected Python. Here is the pseudo code example:")
        st.code(script)
    elif language == "Java":
        st.write("You selected Java. Here is the pseudo code example:")
        st.code(script)
    elif language == "C++":
        st.write("You selected C++. Here is the pseudo code example:")
        st.code(script)


# Function to create a downloadable Excel file
def create_download_link(df, filename):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    processed_data = output.getvalue()
    st.download_button(
        label="Download Excel file",
        data=processed_data,
        file_name=filename,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# Show programming language dropdown only if "Generate test script" is selected
if option == "Generate test script":
    language_option = st.sidebar.selectbox(
        "Select a programming language",
        ("Python", "Java", "C++")
    )

messages = [
    {"role": "system", "content": "You are a document synthesizer. Please synthesize this document"},
    {"role": "user", "content": document_text}
]

if option == "Generate test cases":
    if document_text:
        messages = [
            {"role": "system",
             "content": "You are an expert level test cases generator. Provide multiple test cases. These test cases should be in table format  with the following row labels - Test Case ID, Test Scenario, Test Case, Pre-condition, Test Steps, Test Data, Expected Result, Actual Result, Status (Pass/Fail) and also should be in downloadable excel format. Also under this provide Gherkin Format for every test case"},
            {"role": "user", "content": document_text}
        ]
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=messages,
            temperature=0
        )
        test_cases_content = response.choices[0].message.content.strip()
        st.write(test_cases_content)

        # Assuming test_cases_content is a string that can be converted to a DataFrame
        # Parsing the response into a list of dictionaries
        test_cases_lines = test_cases_content.split('\n')
        test_cases = []
        headers = ["Test Case ID", "Test Scenario", "Test Case", "Pre-condition", "Test Steps", "Test Data",
                   "Expected Result", "Actual Result", "Status (Pass/Fail)"]

        for line in test_cases_lines:
            columns = line.split('|')
            if len(columns) == 10:  # Adjust according to the actual format
                test_cases.append({
                    headers[i]: columns[i].strip() for i in range(len(headers))
                })

        df = pd.DataFrame(test_cases)
        if not df.empty:
            create_download_link(df)
        else:
            st.warning("No test cases generated or parsed correctly.")
    else:
        st.warning("Please upload a document first!")

if option == "Generate test scenarios":
    if document_text:
        messages = [
            {"role": "system",
             "content": "You are an expert level test scenario generator. Please generate complete test scenarios for - Positive test scenarios, Negative test scenarios, Edge test scenarios."},
            {"role": "user", "content": document_text}
        ]
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=messages,
            temperature=0
        )
        st.write(response.choices[0].message.content.strip())

        # Convert the scenarios to a DataFrame
        test_scenarios_content = response.choices[0].message.content.strip()
        df = pd.DataFrame({'Test Scenarios': test_scenarios_content.split('\n')})
        if not df.empty:
            create_download_link(df, 'test_scenarios.xlsx')
        else:
            st.warning("No test scenarios generated or parsed correctly.")
    else:
        st.warning("Please upload a document first!")

if option == "Generate test script":
    if document_text:
        messages = [
            {"role": "system",
             "content": "You are an expert level test script generator. Please generate test scripts for the given document."},
            {"role": "user", "content": document_text}
        ]
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=messages,
            temperature=0
        )
        generated_script = response.choices[0].message.content.strip()

        # Prepare pseudo code based on the selected programming language
        language_messages = [
            {"role": "system",
             "content": f"You are an expert in {language_option}. Convert the following test script into {language_option} pseudo code. The entire pseudo code should be in programming language without any error. It should be a complete script."},
            {"role": "user", "content": generated_script}
        ]
        pseudo_response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=language_messages,
            temperature=0
        )
        pseudo_code = pseudo_response.choices[0].message.content.strip()

        st.write(generated_script)
        # Handle programming language selection
        handle_language_selection(language_option, pseudo_code)
    else:
        st.warning("Please upload a document first!")

if option == "Business Analyst":
    if document_text:
        messages = [
            {"role": "system",
             "content": "You are an expert level business analyst. Generate user stories.For the generated user stories, please include the acceptance criteria based on the functionality."},
            {"role": "user", "content": document_text}
        ]
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=messages,
            temperature=0
        )
        st.write(response.choices[0].message.content.strip())
    else:
        st.warning("Please upload a document first!")

if option == "Test Document":
    if document_text:
        messages = [
            {"role": "system", "content": "You are an expert level test document generator. Provide a test document."},
            {"role": "user", "content": document_text}
        ]
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=messages,
            temperature=0
        )
        st.write(response.choices[0].message.content.strip())
    else:
        st.warning("Please upload a document first!")