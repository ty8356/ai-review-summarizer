import os
import sys, configparser
import fitz  # PyMuPDF
import pandas as pd
import re
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openai import OpenAI

config = configparser.ConfigParser()
config.read('config.ini')
os.environ['OPENAI_API_KEY'] = config.get('DEFAULT', 'gpt_token')
client = OpenAI(api_key=os.environ['OPENAI_API_KEY'])

chatgptprompt = 'I am writing a literature review about the impact of "hope" on chronic illness. Please review the article text provided and pull out the following information with no fluff: title, author names, month-year of publication, country, citation in APA format, chronic disease studied, type of research study, species studied (human, mouse, rat, other), scales used to measure hope/depression/anxiety/etc (list them in a single line separated by commas), number of participants, was quality of life improved or worsened, was functional ability improved or worsened, was treatment adherence improved or worsened, was survival impacted. Additionally, write a brief summary including the following information: methods, results, clinical significance, and any limitations of the study. DO NOT use any markdown around any words in the response. Leave it as plain text ONLY. The results should be in the template as follows: - Title: - Authors: - Month-Year: - Country: - Citation: - Disease: - Study Type: - Species: - Scales Used: - Participants: - Quality of Life: - Functional Ability: - Treatment Adherence: - Survival: - Summary: \n Here is the article text: '

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with fitz.open(pdf_path) as doc:
            for page in doc:
                text += page.get_text()
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    return text

def process_with_chatgpt(text):
    try:
        completion = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "developer", "content": "You are a helpful assistant."},
                {"role": "user", "content": chatgptprompt + text}
            ]
        )

        return completion.choices[0].message.content
    except Exception as e:
        print(f"Error with OpenAI API: {e}")
        return ""
    
def construct_excel_row(text, file_name, raw_file_text):
    patterns = {
        "Title": r"Title: (.+?)(?=\n|\Z)",
        "Authors": r"Authors: (.+?)(?=\n|\Z)",
        "Month-Year": r"Month-Year: (.+?)(?=\n|\Z)",
        "Country": r"Country: (.+?)(?=\n|\Z)",
        "Citation": r"Citation: (.+?)(?=\n|\Z)",
        "Disease": r"Disease: (.+?)(?=\n|\Z)",
        "Study Type": r"Study Type: (.+?)(?=\n|\Z)",
        "Species": r"Species: (.+?)(?=\n|\Z)",
        "Scales Used": r"Scales Used: ([\s\S]+?)(?=\n|\Z)",
        "Participants": r"Participants: (.+?)(?=\n|\Z)",
        "Quality of Life": r"Quality of Life: (.+?)(?=\n|\Z)",
        "Functional Ability": r"Functional Ability: (.+?)(?=\n|\Z)",
        "Treatment Adherence": r"Treatment Adherence: (.+?)(?=\n|\Z)",
        "Survival": r"Survival: (.+?)(?=\n|\Z)",
        "Summary": r"Summary: ([\s\S]+?)(?=#####|\Z)"
    }

    extracted_sections = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            extracted_sections[key] = match.group(1).strip()

    # print(text)

    record = {
        "PMID": file_name.split('.')[0],
        "Title": extracted_sections['Title'],
        "Authors": extracted_sections['Authors'],
        "Month-Year": extracted_sections['Month-Year'],
        "Country": extracted_sections['Country'],
        "Study Type": extracted_sections['Study Type'],
        "Species": extracted_sections['Species'],
        "Participants": extracted_sections['Participants'],
        "Scales": extracted_sections['Scales Used'],
        "Disease": extracted_sections['Disease'],
        "QoL": extracted_sections['Quality of Life'],
        "Func Ability": extracted_sections['Functional Ability'],
        "Tx Adherence": extracted_sections['Treatment Adherence'],
        "Survival": extracted_sections['Survival'],
        "Summary": extracted_sections['Summary'],
        "Excluded": '0',
        "Citation": extracted_sections['Citation'],
        "Text": ILLEGAL_CHARACTERS_RE.sub(r'', raw_file_text)
    }

    return record

def write_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

def main(input_dir, output_file):
    data = []
    for file_name in os.listdir(input_dir):
        if file_name.endswith('.pdf'):
            pdf_path = os.path.join(input_dir, file_name)
            print(f"Processing {file_name}...")
            text = extract_text_from_pdf(pdf_path)
            processed_data = process_with_chatgpt(text)
            record = construct_excel_row(str(processed_data), file_name, text)
            data.append(record)

    write_to_excel(data, output_file)
    print(f"Data written to {output_file}")

if __name__ == "__main__":
    input_directory = config.get('DEFAULT', 'input_directory')
    output_excel = "summarized-pdfs.xlsx"
    main(input_directory, output_excel)