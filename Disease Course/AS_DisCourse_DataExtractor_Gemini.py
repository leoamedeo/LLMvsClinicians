import os
import json
import requests
import datetime
import mammoth
import re
import pandas as pd
import time
import google.generativeai as genai

# Configuration - Replace with your own API key
API_KEY = "YOUR_GOOGLE_GEMINI_API_KEY_HERE"

# Model configuration
model_name = "gemini-2.0-flash-exp"  # Options: gemini-1.5-flash, gemini-2.0-flash-exp
llm_nickname = "gemini-exp"

# Rate limiting: Adjust requests per minute - 15 for flash1.5, 10 for flash2.0
rate_limit = 10

# Directory configuration - Update these paths for your environment
main_folder = r"path/to/your/input/documents"
output_folder = r"path/to/your/output/files"

# Processing parameters
n_subfolders = 211  # Number of subfolders to process
sections = "all_sections"  # Options: ["Verlauf", "Befund", "Beurteilung"] or "all_sections"
iteration_n = 3  # Number of processing iterations

# LLM instruction prompt for medical data extraction
input_instructions = (
    "In the provided document, please analyze the patient's disease course based on the text within <context> and determine the correct values for each of the following requested data point. After providing a summary of the patients disease course, provide a final answer in the form of bullet points following the same structure of the datapoints below."
    "- Any improvement of pain after first surgery (Yes/No/Not provided)"
    "- Completely free of pain after first surgery (Yes/No/Not provided)"
    "- Symptom recurrence after first surgery (Yes/No, if it is not explicitly mentioned, assume there was no recurrence)"
    "- A second surgery was carried out (Yes/No, if it is not explicitly mentioned, assume there was no second surgery)"
    "- Free of pain after second surgery (Yes/No/Not provided)"
    "- Recurrence after second surgery (Yes/No/Not provided)"
    "- Thermocoagulation was carried out (Yes/No/Not provided)"
)

def combine_word_documents(subfolder_path):
    """
    Combine the raw text of all .docx files in the specified folder into a single string.
    Extracts text between 'diagnos' and 'grüße' markers and cleans formatting.
    Note: Assumes documents have already been anonymized.
    """
    combined_text = []
    subfolder_name = os.path.basename(subfolder_path)

    # Process all Word documents in the subfolder
    for file_name in os.listdir(subfolder_path):
        if file_name.endswith(".docx"):
            file_path = os.path.join(subfolder_path, file_name)
            try:
                # Extract raw text from Word document
                with open(file_path, "rb") as docx_file:
                    result = mammoth.extract_raw_text(docx_file)
                    text = result.value

                    # Extract relevant medical content between markers
                    match = re.search(r'(diagnos.*?grüße)', text, re.IGNORECASE | re.DOTALL)
                    if match:
                        extracted_text = match.group(1).strip()
                        # Clean up excessive newlines
                        extracted_text = re.sub(r'\n{2,}', '\n', extracted_text)
                        combined_text.append(extracted_text)
                        
            except Exception as e:
                print(f"Error processing file {file_name}: {e}")

    return "\n".join(combined_text)

def extract_section(text, sections):
    """
    Extract specified sections from the text, including the title and paragraphs 
    immediately after each matching section heading.
    """
    paragraphs = text.split("\n")
    section_content = []

    # Return all content if no specific sections requested
    if sections == "all_sections":
        section_content = paragraphs
    else:
        # Extract only specified sections
        for i, paragraph in enumerate(paragraphs):
            if any(section_heading.lower() in paragraph.lower() for section_heading in sections):
                section_content.append(paragraph)  # Add the section heading
                
                # Include the next paragraph if it exists and is not empty
                if i + 1 < len(paragraphs) and paragraphs[i + 1].strip():
                    section_content.append(paragraphs[i + 1])

    return "\n".join(section_content)

def standardize_variable_name(variable):
    """
    Normalize different spellings and phrasings for the same data points to ensure
    consistent mapping of extracted information.
    """
    variable = variable.lower().strip()
    
    # Explicit matches for precise mapping
    if variable == 'free of pain after second surgery':
        return 'Free of pain after second surgery'
    elif variable == 'recurrence after second surgery':
        return 'Recurrence after second surgery'
    
    # Pattern-based matching for common variations
    if 'improvement' in variable or 'betterment' in variable:
        return 'Symptom-improvement after 1. surgery'
    elif ('free of pain after first' in variable or
          'painfree after first' in variable or
          'painfree after 1' in variable or
          'free of pain after 1' in variable):
        return 'Free of pain after first surgery'
    elif 'recurrence after first' in variable or 'recurrence after 1' in variable:
        return 'Recurrence after first surgery'
    elif ('second surgery' in variable or '2nd surgery' in variable or 
          '2. surgery' in variable or 'a second surgery was carried out' in variable):
        return 'Second surgery'
    elif 'thermocoag' in variable or 'coagulation' in variable:
        return 'Thermocoagulation'
    
    return None

def extract_bullet_points(response_text):
    """
    Extract relevant bullet points from the LLM response text, handling various
    formatting patterns including 'Final answer:' sections and embedded answers.
    """
    relevant_points = {}
    lines = response_text.split('\n')

    # Regex patterns for different response formats
    bullet_point_pattern = re.compile(
        r'^\s*(?:\*\s*|\*\*\*|\*+\s*\*\*?|[\u2022\-]|[1-5]\.)\s*\**\s*([^:]+?)\**\s*:\s*(?:\'|\")?(Yes|No|Ja|Nein|provided|know)(?:\'|\")?.*',
        re.IGNORECASE
    )

    final_answer_pattern = re.compile(r'\bfinal answer\b', re.IGNORECASE)
    standalone_answer_pattern = re.compile(r'^\s*([A-Za-z\s]+):\s*(Yes|No|Ja|Nein|provided|know)', re.IGNORECASE)

    current_variable = None
    in_final_answer_section = False

    # Process each line of the response
    for line in lines:
        line = line.strip()

        # Check if we've reached the final answer section
        if final_answer_pattern.search(line):
            in_final_answer_section = True
            continue

        # Extract bullet point formatted answers
        match = bullet_point_pattern.match(line)
        if match:
            current_variable = match.group(1).strip().lower()
            value = match.group(2).strip()

            # Normalize response values
            value = normalize_response_value(value)
            normalized_variable = standardize_variable_name(current_variable)

            if normalized_variable:
                relevant_points[normalized_variable] = value

        # Extract standalone answers in final answer section
        elif in_final_answer_section:
            match = standalone_answer_pattern.match(line)
            if match:
                variable = match.group(1).strip().lower()
                value = match.group(2).strip()
                
                value = normalize_response_value(value)
                normalized_variable = standardize_variable_name(variable)

                if normalized_variable:
                    relevant_points[normalized_variable] = value

    return relevant_points

def normalize_response_value(value):
    """
    Normalize response values to standard format.
    """
    value_lower = value.lower()
    if value_lower == "ja":
        return "Yes"
    elif value_lower == "nein":
        return "No"
    elif value_lower in ["provided", "know"]:
        return "Zero"
    return value

def save_to_excel(data, output_file):
    """
    Save the extracted data to Excel file, appending to existing data if file exists.
    """
    columns = ['Surname', 'Name', 'Symptom-improvement after 1. surgery', 
               'Free of pain after first surgery', 'Recurrence after first surgery',
               'Second surgery', 'Free of pain after second surgery', 
               'Recurrence after second surgery', 'Thermocoagulation', 
               'AI response', 'Parsed Data']

    # Load existing data or create new DataFrame
    try:
        df_existing = pd.read_excel(output_file)
    except FileNotFoundError:
        df_existing = pd.DataFrame(columns=columns)

    # Create DataFrame with new data
    df_new = pd.DataFrame([data], columns=columns)
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined.to_excel(output_file, index=False)

def process_and_run_llm_for_subfolder(subfolder_path, api_key, model_name=model_name):
    """
    Process documents in a subfolder and extract medical data using LLM.
    Returns extracted data as dictionary.
    """
    subfolder_name = os.path.basename(subfolder_path)
    
    # Note: Assuming patient identifiers have been pre-anonymized
    surname, name = "ANONYMIZED", "ANONYMIZED"

    # Combine and process Word documents
    combined_text = combine_word_documents(subfolder_path)
    context = extract_section(combined_text, sections)

    # Save context for reference
    context_file_path = os.path.join(subfolder_path, "context.txt")
    with open(context_file_path, "w", encoding="utf-8") as file:
        file.write(context)

    # Create formatted query for LLM
    formatted_query = (
        "You are a capable physician assistant. You are given medical documents of patients who underwent microvascular decompression surgery. contained within <context> tags. Your role is to analyze the patient's disease course carefully and then provide specific data points as your final answer. You must strictly rely on the content provided inside <context></context> XML tags.\n. Never fabricate or add external information."
        " If data is missing, respond with I don't know or Not provided.\n"
        f"<context>\n{context}\n</context>\n\n"
        "When answering the user:\n"
        "- If you don't know, just say that you don't know.\n"
        "- If the context doesn't give you the information asked for, say so.\n"
        "Avoid mentioning that you obtained the information from the context.\n"
        "Always strictly stand by the information given in the context.\n\n"
        f"Given the context information, answer the query.\nQuery: {input_instructions}"
    )

    # Configure and run Gemini model
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    try:
        # Generate response from Gemini
        response = model.generate_content(formatted_query)
        full_response = response.text.strip()

    except Exception as e:
        print(f"Error processing '{subfolder_name}': {e}")
        return None

    # Extract structured data from response
    relevant_points = extract_bullet_points(full_response)
    data = relevant_points
    data['Surname'] = surname
    data['Name'] = name
    data['AI response'] = full_response
    data['Parsed Data'] = context

    print(f"Response saved for {subfolder_name}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    return data

def main():
    """
    Main execution function that processes all subfolders through multiple iterations
    with rate limiting to respect API constraints.
    """
    # Get list of subfolders to process
    subfolders = [f for f in os.listdir(main_folder) 
                  if os.path.isdir(os.path.join(main_folder, f))][:n_subfolders]

    # Process through multiple iterations
    for i in range(1, iteration_n + 1):
        output_file = os.path.join(output_folder, f"output_{llm_nickname}_{i}.xlsx")
        all_data = []
        start_time = time.time()

        print(f"Starting iteration {i} of {iteration_n}")

        # Process each subfolder
        for j, subfolder in enumerate(subfolders, start=1):
            subfolder_path = os.path.join(main_folder, subfolder)
            result = process_and_run_llm_for_subfolder(subfolder_path, API_KEY, model_name)
            
            if result:
                all_data.append(result)

            # Rate limiting: pause after every rate_limit requests
            if j % rate_limit == 0:
                elapsed_time = time.time() - start_time
                remaining_time = 61 - elapsed_time  # Wait for full minute
                
                if remaining_time > 0:
                    print(f"Rate limit reached: Waiting {remaining_time:.2f} seconds before continuing...")
                    time.sleep(remaining_time)
                
                start_time = time.time()  # Reset timer

        # Save results for this iteration
        if all_data:
            df = pd.DataFrame(all_data)
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            df.to_excel(output_file, index=False)
            print(f"Iteration {i} completed. Output saved as {output_file}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    main()