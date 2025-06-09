import os
import json
import requests
import datetime
import mammoth
import re
import pandas as pd
import time
import google.generativeai as genai

# Configuration - Replace with your actual API key
API_KEY = "YOUR_GOOGLE_GEMINI_API_KEY_HERE"

# Model configuration
model_name = "gemini-1.5-flash"  # Options: gemini-1.5-flash, gemini-2.0-flash-exp
llm_nickname = "geminiflash"

# Rate limiting: Adjust requests per minute to 15 for flash1.5 and to 10 for flash2.0
rate_limit = 15

# Configurable parameters - Update these paths for your system
main_folder = "/path/to/your/input/documents"  # Folder containing subfolders with .docx files
output_folder = "/path/to/your/output/files"   # Folder where results will be saved

# Processing parameters
n_subfolders = 212  # Number of subfolders to process
sections = "all_sections"  # Use "all_sections" or specify sections like ["Section1", "Section2"]
iteration_n = 1  # Number of processing iterations

# Instructions for the AI model to extract preoperative symptoms
input_instructions = (
    "In the provided document please look for the following preoperative symptoms: "
    "Sudden Severe Facial Pain, Facial Numbness, Vertigo, Lacrimation, Facial Muscle Spasm, and Other (related to trigeminal neuralgia)."
    "In your Answer, first reason whether any of your findings can really be considered a preoperative symptom. "
    "Focus on the fact that it is only a preoperative symptom only if it was already present before the FIRST surgery. "
    "Always consider the first surgery if the patient underwent multiple ones. "
    "Consider the symptom only if it is explicitly mentioned in the documents; if it is not mentioned, always assume the symptom is not present."
    "After reasoning about your findings, provide a final answer in the form of bullet points with 'Name of the Symptom': 'Yes' or 'No' for each individual point."
)

def combine_word_documents(subfolder_path):
    """
    Combine the raw text of all .docx files in the specified folder into a single string.
    Extracts text between 'diagnos' and 'grüße' patterns and cleans formatting.
    
    Args:
        subfolder_path (str): Path to subfolder containing .docx files
        
    Returns:
        str: Combined text from all documents
    """
    combined_text = []

    for file_name in os.listdir(subfolder_path):
        if file_name.endswith(".docx"):
            file_path = os.path.join(subfolder_path, file_name)
            try:
                # Extract raw text from Word document using mammoth
                with open(file_path, "rb") as docx_file:
                    result = mammoth.extract_raw_text(docx_file)
                    text = result.value

                    # Extract text between "diagnos" and "grüße" (case-insensitive), including delimiters
                    match = re.search(r'(diagnos.*?grüße)', text, re.IGNORECASE | re.DOTALL)
                    if match:
                        extracted_text = match.group(1).strip()

                        # Clean up excessive newlines
                        extracted_text = re.sub(r'\n{2,}', '\n', extracted_text)

                        combined_text.append(extracted_text)
            except Exception as e:
                print(f"Error processing file {file_name}: {e}")

    return "\n".join(combined_text)  # Join all text with newlines

def extract_section(text, sections):
    """
    Extracts specified sections from the text, including the title and paragraphs 
    immediately after each matching section heading.
    
    Args:
        text (str): Input text to extract sections from
        sections (str or list): Either "all_sections" or list of section headings to extract
        
    Returns:
        str: Extracted section content
    """
    paragraphs = text.split("\n")
    section_content = []

    if sections == "all_sections":
        section_content = paragraphs
    else:
        # Extract specific sections based on headings
        for i, paragraph in enumerate(paragraphs):
            if any(section_heading.lower() in paragraph.lower() for section_heading in sections):
                section_content.append(paragraph)  # Add the section heading

                # Include the next paragraph if it exists and is not empty
                if i + 1 < len(paragraphs) and paragraphs[i + 1].strip():
                    section_content.append(paragraphs[i + 1])

    return "\n".join(section_content)

def extract_bullet_points(response_text):
    """
    Extract relevant bullet points from the AI response text, including 'Final answer:' 
    formatted lists and embedded answers.
    
    Args:
        response_text (str): AI model response text
        
    Returns:
        dict: Dictionary mapping standardized symptom names to Yes/No values
    """
    relevant_points = {}
    lines = response_text.split('\n')

    # Regex pattern to match various bullet point formats and numbered lists
    bullet_point_pattern = re.compile(
        r'^\s*(?:\*\s*|\*\*\*|\*+\s*\*\*?|[\u2022\-]|[1-5]\.)\s*\**\s*([^:]+?)\**\s*:\s*(?:\'|\")?(Yes|No|Ja|Nein)(?:\'|\")?.*',
        re.IGNORECASE
    )

    # Pattern to detect "Final answer" sections
    final_answer_pattern = re.compile(r'\bfinal answer\b', re.IGNORECASE)

    # Pattern to match standalone answers without bullets (e.g., "Symptom: No")
    standalone_answer_pattern = re.compile(r'^\s*([A-Za-z\s]+):\s*(Yes|No|Ja|Nein)', re.IGNORECASE)

    current_variable = None
    in_final_answer_section = False

    for line in lines:
        line = line.strip()

        # Check if we're entering a "Final answer" block
        if final_answer_pattern.search(line):
            in_final_answer_section = True
            continue

        # Try to match standard bullet point pattern
        match = bullet_point_pattern.match(line)
        if match:
            current_variable = match.group(1).strip().lower()
            value = match.group(2).strip()

            # Normalize German responses to English
            if value.lower() == "ja":
                value = "Yes"
            elif value.lower() == "nein":
                value = "No"

            # Standardize variable names
            normalized_variable = standardize_variable_name(current_variable)
            if normalized_variable:
                relevant_points[normalized_variable] = value

        # Check for standalone answers in final answer section
        elif in_final_answer_section:
            match = standalone_answer_pattern.match(line)
            if match:
                variable = match.group(1).strip().lower()
                value = match.group(2).strip()

                # Normalize German responses
                if value.lower() == "ja":
                    value = "Yes"
                elif value.lower() == "nein":
                    value = "No"

                normalized_variable = standardize_variable_name(variable)
                if normalized_variable:
                    relevant_points[normalized_variable] = value

        # Capture additional explanations for previous bullet points
        elif current_variable and line:
            if current_variable in relevant_points:
                relevant_points[current_variable] += f" - {line}"

    return relevant_points

def standardize_variable_name(variable):
    """
    Normalize different spellings and variations for the same medical symptoms.
    
    Args:
        variable (str): Variable name to standardize
        
    Returns:
        str or None: Standardized variable name or None if not recognized
    """
    variable = variable.lower()
    
    # Map various symptom descriptions to standardized names
    if 'sudden' in variable or 'facial pain' in variable:
        return 'Trigeminal Pain'
    elif 'facial numbness' in variable or 'taub' in variable:
        return 'Facial Numbness'
    elif 'vertigo' in variable or 'dizz' in variable:
        return 'Vertigo'
    elif 'lacrimation' in variable or 'tear' in variable:
        return 'Lacrimation'
    elif 'spasm' in variable or 'muscle' in variable:
        return 'Facial Muscle Spasm'
    elif 'other' in variable:
        return 'Other'
    
    return None

def save_to_excel(data, output_file):
    """
    Save the extracted data to an Excel file, appending to existing data if file exists.
    
    Args:
        data (dict): Dictionary containing extracted data
        output_file (str): Path to output Excel file
    """
    # Define the standard column order
    columns = ['Patient_ID', 'Trigeminal Pain', 'Facial Numbness',
               'Vertigo', 'Lacrimation', 'Facial Muscle Spasm', 'Other', 
               'AI response', 'Parsed Data']

    # Load existing data or create new DataFrame
    try:
        df_existing = pd.read_excel(output_file)
    except FileNotFoundError:
        df_existing = pd.DataFrame(columns=columns)

    # Ensure all columns exist in data
    for col in columns:
        if col not in data:
            data[col] = None

    # Create new row and append to existing data
    df_new = pd.DataFrame([data], columns=columns)
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)

    # Save combined data back to Excel
    df_combined.to_excel(output_file, index=False)

def process_and_run_llm_for_subfolder(subfolder_path, api_key, model_name=model_name):
    """
    Process a single subfolder: extract text from documents, query AI model, and save results.
    
    Args:
        subfolder_path (str): Path to subfolder containing documents
        api_key (str): Google Gemini API key
        model_name (str): Name of the AI model to use
    """
    subfolder_name = os.path.basename(subfolder_path)
    
    # Extract and combine text from all Word documents in the folder
    combined_text = combine_word_documents(subfolder_path)
    context = extract_section(combined_text, sections)

    # Save context to file for debugging/review
    context_file_path = os.path.join(subfolder_path, "context.txt")
    with open(context_file_path, "w", encoding="utf-8") as file:
        file.write(context)

    # Format the query for the AI model with context
    formatted_query = (
        "You are a helpful physician assistant tasked with extracting clinical data for a study."
        "Use the following context as your learned knowledge, inside <context></context> XML tags.\n"
        f"<context>\n{context}\n</context>\n\n"
        "When answering the user:\n"
        "- If the context doesn't give you the information asked for, say so.\n"
        "Avoid mentioning that you obtained the information from the context.\n"
        "Always strictly stand by the information given in the context.\n\n"
        f"Given the context information, answer the query.\nQuery: {input_instructions}"
    )

    # Configure and run the Gemini model
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    try:
        # Generate response from AI model
        response = model.generate_content(formatted_query)
        full_response = response.text.strip()

    except Exception as e:
        print(f"Error processing '{subfolder_name}': {e}")
        return

    # Extract structured data from AI response
    relevant_points = extract_bullet_points(full_response)
    data = relevant_points
    data['Patient_ID'] = subfolder_name  # Use folder name as patient identifier
    data['AI response'] = full_response
    data['Parsed Data'] = context

    # Save results to Excel file
    save_to_excel(data, output_file)
    print(f"Response saved for {subfolder_name}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# -----------------------------
# Main execution with rate limiting
# -----------------------------

def main():
    """
    Main function to process all subfolders with rate limiting to respect API limits.
    """
    # Get list of subfolders to process
    subfolders = [f for f in os.listdir(main_folder) 
                  if os.path.isdir(os.path.join(main_folder, f))][:n_subfolders]

    # Process each iteration
    for i in range(1, iteration_n + 1):
        output_file = os.path.join(output_folder, f"output_{llm_nickname}_{i}.xlsx")
        start_time = time.time()

        print(f"Starting iteration {i}/{iteration_n}")
        
        # Process each subfolder
        for j, subfolder in enumerate(subfolders, start=1):
            subfolder_path = os.path.join(main_folder, subfolder)
            process_and_run_llm_for_subfolder(subfolder_path, API_KEY, model_name)

            # Rate limiting: pause after every `rate_limit` requests
            if j % rate_limit == 0:
                elapsed_time = time.time() - start_time
                remaining_time = 61 - elapsed_time  # Wait until a full minute has passed
                
                if remaining_time > 0:
                    print(f"Rate limit reached: Waiting {remaining_time:.2f} seconds before continuing...")
                    time.sleep(remaining_time)
                
                # Reset timer
                start_time = time.time()

        print(f"Iteration {i} completed. Output saved as {output_file}. "
              f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# Ensure output directory exists
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Run the main processing
if __name__ == "__main__":
    main()