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
API_KEY = "YOUR_GEMINI_API_KEY_HERE"

# Model configuration
model_name = "gemini-2.0-flash-exp"  # Options: gemini-1.5-flash, gemini-2.0-flash-exp

# NOTE: Adjust requests per minute to 15 for flash1.5 and to 10 for flash2.0
rate_limit = 10

# Configurable file paths - Update these paths for your environment
main_folder = r"path/to/your/input/documents"
output_file = r"path/to/your/output/results.xlsx"

# Processing parameters
n_subfolders = 212
sections = "all_sections"  # Use "all_sections" or specify sections like ["Verlauf", "Befund", "Beurteilung"]

# Instructions for AI analysis
input_instructions = (
    "In the provided document, please look for the following postoperative complications: CSF leak, infection, facial palsy, facial numbness, and hearing loss. "
    "In your Answer, first reason whether any of your findings can really be considered a surgical complication. Focus on the fact that it is only a complication if it was not present before surgery and it is present afterwards"
    "Complications, if present, are always explicitly mentioned in the documents; If it is not mentioned, you must assume that the complication is not present. Please be mindful about the fact that the surgical access area is behind the ear, therefore numbness in that area should NOT be considered under facial numbness."
    "After reasoning about your findings, provide a final answer in the form of bullet points with 'Name of the Complication': 'Yes' or 'No' for each individual point. Do not use bold text."
)

def combine_word_documents(subfolder_path):
    """
    Combine the raw text of all .docx files in the specified folder into a single string.
    Extracts text between 'diagnos' and 'grüße' markers and removes excessive newlines.
    
    Args:
        subfolder_path (str): Path to the folder containing .docx files
        
    Returns:
        str: Combined text from all documents in the folder
    """
    combined_text = []

    for file_name in os.listdir(subfolder_path):
        if file_name.endswith(".docx"):
            file_path = os.path.join(subfolder_path, file_name)
            try:
                with open(file_path, "rb") as docx_file:
                    result = mammoth.extract_raw_text(docx_file)
                    text = result.value

                    # Extract text between "diagnos" and "grüße" (case-insensitive), including delimiters
                    match = re.search(r'(diagnos.*?grüße)', text, re.IGNORECASE | re.DOTALL)
                    if match:
                        extracted_text = match.group(1).strip()

                        # Remove excessive newlines (2 or more consecutive newlines become single newline)
                        extracted_text = re.sub(r'\n{2,}', '\n', extracted_text)

                        combined_text.append(extracted_text)
            except Exception as e:
                print(f"Error processing file {file_name}: {e}")

    return "\n".join(combined_text)

def extract_section(text, sections):
    """
    Extract specified sections from the text, including the title and paragraphs after each matching section heading.
    
    Args:
        text (str): Input text to extract sections from
        sections: Either "all_sections" to return all text, or list of section headings to extract
        
    Returns:
        str: Extracted section content
    """
    paragraphs = text.split("\n")
    section_content = []

    if sections == "all_sections":
        section_content = paragraphs
    else:
        for i, paragraph in enumerate(paragraphs):
            # Check if any of the specified section headings are in the current paragraph
            if any(section_heading.lower() in paragraph.lower() for section_heading in sections):
                section_content.append(paragraph)  # Add the section heading

                # Include the next paragraph if it exists and is not empty
                if i + 1 < len(paragraphs) and paragraphs[i + 1].strip():
                    section_content.append(paragraphs[i + 1])

    return "\n".join(section_content)

def extract_bullet_points(response_text):
    """
    Extract relevant complications data from AI response text.
    Handles various bullet formats, bold text, and normalizes responses.
    
    Args:
        response_text (str): Raw AI response text
        
    Returns:
        dict: Dictionary with complication names as keys and Yes/No as values
    """
    relevant_points = {}
    lines = response_text.split('\n')

    # Pattern to match bullet points with complications and Yes/No answers
    bullet_point_pattern = re.compile(
        r'^\s*(?:\*\s*|\*\*\*|\*+\s*\*\*?|[\u2022\-]|[1-5]\.)\s*\**\s*([^:]+?)\**\s*:\s*(?:\'|\")?(Yes|No|Ja|Nein)(?:\'|\")?.*',
        re.IGNORECASE
    )

    current_variable = None

    for line in lines:
        line = line.strip()

        # Try to match the bullet point pattern
        match = bullet_point_pattern.match(line)
        if match:
            current_variable = match.group(1).strip().lower()
            value = match.group(2).strip()

            # Normalize German responses to English
            if value.lower() == "ja":
                value = "Yes"
            elif value.lower() == "nein":
                value = "No"

            # Normalize variable names to standard complication categories
            normalized_variable = None
            if 'leak' in current_variable or 'liquor' in current_variable:
                normalized_variable = 'CSF Leak'
            elif 'infection' in current_variable or 'infektion' in current_variable:
                normalized_variable = 'Infection'
            elif 'facial palsy' in current_variable or 'gesichtslähmung' in current_variable:
                normalized_variable = 'Facial Palsy'
            elif 'facial numbness' in current_variable or 'taub' in current_variable:
                normalized_variable = 'Facial Numbness'
            elif 'hearing loss' in current_variable or 'hörverlust' in current_variable:
                normalized_variable = 'Hearing Loss'

            if normalized_variable:
                relevant_points[normalized_variable] = value

    return relevant_points

def save_to_excel(data, output_file):
    """
    Save the extracted data to an Excel file, appending to existing data if file exists.
    
    Args:
        data (dict): Dictionary containing the data to save
        output_file (str): Path to the output Excel file
    """
    columns = ['Patient_ID', 'CSF Leak', 'Infection', 'Facial Palsy',
               'Facial Numbness', 'Hearing Loss', 'AI response', 'Parsed Data']

    # Load existing data if file exists, otherwise create empty DataFrame
    try:
        df_existing = pd.read_excel(output_file)
    except FileNotFoundError:
        df_existing = pd.DataFrame(columns=columns)

    # Ensure all required columns are present in data
    for col in columns:
        if col not in data:
            data[col] = None

    # Convert data to DataFrame and append to existing data
    df_new = pd.DataFrame([data], columns=columns)
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined.to_excel(output_file, index=False)

def process_and_run_llm_for_subfolder(subfolder_path, api_key, model_name=model_name):
    """
    Process all Word documents in a subfolder and analyze them using Gemini AI.
    
    Args:
        subfolder_path (str): Path to the subfolder containing documents
        api_key (str): Google AI API key
        model_name (str): Name of the Gemini model to use
    """
    subfolder_name = os.path.basename(subfolder_path)
    
    # Combine all Word documents in the folder
    combined_text = combine_word_documents(subfolder_path)
    context = extract_section(combined_text, sections)

    # Save context to file for reference
    context_file_path = os.path.join(subfolder_path, "context.txt")
    with open(context_file_path, "w", encoding="utf-8") as file:
        file.write(context)

    # Format the query for the AI model
    formatted_query = (
        "You are a helpful physician assistant tasked with extracting clinical data for a study. "
        "Use the following context as your learned knowledge, inside <context></context> XML tags.\n"
        f"<context>\n{context}\n</context>\n\n"
        "When answering the user:\n"
        "- If you don't know, just say that you don't know.\n"
        "- If the context doesn't give you the information asked for, say so.\n"
        "Avoid mentioning that you obtained the information from the context.\n"
        "Always strictly stand by the information given in the context.\n\n"
        f"Given the context information, answer the query.\nQuery: {input_instructions}"
    )

    # Configure and initialize the Gemini model
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    try:
        # Generate response from Gemini AI
        response = model.generate_content(formatted_query)
        full_response = response.text.strip()

    except Exception as e:
        print(f"Error processing '{subfolder_name}': {e}")
        return

    # Extract structured data from AI response
    relevant_points = extract_bullet_points(full_response)
    
    # Prepare data for saving
    data = relevant_points.copy()
    data['Patient_ID'] = subfolder_name  # Use folder name as patient identifier
    data['AI response'] = full_response
    data['Parsed Data'] = context

    # Save results to Excel file
    save_to_excel(data, output_file)
    print(f"Response saved for {subfolder_name}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


# -----------------------------
# Main processing loop with rate limiting
# -----------------------------

def main():
    """Main function to process all subfolders with rate limiting."""
    
    # Get list of subfolders to process
    subfolders = [f for f in os.listdir(main_folder) 
                  if os.path.isdir(os.path.join(main_folder, f))][:n_subfolders]

    print(f"Processing {len(subfolders)} subfolders...")
    
    # Start timing for rate limiting
    start_time = time.time()

    for i, subfolder in enumerate(subfolders, start=1):
        process_and_run_llm_for_subfolder(
            os.path.join(main_folder, subfolder), 
            api_key=API_KEY
        )

        # Rate limiting: pause after every 'rate_limit' requests
        if i % rate_limit == 0:
            elapsed_time = time.time() - start_time
            remaining_time = 61 - elapsed_time  # Wait for full minute to pass

            if remaining_time > 0:
                print(f"Rate limit reached: Waiting {remaining_time:.2f} seconds before continuing...")
                time.sleep(remaining_time)

            # Reset timer after waiting
            start_time = time.time()

    print("Processing completed!")

if __name__ == "__main__":
    main()