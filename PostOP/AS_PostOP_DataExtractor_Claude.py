import os
import json
import requests
import datetime
import mammoth
import re
import pandas as pd
import time
import anthropic

# Configuration - Replace with your actual API key
API_KEY = "your-anthropic-api-key-here"
model_name = "claude-3-5-sonnet-20241022"

# Configurable parameters - Update these paths for your environment
main_folder = "path/to/your/input/documents"
output_file = "path/to/your/output/file.xlsx"

# Processing parameters
n_subfolders = 212

# Section extraction configuration
sections = "all_sections"  # Set to list like ["Section1", "Section2"] or "all_sections"

# Instructions for the AI model
input_instructions = (
    "In the provided document, please look for the following postoperative complications: CSF leak, infection, facial palsy, facial numbness, and hearing loss. "
    "In your Answer, first reason whether any of your findings can really be considered a surgical complication. Focus on the fact that it is only a complication if it was not present before surgery and it is present afterwards"
    "Complications, if present, are always explicitly mentioned in the documents; If it is not mentioned, you must assume that the complication is not present. Please be mindful about the fact that the surgical access area is behind the ear, therefore numbness in that area should NOT be considered under facial numbness."
    "After reasoning about your findings, provide a final answer in the form of bullet points with 'Name of the Complication': 'Yes' or 'No' for each individual point. Do not use bold text."
)

def combine_word_documents(subfolder_path):
    """
    Combine the raw text of all .docx files in the specified folder into a single string.
    Extracts text between 'diagnos' and 'grüße' patterns and cleans formatting.
    
    Args:
        subfolder_path (str): Path to the folder containing .docx files
        
    Returns:
        str: Combined text from all documents
    """
    combined_text = []

    for file_name in os.listdir(subfolder_path):
        if file_name.endswith(".docx"):
            file_path = os.path.join(subfolder_path, file_name)
            try:
                # Extract raw text from Word document
                with open(file_path, "rb") as docx_file:
                    result = mammoth.extract_raw_text(docx_file)
                    text = result.value

                    # Extract text between "diagnos" and "grüße" (case-insensitive), including delimiters
                    match = re.search(r'(diagnos.*?grüße)', text, re.IGNORECASE | re.DOTALL)
                    if match:
                        extracted_text = match.group(1).strip()

                        # Remove unnecessary newlines (replace multiple newlines with single)
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
        sections (str or list): Either "all_sections" or list of section names to extract
        
    Returns:
        str: Extracted section content
    """
    paragraphs = text.split("\n")
    section_content = []

    if sections == "all_sections":
        # Return all content if no specific sections requested
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

def extract_bullet_points(response_text):
    """
    Extract relevant bullet points from the AI response text, including 'Final answer:' 
    formatted lists and embedded answers. Handles multiple bullet point formats and languages.
    
    Args:
        response_text (str): AI response text to parse
        
    Returns:
        dict: Dictionary with standardized complication names as keys and Yes/No as values
    """
    relevant_points = {}
    lines = response_text.split('\n')

    # Regex pattern to match various bullet point formats
    bullet_point_pattern = re.compile(
        r'^\s*(?:\*\s*|\*\*\*|\*+\s*\*\*?|[\u2022\-]|[1-5]\.)\s*\**\s*([^:]+?)\**\s*:\s*(?:\'|\")?(Yes|No|Ja|Nein)(?:\'|\")?.*',
        re.IGNORECASE
    )

    # Pattern to detect "Final answer" section
    final_answer_pattern = re.compile(r'\bfinal answer\b', re.IGNORECASE)

    # Pattern to match answers without bullets (e.g., "CSF leak: No")
    standalone_answer_pattern = re.compile(r'^\s*([A-Za-z\s]+):\s*(Yes|No|Ja|Nein)', re.IGNORECASE)

    current_variable = None
    in_final_answer_section = False

    for line in lines:
        line = line.strip()

        # Check if we're entering the "Final answer" section
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

        # Handle plain-text answers in final answer section
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

        # Capture explanations for previous bullet points
        elif current_variable and line:
            if current_variable in relevant_points:
                relevant_points[current_variable] += f" - {line}"

    return relevant_points

def standardize_variable_name(variable):
    """
    Normalize different spellings and languages for the same medical complications.
    
    Args:
        variable (str): Variable name to standardize
        
    Returns:
        str or None: Standardized variable name or None if not recognized
    """
    variable = variable.lower()
    
    # Map various terms to standardized complication names
    if 'leak' in variable or 'liquor' in variable:
        return 'CSF Leak'
    elif 'infection' in variable or 'infektion' in variable:
        return 'Infection'
    elif 'facial palsy' in variable or 'gesichtslähmung' in variable:
        return 'Facial Palsy'
    elif 'facial numbness' in variable or 'taub' in variable:
        return 'Facial Numbness'
    elif 'hearing loss' in variable or 'hörverlust' in variable:
        return 'Hearing Loss'
    
    return None

def save_to_excel(data, output_file):
    """
    Save the data to the Excel file, appending to existing data if file exists.
    
    Args:
        data (dict): Data dictionary to save
        output_file (str): Path to output Excel file
    """
    # Define expected columns for the Excel output
    columns = ['CSF Leak', 'Infection', 'Facial Palsy', 'Facial Numbness', 
               'Hearing Loss', 'AI response', 'Parsed Data']

    # Load existing data if file exists, otherwise create empty DataFrame
    try:
        df_existing = pd.read_excel(output_file)
    except FileNotFoundError:
        df_existing = pd.DataFrame(columns=columns)

    # Ensure the data has all necessary columns
    for col in columns:
        if col not in data:
            data[col] = None

    # Convert the data dictionary to a DataFrame
    df_new = pd.DataFrame([data], columns=columns)

    # Append the new data to the existing data
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined.to_excel(output_file, index=False)

def process_and_run_llm_for_subfolder(subfolder_path, api_key, model_name="claude-3-5-sonnet-20241022"):
    """
    Process a single subfolder: combine documents, extract context, 
    send to Claude AI model, and save response.
    
    Args:
        subfolder_path (str): Path to the subfolder to process
        api_key (str): Anthropic API key
        model_name (str): Claude model name to use
    """
    subfolder_name = os.path.basename(subfolder_path)

    # Combine all Word documents in the folder
    combined_text = combine_word_documents(subfolder_path)
    
    # Extract relevant sections
    context = extract_section(combined_text, sections)

    # Save context for debugging/reference
    context_file_path = os.path.join(subfolder_path, "context.txt")
    with open(context_file_path, "w", encoding="utf-8") as file:
        file.write(context)

    # Format the query with context for Claude
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

    # Initialize the Anthropic client
    client = anthropic.Anthropic(api_key=api_key)

    try:
        # Generate response from Claude
        response = client.messages.create(
            model=model_name,
            max_tokens=1000,
            temperature=0,
            messages=[
                {"role": "user", "content": formatted_query}
            ]
        )
        
        # Extract text response
        full_response = response.content[0].text.strip()

    except Exception as e:
        print(f"Error processing '{subfolder_name}': {e}")
        return

    # Parse the AI response for structured data
    relevant_points = extract_bullet_points(full_response)
    
    # Prepare data structure for output
    data = relevant_points.copy()
    data['AI response'] = full_response
    data['Parsed Data'] = context

    # Save data to Excel file
    save_to_excel(data, output_file)
    print(f"Response saved for {subfolder_name}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

def main():
    """
    Main execution function: processes all subfolders and tracks execution time.
    """
    # Get list of subfolders to process
    subfolders = [f for f in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, f))][:n_subfolders]

    # Start time tracking
    start_time = time.time()

    # Process each subfolder
    for i, subfolder in enumerate(subfolders, start=1):
        print(f"Processing subfolder {i}/{len(subfolders)}: {subfolder}")
        process_and_run_llm_for_subfolder(
            os.path.join(main_folder, subfolder), 
            api_key=API_KEY,
            model_name=model_name
        )

    # Calculate and display total execution time
    end_time = time.time()
    total_time = end_time - start_time
    print(f"Total processing time: {total_time:.2f} seconds ({total_time/60:.2f} minutes)")

# Execute main function if script is run directly
if __name__ == "__main__":
    main()