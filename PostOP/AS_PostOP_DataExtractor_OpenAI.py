import os
import json
import requests
import datetime
import mammoth
import re
import pandas as pd
from openai import OpenAI

# Configuration - Replace with your actual API key
API_KEY = "your-openai-api-key-here"

# Configurable parameters - Update these paths for your environment
main_folder = "path/to/your/input/documents"
output_folder = "path/to/your/output/files"

# Model configuration
model_name = "o3-mini"
llm_nickname = "o3mini"

# Processing parameters
n_subfolders = 211
iteration_n = 3  # Number of iterations

# Section extraction configuration
sections = "all_sections"  # Set to list like ["Section1", "Section2"] or "all_sections"

# Instructions for the AI model
input_instructions = (
    "In the provided document, please look for the following postoperative complications: CSF leak, infection, facial palsy, facial numbness, and hearing loss. "
    "In your Answer, first reason whether any of your findings can really be considered a surgical complication. Focus on the fact that it is only a complication if it was not present before surgery and it is present afterwards"
    "Complications, if present, are always explicitly mentioned in the documents; If it is not mentioned, you must assume that the complication is not present. Please be mindful about the fact that the surgical access area is behind the ear, therefore numbness in that area should NOT be considered under facial numbness."
    "After reasoning about your findings, provide a final answer in the form of bullet points with 'Name of the Complication': 'Yes' or 'No' for each individual point. Feel free provide a short explanation after every statement"
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

def process_and_run_llm_for_subfolder(subfolder_path, api_key, model=model_name):
    """
    Process a single subfolder: combine documents, extract context, 
    send to AI model, and parse response.
    
    Args:
        subfolder_path (str): Path to the subfolder to process
        api_key (str): OpenAI API key
        model (str): Model name to use
        
    Returns:
        dict or None: Processed data dictionary or None if error occurred
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

    try:
        # Initialize OpenAI client and send request
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a helpful physician assistant tasked with extracting clinical data for a study. "
                        "You will be provided with a document. Your task is to look for the following postoperative complications: "
                        "CSF leak, infection, facial palsy, facial numbness, and hearing loss. In your Answer, first reason "
                        "whether any of your findings can really be considered a surgical complication. Focus on the fact "
                        "that it is only a complication if it was not present before surgery and it is present afterwards. "
                        "Complications, if present, are always explicitly mentioned in the documents; If it is not mentioned, "
                        "you must assume that the complication is not present. After reasoning about your findings, provide "
                        "a final answer in the form of bullet points with 'Name of the Complication': 'Yes' or 'No' for each "
                        "individual point. Feel free to provide a short explanation after every statement. Avoid mentioning "
                        "that you obtained the information from the context. Always strictly stand by the information given "
                        "in the context."
                    )
                },
                {
                    "role": "user",
                    "content": context
                }
            ]
        )

        full_response = response.choices[0].message.content

    except Exception as e:
        print(f"Error processing {subfolder_name}: {e}")
        return None

    # Parse the AI response for structured data
    relevant_points = extract_bullet_points(full_response)
    
    # Prepare data structure for output
    data = relevant_points.copy()
    data['AI response'] = full_response
    data['Parsed Data'] = context

    print(f"Response saved for {subfolder_name}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    return data

# Main execution: Run the LLM system for each subfolder multiple times
def main():
    """
    Main execution function: processes all subfolders through multiple iterations
    and saves results to Excel files.
    """
    # Get list of subfolders to process
    subfolders = [f for f in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, f))][:n_subfolders]

    # Run multiple iterations
    for i in range(1, iteration_n + 1):
        output_file = os.path.join(output_folder, f"output_{llm_nickname}_{i}.xlsx")
        all_data = []
        
        # Process each subfolder
        for subfolder in subfolders:
            subfolder_path = os.path.join(main_folder, subfolder)
            result = process_and_run_llm_for_subfolder(
                subfolder_path, 
                api_key=API_KEY, 
                model=model_name
            )
            if result:
                all_data.append(result)
        
        # Save results to Excel
        if all_data:
            df = pd.DataFrame(all_data)
            
            # Ensure output directory exists
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                
            df.to_excel(output_file, index=False)
            print(f"Iteration {i} completed. Output saved as {output_file}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# Execute main function if script is run directly
if __name__ == "__main__":
    main()