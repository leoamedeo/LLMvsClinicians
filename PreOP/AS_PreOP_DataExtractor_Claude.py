import os
import json
import requests
import datetime
import mammoth
import re
import pandas as pd
import time
import anthropic

# Replace with your Anthropic API key or set as environment variable
API_KEY = os.getenv('ANTHROPIC_API_KEY', 'your-anthropic-api-key-here')
model_name = "claude-3-5-sonnet-20241022"

# Configurable parameters
main_folder = "./input_documents"  # Path to folder containing patient document subfolders
output_file = "./output_files/output_claude.xlsx"  # Path to save output Excel file

n_subfolders = 211
sections = "all_sections"  # Name the required sections in square brackets as follows: ["Verlauf", "Befund", "Beurteilung"], otherwise set to "all_sections"
input_instructions = (
    "In the provided document please look for the following preoperative symptoms: Sudden Severe Facial Pain, Facial Numbness, Vertigo, Lacrimation, Facial Muscle Spasm, and Other (related to trigeminal neuralgia)."
    "In your Answer, first reason whether any of your findings can really be considered a preoperative symptom. Focus on the fact that it is only a preoperative symptom only if it was already present before the FIRST surgery. Always consider the first surgery if the patient underwent multiple ones. "
    "Consider the symptom only if it is explicitly mentioned in the documents; if it is not mentioned, always assume the symptom is not present."
    "After reasoning about your findings, provide a final answer in the form of bullet points with 'Name of the Symptom': 'Yes' or 'No' for each individual point."
)

def combine_word_documents(subfolder_path):
    """Combine the raw text of all .docx files in the specified folder into a single string."""
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
                        # Remove unnecessary newlines
                        extracted_text = re.sub(r'\n{2,}', '\n', extracted_text)
                        combined_text.append(extracted_text)
            except Exception as e:
                print(f"Error processing file {file_name}: {e}")

    return "\n".join(combined_text)  # Join all text with newlines

def extract_section(text, sections):
    """Extracts specified sections from the text, including the title and the paragraphs immediately after each matching section heading."""
    paragraphs = text.split("\n")
    section_content = []

    if sections == "all_sections":
        section_content = paragraphs
    else:
        for i, paragraph in enumerate(paragraphs):
            if any(section_heading.lower() in paragraph.lower() for section_heading in sections):
                section_content.append(paragraph)  # Add the section heading

                # Include the next paragraph if it exists and is not empty
                if i + 1 < len(paragraphs) and paragraphs[i + 1].strip():
                    section_content.append(paragraphs[i + 1])

    return "\n".join(section_content)

def extract_bullet_points(response_text):
    """Extract relevant bullet points from the response text, including 'Final answer:' formatted lists and embedded answers."""
    relevant_points = {}
    lines = response_text.split('\n')

    # Regex pattern to match:
    # - "•", "*", "-", "***", and numbered bullet points (1., 2., etc.)
    # - Also detects variables listed after "Final answer:"
    bullet_point_pattern = re.compile(
        r'^\s*(?:\*\s*|\*\*\*|\*+\s*\*\*?|[\u2022\-]|[1-5]\.)\s*\**\s*([^:]+?)\**\s*:\s*(?:\'|\")?(Yes|No|Ja|Nein)(?:\'|\")?.*',
        re.IGNORECASE
    )

    # Regex pattern to detect a plain text "Final answer" section, allowing variations like:
    # "Final answer in bullet points:", "Final decision:", etc.
    final_answer_pattern = re.compile(
        r'\bfinal answer\b', re.IGNORECASE  # Look for "Final answer" anywhere in a sentence
    )

    # Regex to match answers listed without bullets (e.g., "CSF leak: No")
    standalone_answer_pattern = re.compile(
        r'^\s*([A-Za-z\s]+):\s*(Yes|No|Ja|Nein)', re.IGNORECASE
    )

    current_variable = None
    in_final_answer_section = False  # Track if we are inside the "Final answer" block

    for line in lines:
        line = line.strip()

        # Check if the line indicates the start of a "Final answer" block
        if final_answer_pattern.search(line):
            in_final_answer_section = True
            continue  # Skip to the next line

        # Try to match the standard bullet point pattern
        match = bullet_point_pattern.match(line)
        if match:
            current_variable = match.group(1).strip().lower()
            value = match.group(2).strip()

            # Normalize "Ja" to "Yes" and "Nein" to "No"
            if value.lower() == "ja":
                value = "Yes"
            elif value.lower() == "nein":
                value = "No"

            # Standardize variable names
            normalized_variable = standardize_variable_name(current_variable)

            if normalized_variable:
                relevant_points[normalized_variable] = value

        # If in "Final answer" mode, check for plain-text answers
        elif in_final_answer_section:
            match = standalone_answer_pattern.match(line)
            if match:
                variable = match.group(1).strip().lower()
                value = match.group(2).strip()

                # Normalize "Ja" to "Yes" and "Nein" to "No"
                if value.lower() == "ja":
                    value = "Yes"
                elif value.lower() == "nein":
                    value = "No"

                # Standardize variable names
                normalized_variable = standardize_variable_name(variable)

                if normalized_variable:
                    relevant_points[normalized_variable] = value

        elif current_variable and line:  # Capture explanations for previous bullet points
            if current_variable in relevant_points:
                relevant_points[current_variable] += f" - {line}"

    return relevant_points

def standardize_variable_name(variable):
    """Normalize different spellings for the same preoperative symptoms."""
    variable = variable.lower()
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
    """Save the data to the Excel file."""
    # Define the column order for symptoms
    columns = ['Trigeminal Pain', 'Facial Numbness',
               'Vertigo', 'Lacrimation', 'Facial Muscle Spasm', 'Other', 'AI response', 'Parsed Data']

    # Load the existing data if the file exists, otherwise create an empty DataFrame
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

    # Save the combined data back to the Excel file
    df_combined.to_excel(output_file, index=False)

def process_and_run_llm_for_subfolder(subfolder_path, api_key, model_name="claude-3-5-sonnet-20241022"):
    """
    Process a subfolder, extract clinical data, and call Claude API for analysis.
    """
    subfolder_name = os.path.basename(subfolder_path)

    # Combine all Word documents in the subfolder
    combined_text = combine_word_documents(subfolder_path)
    context = extract_section(combined_text, sections)

    # Save extracted context to text file for reference
    context_file_path = os.path.join(subfolder_path, "context.txt")
    with open(context_file_path, "w", encoding="utf-8") as file:
        file.write(context)

    # Format the query for Claude
    formatted_query = (
        "You are a helpful physician assistant tasked with extracting clinical data for a study. "
        "Use the following context as your learned knowledge, inside <context></context> XML tags.\n"
        f"<context>\n{context}\n</context>\n\n"
        "When answering the user:\n"
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

    # Extract and parse the response
    relevant_points = extract_bullet_points(full_response)
    data = relevant_points
    data['AI response'] = full_response
    data['Parsed Data'] = context

    # Save to Excel file
    save_to_excel(data, output_file)
    print(f"Response saved for {subfolder_name}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# Main execution
# Get list of subfolders to process
subfolders = [f for f in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, f))][:n_subfolders]

# Start time tracking
start_time = time.time()

# Process each subfolder
for i, subfolder in enumerate(subfolders, start=1):
    process_and_run_llm_for_subfolder(
        os.path.join(main_folder, subfolder), 
        api_key=API_KEY
    )

# Print total processing time
end_time = time.time()
total_time = end_time - start_time
print(f"Total processing time: {total_time:.2f} seconds")