import os
import json
import requests
import datetime
import mammoth
import re
import pandas as pd

# Configurable parameters - Update these paths for your environment
main_folder = r"path/to/your/input/documents"
output_folder = r"path/to/your/output/files"

# Local LLM configuration
llm_url = "http://localhost:11434"  # Ollama default URL
llm_model = "deepseek-r1:70B"  # Replace with your preferred model
llm_nickname = "deepseek70B"  # Used for output file naming

# Model parameters
temperature = 0  # Controls randomness (0 = deterministic)
context_length = 32768  # Options: 32768 for 32k, 131072 for 128k context
n_subfolders = 211  # Number of subfolders to process
sections = "all_sections"  # Use "all_sections" or specify sections like ["Verlauf", "Befund", "Beurteilung"]

# Number of iterations to run (for consistency checking)
iteration_n = 3

# Instructions for AI analysis
input_instructions = (
    "In the provided document, please look for the following postoperative complications: CSF leak, infection, facial palsy, facial numbness, and hearing loss. "
    "In your Answer, first reason whether any of your findings can really be considered a surgical complication. Focus on the fact that it is only a complication if it was not present before surgery and it is present afterwards"
    "Complications, if present, are always explicitly mentioned in the documents; If it is not mentioned, you must assume that the complication is not present."
    "After reasoning about your findings, provide a final answer in the form of bullet points with 'Name of the Complication': 'Yes' or 'No' for each individual point. Feel free provide a short explanation after every statement"
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
        r'\s*(?:[\*\u2022\-]|[1-5]\.)\s*([^:]+):\s*(?:\*\*)?(Yes|No|Ja|Nein)(?:\*\*)?\s*(?:\([^)]*\))?',
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
        elif current_variable and line:
            # Add explanation to the existing bullet point's value
            if current_variable in relevant_points:
                relevant_points[current_variable] += f" - {line}"

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

def process_and_run_llm_for_subfolder(subfolder_path):
    """
    Process all Word documents in a subfolder and analyze them using local LLM.
    
    Args:
        subfolder_path (str): Path to the subfolder containing documents
        
    Returns:
        dict: Dictionary containing extracted data and AI response
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

    # Send request to local LLM (Ollama)
    try:
        response = requests.post(
            f"{llm_url}/api/generate",
            json={
                "prompt": formatted_query,
                "model": llm_model,
                "temperature": temperature,
                "num_ctx": context_length
            }
        )

        if response.status_code != 200:
            print(f"Error: Received status code {response.status_code}")
            return None

        # Parse streaming response from Ollama
        responses = []
        for line in response.iter_lines(decode_unicode=True):
            if line:
                try:
                    response_data = json.loads(line)
                    if "response" in response_data:
                        responses.append(response_data["response"])
                except ValueError:
                    continue

        full_response = "".join(responses)

    except requests.RequestException as e:
        print(f"Error connecting to LLM: {e}")
        return None

    # Extract structured data from AI response
    relevant_points = extract_bullet_points(full_response)
    
    # Prepare data for saving
    data = relevant_points.copy()
    data['Patient_ID'] = subfolder_name  # Use folder name as patient identifier
    data['AI response'] = full_response
    data['Parsed Data'] = context

    print(f"Response saved for {subfolder_name}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    return data

def main():
    """
    Main function to run multiple iterations of processing all subfolders.
    Creates separate output files for each iteration to enable consistency analysis.
    """
    # Get list of subfolders to process
    subfolders = [f for f in os.listdir(main_folder) 
                  if os.path.isdir(os.path.join(main_folder, f))][:n_subfolders]
    
    print(f"Processing {len(subfolders)} subfolders across {iteration_n} iterations...")

    # Run multiple iterations for consistency checking
    for i in range(1, iteration_n + 1):
        print(f"\nStarting iteration {i}/{iteration_n}...")
        
        # Create output file name with iteration number
        output_file = os.path.join(output_folder, f"output_{llm_nickname}_{i}.xlsx")
        all_data = []
        
        # Process each subfolder
        for subfolder in subfolders:
            subfolder_path = os.path.join(main_folder, subfolder)
            result = process_and_run_llm_for_subfolder(subfolder_path)
            if result:
                all_data.append(result)
        
        # Save results for this iteration
        if all_data:
            df = pd.DataFrame(all_data)
            
            # Create output folder if it doesn't exist
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                
            df.to_excel(output_file, index=False)
            print(f"Iteration {i} completed. Output saved as {output_file}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            print(f"No data collected for iteration {i}")

    print("\nAll iterations completed!")

if __name__ == "__main__":
    main()