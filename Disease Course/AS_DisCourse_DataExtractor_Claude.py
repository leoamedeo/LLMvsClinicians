import os
import datetime
import mammoth
import re
import pandas as pd
import time
import anthropic

# Configuration - Replace with your actual API key
API_KEY = "your-anthropic-api-key-here"

# LLM Configuration
model_name = "claude-3-5-sonnet-20241022"  #choose the desired model here
llm_nickname = "claude"

# Configurable parameters
main_folder = "path/to/your/input/documents"  # Path to folder containing patient document subfolders
output_folder = "path/to/your/output/files"  # Path where Excel output files will be saved

# Processing parameters
n_subfolders = 211  # Number of subfolders to process
sections = "all_sections"  # Specify sections to extract: ["Section1", "Section2"] or "all_sections"
iteration_n = 3  # Number of iterations to run the analysis

# Instructions for the LLM to analyze medical documents
input_instructions = (
    "In the provided document, please analyze the patient's disease course based on the text within <context> and determine the correct values for each of the following requested data point. After providing a summary of the patients disease course, provide a final answer in the form of bullet points following the same structure of the datapoints below."
    "- Any improvement of pain after first surgery (Yes/No)"
    "- Completely free of pain after first surgery (Yes/No)"
    "- Symptom recurrence after first surgery (Yes/No, if it is not explicitly mentioned, assume there was no recurrence)"
    "- A second surgery was carried out (Yes/No, if it is not explicitly mentioned, assume there was no second surgery)"
    "- Free of pain after second surgery (Yes/No/Not provided)"
    "- Recurrence after second surgery (Yes/No/Not provided)"
    "- Thermocoagulation was carried out (Yes/No/Not provided)"
)

def combine_word_documents(subfolder_path):
    """
    Combine the raw text of all .docx files in the specified folder into a single string.
    Assumes documents have already been anonymized.
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
                    # This extracts the relevant medical content from the document
                    match = re.search(r'(diagnos.*?grüße)', text, re.IGNORECASE | re.DOTALL)
                    if match:
                        extracted_text = match.group(1).strip()

                        # Clean up excessive newlines for better readability
                        extracted_text = re.sub(r'\n{2,}', '\n', extracted_text)

                        combined_text.append(extracted_text)
            except Exception as e:
                print(f"Error processing file {file_name}: {e}")

    return "\n".join(combined_text)  # Join all text with newlines

def extract_section(text, sections):
    """
    Extract specified sections from the text, including the title and paragraphs 
    immediately after each matching section heading.
    
    Args:
        text: The full document text
        sections: List of section names to extract, or "all_sections" for complete text
    
    Returns:
        String containing the extracted sections
    """
    paragraphs = text.split("\n")
    section_content = []

    if sections == "all_sections":
        section_content = paragraphs
    else:
        for i, paragraph in enumerate(paragraphs):
            # Check if any of the specified sections appear in this paragraph
            if any(section_heading.lower() in paragraph.lower() for section_heading in sections):
                section_content.append(paragraph)  # Add the section heading

                # Include the next paragraph if it exists and is not empty
                if i + 1 < len(paragraphs) and paragraphs[i + 1].strip():
                    section_content.append(paragraphs[i + 1])

    return "\n".join(section_content)

def standardize_variable_name(variable):
    """
    Normalize different spellings and phrasings for the same data points to ensure
    consistent categorization of extracted information.
    
    Args:
        variable: The variable name to standardize
        
    Returns:
        Standardized variable name or None if not recognized
    """
    variable = variable.lower().strip()
    
    # Check explicit matches first for precise categorization
    if variable == 'free of pain after second surgery':
        return 'Free of pain after second surgery'
    elif variable == 'recurrence after second surgery':
        return 'Recurrence after second surgery'
    
    # Pattern matching for various phrasings of the same concept
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
    Extract relevant bullet points from the LLM response text, parsing both
    'Final answer:' formatted lists and embedded answers throughout the response.
    
    Args:
        response_text: The full response from the LLM
        
    Returns:
        Dictionary mapping standardized variable names to their values
    """
    relevant_points = {}
    lines = response_text.split('\n')

    # Regex pattern to match various bullet point formats
    bullet_point_pattern = re.compile(
        r'^\s*(?:\*\s*|\*\*\*|\*+\s*\*\*?|[\u2022\-]|[1-5]\.)\s*\**\s*([^:]+?)\**\s*:\s*(?:\'|\")?(Yes|No|Ja|Nein|provided|know)(?:\'|\")?.*',
        re.IGNORECASE
    )

    # Pattern to identify "final answer" sections
    final_answer_pattern = re.compile(
        r'\bfinal answer\b', re.IGNORECASE
    )

    # Pattern for standalone answers without bullet formatting
    standalone_answer_pattern = re.compile(
        r'^\s*([A-Za-z\s]+):\s*(Yes|No|Ja|Nein|provided|know)', re.IGNORECASE
    )

    in_final_answer_section = False

    for line in lines:
        line = line.strip()

        # Check if we've reached the final answer section
        if final_answer_pattern.search(line):
            in_final_answer_section = True
            continue

        # Try to match bullet point format
        match = bullet_point_pattern.match(line)
        if match:
            current_variable = match.group(1).strip().lower()
            value = match.group(2).strip()

            # Normalize German responses and handle "Not provided" cases
            if value.lower() == "ja":
                value = "Yes"
            elif value.lower() == "nein":
                value = "No"
            elif value.lower() == "provided":
                value = "Zero"
            elif value.lower() == "know":
                value = "Zero"

            normalized_variable = standardize_variable_name(current_variable)

            if normalized_variable:
                relevant_points[normalized_variable] = value

        # Handle standalone answers in final answer section
        elif in_final_answer_section:
            match = standalone_answer_pattern.match(line)
            if match:
                variable = match.group(1).strip().lower()
                value = match.group(2).strip()

                # Apply same normalization as above
                if value.lower() == "ja":
                    value = "Yes"
                elif value.lower() == "nein":
                    value = "No"
                elif value.lower() == "provided":
                    value = "Zero"
                elif value.lower() == "know":
                    value = "Zero"

                normalized_variable = standardize_variable_name(variable)

                if normalized_variable:
                    relevant_points[normalized_variable] = value

    return relevant_points

def process_and_run_llm_for_subfolder(subfolder_path, api_key, model_name="claude-3-5-sonnet-20241022"):
    """
    Process a single subfolder containing medical documents:
    1. Combine all Word documents in the folder
    2. Extract relevant sections
    3. Send to Claude for analysis
    4. Parse the response and extract structured data
    
    Args:
        subfolder_path: Path to the subfolder containing documents
        api_key: Anthropic API key
        model_name: Claude model to use
        
    Returns:
        Dictionary containing extracted data points
    """
    subfolder_name = os.path.basename(subfolder_path)

    # Combine all Word documents in the subfolder
    combined_text = combine_word_documents(subfolder_path)
    
    # Extract relevant sections from the combined text
    context = extract_section(combined_text, sections)

    # Save the context for debugging/review purposes
    context_file_path = os.path.join(subfolder_path, "context.txt")
    with open(context_file_path, "w", encoding="utf-8") as file:
        file.write(context)

    # Format the query for Claude with proper context structure
    formatted_query = (
        "You are a capable physician assistant. You are given medical documents of patients who underwent microvascular decompression surgery. Your role is to analyze the patient's disease course carefully and then provide specific data points as your final answer. "
        "Use the following context as your learned knowledge, inside <context></context> XML tags.\n"
        f"<context>\n{context}\n</context>\n\n"
        "When answering the user:\n"
        "- If you don't know, just say that you don't know.\n"
        "- If the context doesn't give you the information asked for, say so.\n"
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
        
        # Extract text response from Claude's response format
        full_response = response.content[0].text.strip()

    except Exception as e:
        print(f"Error processing '{subfolder_name}': {e}")
        return None

    # Extract structured data from the LLM response
    relevant_points = extract_bullet_points(full_response)
    
    # Prepare the final data structure
    data = relevant_points.copy()
    data['Subfolder'] = subfolder_name  # Using generic identifier instead of personal info
    data['AI response'] = full_response
    data['Parsed Data'] = context

    print(f"Response saved for {subfolder_name}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    return data

def main():
    """
    Main execution function that processes all subfolders through multiple iterations
    and saves results to Excel files.
    """
    # Get list of subfolders to process
    subfolders = [f for f in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, f))][:n_subfolders]

    # Run multiple iterations for consistency checking
    for i in range(1, iteration_n + 1):
        output_file = os.path.join(output_folder, f"output_{llm_nickname}_{i}.xlsx")
        all_data = []
        
        # Process each subfolder
        for subfolder in subfolders:
            subfolder_path = os.path.join(main_folder, subfolder)
            result = process_and_run_llm_for_subfolder(
                subfolder_path, 
                api_key=API_KEY,
                model_name=model_name
            )
            if result:
                all_data.append(result)
        
        # Save results to Excel file
        if all_data:
            df = pd.DataFrame(all_data)
            
            # Create output directory if it doesn't exist
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                
            df.to_excel(output_file, index=False)
            print(f"Iteration {i} completed. Output saved as {output_file}. {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    main()