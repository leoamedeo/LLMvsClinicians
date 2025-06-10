# LLMvsClinicians

## Overview
This repository contains scripts and resources supporting the study: Large Language Models and Human Auditors in Extracting Clinical Information from Unstructured Medical Records: a Comparative Study by Aldo Spolaore et al. 
We evaluated multiple cloud‑based and local LLMs (GPT‑4o, o3‑mini, Sonnet 3.5, Gemini Flash, llama3.3, deepseek‑r1) against two human auditors on three extraction tasks:

1. Preoperative symptoms
2. Postoperative complications
3. Disease course

Results demonstrate that state‑of‑the‑art LLMs match or exceed human performance in accuracy, sensitivity, F1 score, and consistency.

## General Setup Instructions

Before running the scripts, please ensure the following setup steps are completed:

1. **Python Installation**: Make sure Python is installed on your system. The scripts are compatible with Python 3.10.10
2. **Dependency Installation**: Install the required Python packages. You can do this easily by using the `requirements.txt` file provided:
   ```bash
   pip install -r requirements.txt
   ```

## Data Preparation

Place your dataset files in accessible paths on your system.

## Configuration
1. API Keys:

- Set OPENAI_API_KEY for OpenAI models
- Set ANTHROPIC_API_KEY for Anthropic Claude
- Set GENAI_API_KEY for Google Gemini

2. Paths:

- In each script, update MAIN_FOLDER to point at your local folder of anonymized .docx patient records.
- Update OUTPUT_FOLDER or OUTPUT_FILE to your desired results directory.

Each script will iterate through subfolders, extract relevant text snippets, query the specified LLM zero‑shot, parse the "Symptom: Yes/No" bullet points, and save results to an Excel file.

## License

This project is released under the MIT License. See LICENSE for details.


## Contact
For questions or data/code requests, please contact:

Aldo Spolaore, MD <aldo.spolaore@med.uni-tuebingen.de>
