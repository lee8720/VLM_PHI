import openai
import pandas as pd
import os
import time
import json
import re
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
from PIL import Image
import requests
from io import BytesIO


# Load the Excel file that contains the file names and URLs
file_path = "####.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)
file_names = df['FileName'].tolist()
urls = df['URL'].tolist()

# Google Gemini API Key
genai.configure(api_key="####")
model = genai.GenerativeModel("gemini-1.5-pro")
generation_config = genai.GenerationConfig(temperature=0.0, max_output_tokens=1000)

# message prompts
messages_list = [
    # prompt 1
    """
    The file contains an image that requires analysis. Please review the image to confirm whether the text-based personal information related to the patient has been fully de-identified. 
    It is possible that the image does not contain any text.

    After completing the analysis, create a table with two columns: 
    1. The first column should indicate whether the personal information has been de-identified ('Yes' or 'No'). 
    2. The second column should provide a brief explanation for your determination.

    Provide the result as a JSON object structured as follows:
    {
      "de-identified": "Yes/No",
      "explanation": "Brief explanation for the determination"
    }
    """,

    # Prompt 2
    """
    The file contains an image that requires analysis. Please review the image to identify any text-based personal information related to the patient. 
    It is possible that the image does not contain any text.

    After completing the analysis, create a table with three columns: 
    1. The first column should list the categories of personal information detected in the image (with no limitation on the number or type of categories). 
    2. The second column should indicate whether each category is present in the image ('Yes' or 'No'). 
    3. The third column should provide details for any category where personal information is detected.

    Based solely on this analysis, confirm whether the text-based personal information has been fully de-identified ('Yes' or 'No') and provide a brief explanation for your determination.

    Provide the result as a JSON object structured as follows:
    {
      "categories": [
        {"category": "Detected Category 1", "present": "Yes/No", "details": "Detected details if present"},
        {"category": "Detected Category 2", "present": "Yes/No", "details": "Detected details if present"},
        ...
      ],
      "de-identified": "Yes/No"
      "explanation": "Brief explanation for the determination"
    }
    """,

    # Prompt 3
    """
    The file contains an image that requires analysis. Please review the image to identify any text-based personal information related to the patient. Personal information is defined by the following five categories: 1. Patient Name; 2. Patient ID; 3. Patient Date of Birth; 4. Patient Age; 5. Patient Gender. 
    It is possible that the image does not contain any text.

    After completing the analysis, create a table with three columns: 
    1. The first column should list the five categories mentioned above. 
    2. The second column should indicate whether each category is present in the image ('Yes' or 'No'). 
    3. The third column should provide details for any category where personal information is detected.

    Based solely on this analysis, confirm whether the text-based personal information has been fully de-identified ('Yes' or 'No') and provide a brief explanation for your determination.

    Provide the result as a JSON object structured as follows:
    {
      "categories": [
        {"category": "Patient Name", "present": "Yes/No", "details": "Detected details if present"},
        {"category": "Patient ID", "present": "Yes/No", "details": "Detected details if present"},
        {"category": "Patient Date of Birth", "present": "Yes/No", "details": "Detected details if present"},
        {"category": "Patient Age", "present": "Yes/No", "details": "Detected details if present"},
        {"category": "Patient Gender", "present": "Yes/No", "details": "Detected details if present"}
      ],
      "de-identified": "Yes/No"
      "explanation": "Brief explanation for the determination"

    }
    """
]

# Output file path
output_file_path = "####.xlsx"

# Check if the file already exists
if os.path.exists(output_file_path):
    output_df = pd.read_excel(output_file_path)
else:
    output_df = pd.DataFrame(columns=[
        'File Name', 'Prompt Number', 'De-Identified', 'Explanation',
        'Category 1', 'Present 1', 'Details 1',
        'Category 2', 'Present 2', 'Details 2',
        'Category 3', 'Present 3', 'Details 3',
        'Category 4', 'Present 4', 'Details 4',
        'Category 5', 'Present 5', 'Details 5'
    ])

# Set the number of retry attempts
max_retries = 5
num_repetition = 3

# Loop over file names and URLs
for r in range(num_repetition):
    for file_name, url in zip(file_names, urls):
        for idx, message in enumerate(messages_list):
            retries = 0
            success = False

            # Retry loop
            while retries < max_retries and not success:
                try:
                    response = requests.get(url)
                    image = Image.open(BytesIO(response.content))
                    response = model.generate_content([image, message], generation_config=generation_config,
                                              safety_settings={HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
                                                               HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
                                                               HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
                                                               HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE
                                                               }
                                                      )
                    # Validate if the response is empty
                    # Validate if the response is empty
                    if not response.text:
                        raise ValueError("Received an empty response from the API.")

                    # Parse the JSON response content
                    response_json = response.text

                    # Remove comments or non-JSON parts using regular expressions
                    # This will match and extract the JSON part within the response
                    json_match = re.search(r"\{.*\}", response_json, re.DOTALL)

                    if json_match:
                        # Extract the JSON part
                        response_json = json_match.group(0)
                        response_data = json.loads(response_json)  # Parsing the string as JSON
                    else:
                        raise ValueError("No valid JSON object found in the response.")

                    # Initialize the new row with basic info
                    new_row = {
                        'File Name': file_name,
                        'Prompt Number': idx + 1,
                        'De-Identified': response_data.get("de-identified", "N/A"),
                        'Explanation': response_data.get("explanation", "N/A")
                    }

                    # Handle category details for Prompt 2 (no limit on category count)
                    if ((idx == 1) or (idx == 2))and "categories" in response_data:
                        for i, category_info in enumerate(response_data["categories"]):
                            new_row[f'Category {i + 1}'] = category_info['category']
                            new_row[f'Present {i + 1}'] = category_info['present']
                            new_row[f'Details {i + 1}'] = category_info['details']

                    # Add the row to the DataFrame
                    output_df = pd.concat([output_df, pd.DataFrame([new_row])], ignore_index=True)
                except Exception as e:
                    retries += 1
                    print(f"Error processing {file_name} for prompt {idx + 1}: {str(e)}")
                    print(f"Retrying {retries}/{max_retries}...")

                    time.sleep(2)

            if not success:
                error_row = pd.DataFrame({'File Name': [file_name],'Prompt Number': [idx + 1],
                                          'De-Identified': ["ERROR"], 'Explanation': [response_json]})
                output_df = pd.concat([output_df, error_row], ignore_index=True)
    # Save the updated DataFrame to Excel after each iteration
output_df.to_excel(output_file_path, index=False)

print("The results have been saved to the Excel file.")
