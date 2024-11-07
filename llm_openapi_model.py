import os
import pandas as pd
import time
from openai import OpenAI

# Set up the OpenAI API key directly or through an environment variable
os.environ["OPENAI_API_KEY"] = "key_here"

# Initialize the OpenAI client
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

# Load the spreadsheet
file_path = 'Color Normalization.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load the relevant sheets
normalized_colors_df = excel_data.parse('Normalized List of Colors')
retailer_colors_df = excel_data.parse('RetailerColors')
retailer_garment_colors_df = excel_data.parse('RetailerGarmentColors')

# Generate a list of unique color terms to reduce API calls
unique_terms = pd.concat([retailer_colors_df['retailerColors'], retailer_garment_colors_df['retailerGarmentColors']]).unique()

# Cache to store the results from OpenAI API calls
term_mappings_cache = {}

# Function to query OpenAI API for color mapping
def query_openai_for_color(term):
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",  # Use "gpt-4" if you have access to it
            messages=[
                {"role": "system", "content": "You are an assistant that helps identify colors in product descriptions."},
                {"role": "user", "content": f"Identify the main color for '{term}'. Return '[blank]' if it is not a color-related term, or 'check URL' if it implies multiple colors."}
            ],
            max_tokens=10,
            temperature=0,
            timeout=10  # Timeout in seconds
        )
        
        # Access the content of the message directly
        message_content = response.choices[0].message.content.strip()
        return message_content

    except Exception as e:
        print(f"Error with term '{term}': {e}")
        return "check URL"  # Fallback in case of an API error

# Query the API for each unique term with progress tracking
for i, term in enumerate(unique_terms, 1):
    print(f"Processing term {i}/{len(unique_terms)}: '{term}'")  # Progress tracking
    if term not in term_mappings_cache:
        term_mappings_cache[term] = query_openai_for_color(term)
        time.sleep(1)  # Add a delay to avoid overloading the server

# Apply the cached mappings back to each dataframe
retailer_colors_df['mapped values'] = retailer_colors_df['retailerColors'].apply(lambda x: term_mappings_cache.get(x, "check URL"))
retailer_garment_colors_df['mapped value(s)'] = retailer_garment_colors_df['retailerGarmentColors'].apply(lambda x: term_mappings_cache.get(x, "check URL"))

# Save the updated dataframes to a new Excel file
output_path_openai = 'Color_Normalization_OpenAI_Enhanced.xlsx'
with pd.ExcelWriter(output_path_openai) as writer:
    retailer_colors_df.to_excel(writer, sheet_name='RetailerColors Mapped', index=False)
    retailer_garment_colors_df.to_excel(writer, sheet_name='RetailerGarmentColors Mapped', index=False)

print("OpenAI-based color normalization completed. Results saved to 'Color_Normalization_OpenAI_Enhanced.xlsx'")
