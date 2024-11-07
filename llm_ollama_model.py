import pandas as pd
import requests
import time

# Load the spreadsheet
file_path = 'Color Normalization.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load the relevant sheets
normalized_colors_df = excel_data.parse('Normalized List of Colors')
retailer_colors_df = excel_data.parse('RetailerColors')
retailer_garment_colors_df = excel_data.parse('RetailerGarmentColors')

# Generate a list of unique color terms to reduce API calls
unique_terms = pd.concat([retailer_colors_df['retailerColors'], retailer_garment_colors_df['retailerGarmentColors']]).unique()

# Cache to store the results from Ollama API calls
term_mappings_cache = {}

# Function to query Ollama API for color mapping
def query_ollama_for_color(term):
    try:
        url = "http://localhost:11434/api/generate"
        headers = {
            "Content-Type": "application/json"
        }
        payload = {
            "model": "llama3.1:8b",
            "prompt": f"Identify the main color for '{term}'. Return '[blank]' if it is not a color-related term, or 'check URL' if it implies multiple colors."
        }
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        result = response.json()
        return result['completion'].strip()
    except Exception as e:
        print(f"Error with term '{term}': {e}")
        return "check URL"  # Fallback in case of an API error

# Query the API for each unique term
for term in unique_terms:
    if term not in term_mappings_cache:
        term_mappings_cache[term] = query_ollama_for_color(term)
        time.sleep(1)  # Add a delay to avoid overloading the server

# Apply the cached mappings back to each dataframe
retailer_colors_df['mapped values'] = retailer_colors_df['retailerColors'].apply(lambda x: term_mappings_cache.get(x, "check URL"))
retailer_garment_colors_df['mapped value(s)'] = retailer_garment_colors_df['retailerGarmentColors'].apply(lambda x: term_mappings_cache.get(x, "check URL"))

# Save the updated dataframes to a new Excel file
output_path_ollama = 'Color_Normalization_Ollama_Enhanced.xlsx'
with pd.ExcelWriter(output_path_ollama) as writer:
    retailer_colors_df.to_excel(writer, sheet_name='RetailerColors Mapped', index=False)
    retailer_garment_colors_df.to_excel(writer, sheet_name='RetailerGarmentColors Mapped', index=False)

print("Ollama-based color normalization completed. Results saved to 'Color_Normalization_Ollama_Enhanced.xlsx'")
