import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.neighbors import NearestNeighbors

# Load the spreadsheet
file_path = 'Color Normalization.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load the relevant sheets
normalized_colors_df = excel_data.parse('Normalized List of Colors')
color_df = excel_data.parse('Color')
retailer_colors_df = excel_data.parse('RetailerColors')
retailer_garment_colors_df = excel_data.parse('RetailerGarmentColors')

# Fill NaN values in aliases with empty strings
normalized_colors_df['aliases'].fillna('', inplace=True)

# Prepare the training data for the ML model
aliases = normalized_colors_df['aliases'].str.replace("\n", " ").values
colors = normalized_colors_df['colors'].values

# Use TF-IDF vectorizer to convert the text into numerical vectors
vectorizer = TfidfVectorizer()
X_train = vectorizer.fit_transform(aliases)

# Build a Nearest Neighbors model to find the closest matching normalized color for any given retailer color
nn_model = NearestNeighbors(n_neighbors=3, metric='cosine')
nn_model.fit(X_train)

# Function to map retailer colors to normalized values
def map_colors_batch(input_colors, model, vectorizer, color_labels):
    # Convert all input colors at once to their vector representations
    color_vectors = vectorizer.transform(input_colors)
    # Find the nearest neighbors in the training data for all input colors
    distances, indices = model.kneighbors(color_vectors)
    # Map the indices to the corresponding color labels
    mapped_values = [color_labels[idx[0]] for idx in indices]
    return mapped_values

# Clean non-string values in 'retailerColors' and 'retailerGarmentColors' columns
retailer_colors_df['retailerColors'] = retailer_colors_df['retailerColors'].astype(str)
retailer_garment_colors_df['retailerGarmentColors'] = retailer_garment_colors_df['retailerGarmentColors'].astype(str)

# Map the retailer colors and retailer garment colors
retailer_colors_mapped = map_colors_batch(retailer_colors_df['retailerColors'].values, nn_model, vectorizer, colors)
retailer_garment_colors_mapped = map_colors_batch(retailer_garment_colors_df['retailerGarmentColors'].values, nn_model, vectorizer, colors)

# Update the dataframes with the mapped values
retailer_colors_df['mapped values'] = retailer_colors_mapped
retailer_garment_colors_df['mapped value(s)'] = retailer_garment_colors_mapped

# Save the updated dataframes to new sheets in the Excel file
with pd.ExcelWriter('Color_Normalization_Updated.xlsx') as writer:
    retailer_colors_df.to_excel(writer, sheet_name='RetailerColors Mapped', index=False)
    retailer_garment_colors_df.to_excel(writer, sheet_name='RetailerGarmentColors Mapped', index=False)

print("Color mapping completed and saved to 'Color_Normalization_Updated.xlsx'")
