import pandas as pd

# Read the original Excel file into a DataFrame
df = pd.read_excel('original_file.xlsx')

# Explode the DataFrame
exploded_df = df.explode('Name').explode('Age').explode('Nationality').explode('Sex')

# Remove rows with any missing values
exploded_df = exploded_df.dropna()

# Write the exploded data to a new sheet in a new Excel file
with pd.ExcelWriter('new_file.xlsx') as writer:
    exploded_df.to_excel(writer, index=False)
