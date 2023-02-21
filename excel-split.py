import pandas as pd

# Load the original Excel sheet into a pandas DataFrame
df = pd.read_excel('original_file.xlsx')

# Split the values in the "Name", "Age", "Nationality", and "Sex" columns
df = df.assign(Name=df['Name'].str.split(', '), 
               Age=df['Age'].str.split(',').apply(lambda x: [int(i) for i in x]), 
               Nationality=df['Nationality'].str.split(', '), 
               Sex=df['Sex'].str.split(','))

# Explode the DataFrame on the "Name", "Age", "Nationality", and "Sex" columns
df = df.explode(['Name', 'Age', 'Nationality', 'Sex'])

# Reset the index of the DataFrame
df = df.reset_index(drop=True)

# Create a new DataFrame with the desired format
new_df = pd.DataFrame({
    'Name': df['Name'],
    'Age': df['Age'],
    'Nationality': df['Nationality'],
    'Sex': df['Sex']
})

# Write the new DataFrame to a new Excel file
new_df.to_excel('new_file.xlsx', index=False)
