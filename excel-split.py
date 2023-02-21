import openpyxl

# Load the original Excel file and sheet
workbook = openpyxl.load_workbook('original_file.xlsx')
sheet = workbook.active

# Get the data from the sheet and create a list of dictionaries
data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    name, age, nationality, sex = row
    age = age.split(',') if isinstance(age, str) else ['']
    data.append({
        'Name': name,
        'Age': age,
        'Nationality': nationality.split(','),
        'Sex': sex.split(',')
    })

# Explode the data into a list of dictionaries
exploded_data = []
for item in data:
    for i in range(len(item['Name'])):
        exploded_data.append({
            'Name': item['Name'][i],
            'Age': item['Age'][i] if i < len(item['Age']) else '',
            'Nationality': item['Nationality'][i] if i < len(item['Nationality']) else '',
            'Sex': item['Sex'][i] if i < len(item['Sex']) else ''
        })

# Write the exploded data to a new sheet
new_workbook = openpyxl.Workbook()
new_sheet = new_workbook.active
new_sheet.append(['Name', 'Age', 'Nationality', 'Sex'])
for item in exploded_data:
    new_sheet.append([item['Name'], item['Age'], item['Nationality'], item['Sex']])

# Save the new workbook to a new file
new_workbook.save('new_file.xlsx')
