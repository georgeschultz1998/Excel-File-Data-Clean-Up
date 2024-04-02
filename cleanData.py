import pandas as pd
import re
import datetime


# Define the path to your Excel file
file_path = 'C:/Users/georg/OneDrive/Desktop/ProjectsForEmployers/Data Management/SSM_Data_Set_2018.xlsx'

# Load the Excel file into a pandas DataFrame without headers
df = pd.read_excel(file_path, header=None)

# If the last two columns are blank and you want to remove them, select only the first 6 columns
df = df.iloc[:, :6]

# Define your new headers
new_headers = ['Name', 'Date', 'Comments', 'Code', 'Reason', 'Position']

# Set the new headers
df.columns = new_headers


################################
# Names corrections
name_corrections = {
    'Alexander Nardone': ['Alexander T. Nardone', 'Alexander Troy Nardone'],
    'Christina Kishitomato': ['Christina tomato?'],
    'George Quincy': ['George Q'],
    'Matt Shawleyey': ['Matthias Shawleyey'],
    'Tess Montehue': ['Tess Mont.']
}

for correct_name, variations in name_corrections.items():
    df['Name'] = df['Name'].replace(variations, correct_name)

################################
# Dates corrections
df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.strftime('%m/%d/%Y')

################################
# Comments corrections

def standardize_time(comment):
    # Convert the comment to a string if it's a datetime.time object
    if isinstance(comment, datetime.time):
        comment_str = comment.strftime('%H:%M:%S')
    else:
        comment_str = str(comment)
    
    # Define a time pattern that matches hours, optional minutes and seconds
    time_pattern = re.compile(r'(\b\d{1,2}:\d{2}(?::\d{2})?)')
    
    # Check if the comment contains a time pattern
    times = time_pattern.findall(comment_str)
    
    # If there are time patterns found, concatenate them with 'arrival'
    if times:
        # Join the times with 'arrival', removing leading zeros from hours and seconds
        standardized_times = []
        for time in times:
            parts = time.split(':')
            hour = parts[0].lstrip("0")
            minute = parts[1]
            second = parts[2].lstrip("0") if len(parts) == 3 else None
            if second:
                standardized_times.append(f'{hour}:{minute}:{second} arrival')
            else:
                standardized_times.append(f'{hour}:{minute} arrival')
        return ', '.join(standardized_times)
    else:
        return comment

# Apply the function to the 'Comments' column
df['Comments'] = df['Comments'].apply(standardize_time)
df['Comments'] = df['Comments'].str.capitalize()


################################
# Code Corrections using a method similar to the 'Positions Corrections'
df['Code'] = df['Code'].str.replace('1/2', 'Half-Day Absence', regex=False) \
                        .replace('Abs.', 'Absent') \
                        .replace('T', 'Tardy') \
                        .fillna('Unspecified')


################################
# Positions Corrections
df['Position'] = df['Position'].str.replace('ops', 'Operations', regex=False).fillna('Unspecified')

################################
# Return updated excel file
updated_file_path = 'C:/Users/georg/OneDrive/Desktop/ProjectsForEmployers/Data Management/SSM_Data_Set_2018_Cleaned.xlsx'
with pd.ExcelWriter(updated_file_path, engine='openpyxl', date_format='m/d/yyyy') as writer:
    df.to_excel(writer, index=False)

print(f"DataFrame has been cleaned and saved to {updated_file_path}.")

