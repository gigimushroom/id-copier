import pandas as pd
import os
import shutil

# Read Excel file and convert DataFrame to dict
df = pd.read_excel('a.xlsm', usecols = "N:O", header=None, sheet_name='Input', skiprows=range(1, 6))
data_dict = {d[13]: d[14] for d in df.to_dict('records') if pd.notna(d[13]) and pd.notna(d[14])}

# Define directories
img_dirs = ['idss/', 'idss2/']
#img_dirs = ['input/']

dst_dir = 'output/'

def find_and_copy(file_name, directory, destination):
    for img_file in os.listdir(directory):
        if img_file.startswith(file_name):
            shutil.copy(directory + img_file, destination)
            return True
    return False

count = 0
# Iterate through the dictionary
for key, value in data_dict.items():
    # Concatenate key and value
    file_name = str(key) + str(value)
    
    # Try to find the matching image in the directories
    if any(find_and_copy(file_name, dir, dst_dir) for dir in img_dirs):
        count += 1
    else:
        print(f"{file_name} not found")

print(count)
