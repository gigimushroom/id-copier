import pandas as pd
import os
import shutil

# Reading Excel file
df = pd.read_excel('a.xlsm', usecols = "N:O", header=None, sheet_name='Input', skiprows=range(1, 6))
# Convert the DataFrame to a list of dictionaries
list_of_dicts = df.to_dict('records')

data_dict = {d[13]: d[14] for d in list_of_dicts if pd.notna(d[13]) and pd.notna(d[14])}
print(data_dict)

#Directory containing images
img_dir = 'input/'

# Directory to copy images to
dst_dir = 'output/'

# Iterate through the dictionary
for key, value in data_dict.items():
    # Concatenate key and value
    file_name = str(key) + str(value)

    # Try to find the matching image in the directory
    for img_file in os.listdir(img_dir):
        # If the file name starts with the concatenated string
        if img_file.startswith(file_name):
            # Copy the image to the destination directory
            shutil.copy(img_dir + img_file, dst_dir)

