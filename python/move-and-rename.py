'''
This is a simple script to:
1. Search for files in a directory with a specific pattern
2. Create a new name for the file
3. Move the file to a new location
'''

import os
import re

# Inputs
origin_directory = r'C:\Test-dir\Move_and_rename\Archive'
target_directory = r'C:\Test-dir\Move_and_rename'
regex_search = r'AUDIT[.]*'
regex_replace = r'.* - '

# Get list of files
files = os.listdir(origin_directory)

# Loop over list of files
for file in files:
    # If file name contains search string, create new name and move file
    if re.search(regex_search, file) is not None:
        new_file = re.sub(regex_replace, '', file)
        full_new_file = os.path.join(target_directory, new_file)
        full_file = os.path.join(origin_directory, file)
        os.rename(full_file, full_new_file)
