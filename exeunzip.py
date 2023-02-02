import os
import zipfile

def unzip_exes(directory):
    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        if os.path.isdir(filepath):
            unzip_exes(filepath)
        elif filename.endswith('.exe'):
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(directory)

unzip_exes('/path/to/folder')
