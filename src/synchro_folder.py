import os
import json


def create_file_dict(directory,out):
    # Create an empty dictionary to store file names and paths
    file_dict = {}

    # Iterate through all files in the folder
    for filename in os.listdir(directory):
        if filename.endswith('.xlsm'):

            # Get the full path of the file
            filepath = os.path.join(directory, filename)
            # Add the file name and path to the dictionary
            file_dict[filename] = filepath
            print(filename)

    # Create a JSON file and write the file dictionary to it
    with open(out, "w") as outfile:
        json.dump(file_dict, outfile)
        print(outfile)

if __name__ == '__main__':
    print('done')

