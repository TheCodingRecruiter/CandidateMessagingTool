import openpyxl
import pandas as pd
from pprint import pprint
import json

# Read excel document - Insert excel document in there with the 'ActiveJobs.xlsx' title or rename the file below.  Sheet_name starts at 0 and you can change sheet by iterating up until final sheet if multiple sheets are used
excel_data_df = pd.read_excel('ActiveJobs.xlsx', sheet_name=2)

# Convert excel to string 
# (define orientation of document in this case from up to down)
thisisjson = excel_data_df.to_json(orient='records')

# Make the string into a list to be able to input in to a JSON-file
thisisjson_dict = json.loads(thisisjson)

# Define file to write to and 'w' for write option -> json.dump() 
# defining the list to write from and file to write to
with open('data.json', 'w') as json_file:
    json.dump(thisisjson_dict, json_file)

# Working with the JSON file data and converting it to a dictionary
file = open("data.json")
json_dictionary = json.load(file)
file.close()

######### Review The Dictionary Items to create variables if required for message rewording.  Using Pretty Print for easier reading and calling a single instance to eliminate too much screen real estate being taken ###########
pprint(json_dictionary[0])


######## Variables for the message
# New line variable
nl = '\n'

# Personal self scheduling link which can be altered as needed
availability = 'https://prelude.amazon.com/'

# Iterating over the dicitonary objects and accessing the value of the keys
for i in json_dictionary:
  # Define variables based upon heading of each column to extract the value of the cell for each of these
  name = (i["Candidate Name"])
  email = (i["Email"])
  phone = (i["Phone Number"])
  role = (i["Role"])
  id = int((i["ID"]))
  status = (i["Status"])
  
  # Creating the message for each candidate using F-Strings
  message = (
    f"{phone}"
    f"{nl}"
    f"The status is currently: {status}"
    f"{nl}"
    f"Hello {name}.  I'm a recruiter at Amazon and was sending you this message in regards to your application for the {role} position ({id}).  I would like to schedule some time for us to meet. Here is a link to my availability.  Please feel free to schedule some time for us. at your earliest convienence.{nl} {availability} "
  )
  print(message)
  print()

