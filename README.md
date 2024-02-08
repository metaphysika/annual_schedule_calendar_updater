# annual_schedule_calendar_updater

## Clear report file
A new report file is created with each instance of the program running.  The old report file is opened and cleared of data first.

## Directory Crawl
Next, the program crawls through a directory looking for specified file types.  In this case, it is .xlsx and .xlsm.
All the files in the directory structure follow a naming convention of YYYYMMDD_ file.xlsx

The script adds all file names matching the correct file type in a subdictory to a list and then finds the file name with the max number value in it.  This is the newest file.  It then opens the file and looks for specific values in named cells.  

## Update data in a database 
The program then updates the named cell values and in a database file. 


## Add data to report that are older than 1 year or missing Review Date
Next the program checks how old two date values are in the report.  If the date in the file named cell is less than one year, the program moves on to check the review date.

Anything that is over one year old in the Survey Date named cell gets added to the report file. 

If the "Review Date" named cell is blank the program will add this data to a second sheet on the report file.


# Equipment Testing Library Automation

This script automates the process of managing equipment testing reports by scanning directories for report spreadsheets, extracting relevant data, and updating a summary sheet with testing and review dates.

## Description

The script is designed to work within a specific organizational structure, handling surveys for Radiography, Fluoroscopy, CT, Mammography, MRI, NM, PET, and other related equipment. It utilizes the `openpyxl` library to interact with Excel spreadsheets, updating records with new testing dates, and marking equipment as due or overdue for testing or review. It also makes use of a custom email script. 

## Installation

To run this script, you'll need Python installed on your system along with the required dependencies.

1. Ensure that you have a Python environment set up.
2. Install the required libraries using pip:


pip install openpyxl shutil datetime time py os getpass

## Usage

Before running the script, make sure to update the dirBase, dirBase2, and dirBase3 variables to reflect the paths where your reports and summary sheets are stored.

To execute the script, run the following command in your terminal:

python path_to_script.py

When prompted, enter the password for the email account that will be used to send notifications.

## Contributing

Contributions to improve the script or adapt it for other organizational structures are welcome. Please ensure to comment on your code and maintain the existing structure for readability and maintainability.

## Notes

    Ensure you have the necessary permissions to read from and write to the specified directories.
    The email functionality requires proper configuration of the emailsender_webmail module with your email server settings.
    Always back up your spreadsheets before running the script to prevent accidental data loss.

## Disclaimer

This script is provided "as is", without warranty of any kind. Use it at your own risk.