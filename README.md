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
