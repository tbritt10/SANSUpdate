# SANSUpdate
A python script to compile an excel file report based on two reports from seperate databases. Made while working as a student assistant at JSU.

This works by taking a report from the offsite training database and the current employee database report and comparing them to find mismatches. If an employee is present in the training database but not in the employee database locally, then they are considered "inactive" (i.e. they no longer worked for JSU). If an employee is present in the employee report but not in SANS, they are considered new. The report generated at the end can be directly uploaded to the SANS database to remove old users and add new ones. This helped to prevent running out of licenses for training due to old employees.

A second use case was eventually added just to generate a report of users who have not completed training with all inactive users stripped from the file. This also adds the supervisor names for each user to the report since they were not present in the original files used for the comparison.

Requirements: python 3, openpyxl, SANS.xlsx (SANS database report), ActiveEE (JSU Employee database report), and Employees.xlsx (template file for updates on SANS).

In order to do a demo of the script:

Must have python version 3 or greater

Download all files and keep them in the same directory

Install openpyxl by running 'pip install openpyxl' in the command prompt or console.

Ensure the file names match the global filename variables in the python file.

File names and column letters are declared as global variables to allow for easily changing.

See documentation for openpyxl here: https://openpyxl.readthedocs.io/en/stable/index.html

--------------------------------------------------------------------------------------------------------------------------------

Class IncompleteRow:
This class is used to import the data for the report of employees who haven't completed training and for data that respresents a snapshot of the current users with licenses in SANS. Data is already present in the SANS report file, but users who are inactive must be stripped, and the supervisor must be pulled from another file and added. This class overrides == and != behavior to compare self.email rather than the entire object. This is to allow for comparisons between the different classes.

Class ActiveRow:
This class is used to create objects from the current employee report for comparisons against the SANS report. This class overrides == and != behavior to compare self.email rather than the entire object. This is to allow for comparisons between the different classes. This class features the compareRow() method that compares between different "row" objects (i.e. ActiveRow == IncompleteRow). 

For each files the program features a respective load function that uses the global column letter variables to choose where to pull data. Each has output to show the progress of loading files. 

Each of the respective files has a comparison function to do the job of finding inactive employees, new employees, and finding incomplete employees; as well as an export function to output these to a new excel files report.

The dataOutput() function has exception catches for FileNotFound errors which can occur if the user has a report open. The error output reflects solutions to the issues I encountered during testing. 

Tha main function checks the global variable options to determine which functions to run (i.e. you can disable incomplete) to cut down run time. 
