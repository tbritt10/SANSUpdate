# SANSUpdate
A python script to compile an excel file report based on two reports from seperate databases. Made while working as a student assistant at JSU.

This essentially works by taking a report from the offsite training database and the current employee database report and comparing them to find mismatches. If an employee is present in the training database but not in the employee database locally, then they are considered "inactive" (i.e. they no longer worked for JSU). 

A second use case was eventually added just to generate a report of users who have not completed training with all inactive users stripped from the file. This also adds the supervisor names for each user to the report since they were not present in the original files used for the comparison.

Requirements: python 3, openpyxl, SANS.xlsx (SANS database report), ActiveEE (JSU Employee database report), and Employees.xlsx (template file for updates on SANS).

In order to do a demo of the script:

Must have python version 3 or greater

Download all files and keep them in the same directory

Install openpyxl by running 'pip install openpyxl' in the command prompt or console.

Ensure the file names match the global filename variables in the python file.

See documentation for openpyxl here: https://openpyxl.readthedocs.io/en/stable/index.html

File names and column letters are declared as global variables to allow for easily changing.
