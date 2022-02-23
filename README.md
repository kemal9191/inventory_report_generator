# Packages

### python-docx
Use `pip install python-docx` to install python-docx. After the installation, use `import docx` to import the package, not `import python-docx`

### pandas
Use `pip install pandas` to install pandas library. After the installation, use `import pandas as pd` to import the package.

### openpyxl
Use `pip install openpyxl` to install optional dependency ***openpyxl***. You do not need to import this dependency on `app.py`, pandas needs it to operate.

# How It Works

## What It Does
This small program has been designed to automate the creation of personalized reports for a personality inventory.

## Input
It takes results of people who complete the inventory from an excel document named as **results.xlsx** located on **/data** directory. If you want to process some results for a sample, you should modify that document.

The program uses input from **report_data.xlsx** located on **/data** directory to generate personal reports. The file has report statements for low, average, and high scores in terms of all factors. 

## Output
After taking necessary data from excel sheet, the program processes all the data in order to create personalized inventory reports. It loops through each person, and creates reports for each of them. Personalized reports are saved on **/results** directory, named after the person they belong.


