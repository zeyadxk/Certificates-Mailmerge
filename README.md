# Certificates Mailmerge for macOS

### python script to add names from an xlsx spreadsheet on a jpeg image then exports them as PDFs.
### applescript to mailmerge the pdf files with a personalized message to the recipitent using Apple Mail.
# Dependencies
- Pillow library (pip install --upgrade Pillow)
- OpenPyXL library (pip install --upgrade openpyxl)
- Arabic Reshaper (pip install --upgrade arabic-reshaper)
- BiDi (pip install --upgrade python-bidi)
- Microsoft Excel for mailmerge feature

# How to Use?
## Spreadsheet:

- Create an excel spreadsheet
- Add names in column A, email address in column B, and serial number of the ceritficate in column C. You can add as many entries as you want. (a template is available named template.xlsx)
- Save the file. You can use any name, but if you change the name to anything other than template.xlsx, you will need to change that in the main.py file.

if you don't want to use the mailmerge feature and only the python script, you don't have to fill column B. However the name has to be in column A and the serial number in column C

## Certificate:

- Should be a JPG and it shoud be named cert.jpg or change it in main.py (sample is available)
- Find out the position you'd like the text to be in and edit the value in main.py

## The main.py script:

- Create an new folder named "Result" (make sure it's the same exactly)
- Put the script, the jpeg file, the xlsx file, and the Result folder into one folder
- Using terminal go to the same directory and run the command [python3 main.py]
- The command will output the PDF files in the Result folder

## The mailmerge applescript:

- Open the script in Script Editor
- Add the email content and the Pathname to the Result folder in the script (there are instructions there)
- Save the changes
- Open the excel spreadsheet you made earlier for the python script
- Save the spreadsheet as csv with the same name (template)
- Open Apple Mail
- Make sure you have selected the inbox of the eamil address you want the email to be sent by
- Open the csv file you just created using Microsoft Excel
- Make sure it is the only excel file that is open and NO OTHER FILE (the file should have the 3 columns as follows A for name, B for email address, and C for certificate serial number)
- Run the applescript in Script Editor.app

### Please use a test csv file with a small number of entries and your email address to test the script first

# How to contribute

- Raise an issue explaining what changes you intend to make and I will contact you there
- If the changes are approved, fork the repository and make your changes.
- Created a pull request
