# Excel VB Script to import files and process the data in work sheet
Features
* Import multiple csv files to Excel sheet named after file name.
* Delete irrelavant rows
* Rename the selected column header
* Delete irrelavant columns

Issues:
* On Mac, list files function using DIR does not work. It lists only first file. 
* On Mac, list files using dir with wild card does not work.
* On Mac, list files using CreateObject does not work since mac does not have COM/ActiveX support.
* On Mac, there is no VBStudio for development.