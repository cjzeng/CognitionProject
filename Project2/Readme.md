###Features
- Read all clicker result excel files from directory
- Combine the answers and user to single excel worksheet
- Discard call answers columns without header
- Convert letter answer into number answer

### Algorithm
- Load all filenames into array
- Open a new "summary" workbook
- For each answer workbooks defined in file list array
-- Open the workbook, and select the first worksheet
-- Load the workbook answer headers into source definition
-- Identify the row and column with data
-- Load data into array, also convert the answer value as it was read.
-- Close the source workbook
-- Select the "summary" workbook and worksheet
-- Merge the source definition into destination definition
-- Write out the combined header
-- Write out the stadard answer row
-- Insert the source data into destination row by
--- Use combined firstname/lastname as the key
--- Look up the row in look up table. If not found, add the key to the lookup table and get the row
--- Copy the data row from source to destination excel row


### Implementation Object
#### Class Source
- Stores header definition
- Map of header to excel column index
- Create field name list
- Map of Excel column index to field name
- Map of Excel data column to data table
- Data table to store the data

#### Class Destination
- Header definition
- Map of header name to Excel column index

#### Class NameLookup
- Look up array to store the keys
- Function to add key to lookup array
- Function to look up key and returns the row number

#### Class IngestTool
- Service to perform the process of combining the data from each Clicker Excel files into consolidated worksheet.

#### Module Work
Scarfold for calling the services

#### Module Test
Unit test for classes and modules

