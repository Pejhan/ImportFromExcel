# ImportFromExcel
This Stored Procedure can make importing excel worksheets into SQL Server Database as easy as copy-pasting.

## Purpose
The purpose of this script is to ease the process of importing Excel worksheets into Database Table Objects. As per the shorcomings that will be mentioned later bear in mind that this is not the best nor the most effecient way to import your data into a SQL Database But if your looking for the easiest way to get things on the roll scroll to the next section.

## How it works
After executing the imp.sql file against your database you could easily import your excel by calling the Stored Procedure dbo.sp_ImportFromExcel and feeding it the content of your excel spreadsheet (The data should be selected and copied to clipboard) as your first argument. The second argument is optional and is the name of the table that you want to be created and the data to be imported into. If left empty the table's name will be the default value of ImportedFromExcel.
This Procedure does not violate SQL's constraints and brings some of it's own so while working with it consider the following:
- Your data must have headers and the first row will be chosen as column headers.
- You may not have two columns with the same name.
- You may not create a temp table using this method.
- You may not import rows with different number of columns.

## Shortcomings
Please do consider that if you intend to work with a large amount of data (say a million records or less considering the number of columns), or a normalized and effecient handling of such task please consult other guides or try SQL Server's own Import and Export Wizard that prompts you on each step on how do you want to handle things in order to get just the right result.\
Inserting the way this script is trying to do is considered non-effecient and it could overuse a lot of unneccessary resources. There seems to be better alternatives such as BULK INSERT or OPEN ROWSET, which will be examined later maybe.
Also for the sake of ease and comprehensiveness the data type for all columns no matter the formatting is chosen to be VARCHAR(MAX) which in some cases might be considered over the top. I have no plans for this yet since it hasn't caused me any issues yet.
