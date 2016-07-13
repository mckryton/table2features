# table2features
creates BDD feature files from an Excel table

## usage

* open table2features.xlsm with macros activated
* create a table with 4 columns domains, aggregates, features and scenarios
* select the domain in the first row of your table
* run the macro 'exportFeatures'

## system requirements
The current version of this macro runs only with Excel 2011 for Mac.

## background
The script was made for the case when you have to model an existing system. If the system has a lot of features it is quite exhausting to cerate every feature file manually. To save some effort you might now list all features in a table and let the script create all the file for you. 

The Excel file [table2features.xslm](table2features.xslm) contains a sample table. If you run the macro you should receive feature files like in [doc/sample_data_features/](doc/sample_data_features/).