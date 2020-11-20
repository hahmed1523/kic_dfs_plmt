# Current Process

 - Use Salesforce report builder to create needed report. 
 - Export the report to a Microsoft Excel spreadsheet.
 - Open the report and add a column to the with a forumula to determine if the person ID was in the previous week's report.
 - If the ID is in the previous report that the value for the cell should be "--" otherwise it should say "NEW".
 - Highlight the rows if the ID is new.
 - Sort by the new IDs.

 # New Process

 This is an extension to the other kic program. As the other program, I directly query Salesforce through python using the simple salesforce module and a SOQL query. I then normalize the json result and create a dataframe. With the dataframe I can add the NEW formula, color the rows, and sort the data.

 Since this program is an extension of the other kic program, I do not want to query salesforce twice for the same data. The goal for this program is to be part of a main program that will run all of the weekly reports. Therefore, I check if there is already a dataframe passed into the program function and if so, then it uses that dataframe rather than sending another query request to Salesforce. 

 I also added logic to the program to look at the folder where all the reports are stored and to search for the last report using the datetime module. 

 This program combines different techiques thats I have used in my other programs such as using simple salesforce, pandas, and openpyxl. 