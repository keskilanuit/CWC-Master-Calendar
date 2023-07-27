# CWC-Master-Calendar

-Populate Master Calendar
this Workflow is design to manage mass amount of return items and task end items for monthly basis project.
the coding will first loop through each line items from worksheet "Matrix" starts from Cell A3 (that's very first cell below the header line) through all rows that value ends at the column Z
we are going to set i = 3, as the data will starts from the 3rd row
after all rows are looped, it will looking at column M of each line items, if column M of a line that contains value of "X", code will copy the entire row from column A to L for all rows that has value of "X" from their column M.
then copy  range A to L and paste as value to the worksheet named "Master Calendar"
and set cell M from "Master Calendar" worksheet as text value "Not Start" after done pasting each line item.


-Export Master Calendar
this coding work as exporting the entire worksheet to a user selected folder with naming convention of "Master Calendar" + Current month  & Current year
code will take last row of data all the way up to row A.
set current year and current month as string 
show user folder picker to select a designated folder and start exporting worksheet and it's title + Month + year

if a file with same name already exist within the folder, provide option for user to overwrite the file.
