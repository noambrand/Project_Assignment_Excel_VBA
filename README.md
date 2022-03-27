# Project_Assignment_Excel_VBA
Automatically create personalized tables of assignments for each employee, allowing them to focus on only their assignments and priorities and not be distracted by other employee assignments that do not concern them. Create management reports with all projects and employees together.

How does it work: It combines different sheets with the same fields to one table, copies the combined table to a new sheet and
Filters it by employee name to separate sheets, AutoFits the table nicely and finally saves the sheets to separate excel files
in a new folder that has the name of the workbook.
The AutoFilter in the code is in columns E,F (See Fields 5,6 in the AutoFilter)

Macro steps:
1) Combines Number of first sheets with the same fields to one sheet.
2) Makes copies the combined sheet to the end of the workbook.
3) Rename sheets by each employee name.
4) Filter each sheet by employee name +not yet complite assignments.
5) Each sheet still has all of the assignments(the user can cancel filter by autofilter).
6) Autofit columns width and sheet from right to left
7) Split the workbook sheets to separate files in a folder.
   The file name will be the sheet name + date of creation.

The example has 4 sheets- 3 big projects that have many assignments and 1 sheet that has 4 small projects with few assignments.
Ps. If the assignment is e.g. "Jerry+Noam" it will be present in Jerry's table and also Noam's table.

Question: Why not create a table with all the projects and assignments in the first place that will save the need to combine different sheets?
Answer: When your table gets filled with projects that have many assignments it becomes difficult to manage efficiently, the table gets messy and editing is more complicated with autofilter on,  you have to keep filtering the table which is time-consuming and prone to errors.
