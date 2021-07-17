# Excel_PowerQuery_ETL_Demonstration_2
Brief example of data import from different files (.accdb, .XLSM, .xslx, .txt), cleaning using Power Query and load to a single Pivot Table Report

Steps made it:
1) Import Excel Files From Folder
2) Transform extensions to all lowercase
3) Filter to include only Excel Files in import process
4) Extract Excel File Name to create New Column for City. Split By Delimiter
5) Rename Column and Remove unwanted columns
6) Add Custom Column with Excel.Workbook Function (M Code Function). Explanation of what functions extracts from the Excel Files
7) Filter Out Excel Objects that do not meet Criteria = Sheet
8) Filter out names that Do Not Begin With Sheet. Extract Worksheet Name to create New Column for SalesRep
9) Final Append to get all Excel Worksheet that contain Proper Data Sets with a proper SalesRep Name
10) Apply correct Data Types
11) Load to Excel Sheet
12) Change Default PivotTable Layout & Options
13) Build PivotTable Report
14) Add New Excel Workbook Files to the Folder & Refresh the Query and PivotTable
