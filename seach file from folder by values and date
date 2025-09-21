Project Overview

This Excel VBA project provides a dynamic file search tool within a worksheet. It allows you to search for files in a specific folder by name or date, automatically list them in a formatted table, and create clickable hyperlinks to open files. The table updates in real-time when you enter a search value or date.

Features

Search by Value (File Name)

Type a keyword in cell A2.

The macro filters files containing that keyword in the folder.

Search by Date

Type a date in cell B2.

The macro lists files with a modified date on or after the entered date.

Dynamic Table

Table headers start at row 3 (File Name, File Type, Date Modified, Open Link).

Alternating row colors for readability.

Borders applied for a professional look.

Columns auto-fit for optimal display.

Minimum column widths enforced for clarity.

File Details

Column A: File Name (matching part highlighted in red and bold).

Column B: File Type (extension in uppercase).

Column C: Date Modified (dd-mmm-yyyy hh:mm format).

Column D: Clickable hyperlink to open the file.

User Interface Enhancements

Cells A1 (Search by value) and B1 (Search by date) formatted with colors, bold text, and font size 20.

Borders applied to header and search input cells.

Freeze panes applied from row 4 so headers stay visible while scrolling.

How to Use

Open the Excel workbook containing the VBA macros.

Make sure macros are enabled.

Navigate to the sheet where the macro is installed.

The top row should display:

A1: "Search by value"

B1: "Search by date"

Enter your search criteria:

A2: Type any part of the file name you want to search for.

B2: Type a date (optional) to filter files modified on or after that date.

The table below (starting at row 3) will update automatically when you change A2 or B2.

Click the “Open File” hyperlink in column D to open any listed file.

Folder Configuration

By default, the macro searches the folder:

D:\My Documents\Desktop\file\


To change the folder, edit this line in the Worksheet_Change event:

folderPath = "D:\My Documents\Desktop\file\"


Make sure the folder path ends with a backslash \.

Notes & Tips

Only files matching the search keyword and/or date will be listed.

If no search criteria are entered, all files in the folder are displayed.

Column widths auto-adjust but maintain minimum widths for readability.

Headers are frozen so you can scroll through long lists without losing context.

The macro automatically highlights the search term within file names.

Macros Included

AddSearchLabels

Creates the formatted search input cells (A1, B1).

Sets font size, color, boldness, and borders.

Worksheet_Change

Triggered when A2 or B2 changes.

Filters folder files based on search criteria.

Creates a dynamic, formatted table with hyperlinks and alternating colors.

callall

Simple macro to call AddSearchLabels and initialize the sheet.

Requirements

Microsoft Excel 2010 or later.

Macros must be enabled.

The folder specified must exist and contain files to search.

This setup provides a modern, visually appealing, and interactive way to search and access files directly from Excel.
