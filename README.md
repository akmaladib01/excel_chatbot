Excel Chatbot
Overview

This Python script allows you to search through multiple Excel (.xlsx, .xls) and CSV files in a folder.

Searches are case-insensitive.

Only rows containing the search keyword are returned.

Matching rows are saved as HTML files and opened automatically in your default browser.

Original files remain unchanged.

Requirements

Python 3.7+ installed on your system

Python libraries: pandas

Install required library
pip install pandas

Folder Setup

Create a folder to store all your Excel/CSV files (example: Excel_Files).

Place all your .xlsx, .xls, and .csv files in that folder.

Each Excel file can have multiple sheets.

CSV files are treated as a single sheet.

Update the folder path in the script if needed:

folder_path = "Excel_Files"  # change to your folder path

How the Script Works

Scans the folder for all Excel and CSV files.

Reads each file and automatically detects the header row.

Cleans the data by removing empty rows and columns.

Loads all sheets/files into memory for fast searching.

Starts a chat loop in the terminal where you can type your search query.

Searches all sheets/files for rows containing your keyword (case-insensitive).

Saves matching rows into HTML files inside a folder called search_results and opens them in your browser.

How to Use

Run the script:

python chatbot_excel_search.py


The script will display all detected files:

Found 3 file(s) in 'Excel_Files':
1) file1.xlsx
2) file2.xlsx
3) file3.csv


Type your search keyword:

You: critical


If matches are found, you’ll see something like:

📄 Matches found in File: file1.xlsx | Sheet: Sheet1 (3 row(s))


The matching rows will open automatically in your default browser as an HTML table.

Type exit to stop the chatbot:

You: exit
Goodbye!

Output

HTML files are saved in the folder search_results.

File names format:

<OriginalFileName>_<SheetName>_results.html


Example: file1_Sheet1_results.html

Notes / Tips

Only rows containing the keyword are included in the results.

The search is case-insensitive.

Original Excel/CSV files are never modified.

Works with multiple sheets in Excel.

You can search multiple keywords in separate prompts.
