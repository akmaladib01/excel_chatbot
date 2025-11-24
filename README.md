🐍 Excel/CSV Search Chatbot

Hey there! 👋 This little Python chatbot helps you search through all your Excel and CSV files super easily. Just type what you’re looking for, and it will find the matching rows and show them nicely in your browser.

No more scrolling endlessly in Excel!

✨ What It Can Do

Search multiple Excel sheets and CSV files at once

Case-insensitive search (so CRITICAL = critical)

Shows only the rows containing your keyword

Saves results as HTML for clean viewing in your browser

Leaves your original files completely untouched

🛠 Requirements

Python 3.7 or higher

Python library: pandas

Install pandas if you haven’t yet:

pip install pandas

📂 How to Set Up Your Files

Make a folder for all your Excel/CSV files. Example: Excel_Files

Put all your .xlsx, .xls, and .csv files there

Excel files can have multiple sheets

CSV files are treated as single sheets

In the script, update the folder path if needed:

folder_path = "Excel_Files"

🚀 How to Use It

Run the script:

python chatbot_excel_search.py


The script will show all your files:

Found 3 file(s) in 'Excel_Files':
1) file1.xlsx
2) file2.xlsx
3) file3.csv


Type your search keyword when prompted:

You: critical


If matches are found, you’ll see:

📄 Matches found in File: file1.xlsx | Sheet: Sheet1 (3 row(s))


The matching rows automatically open in your browser as a clean HTML table.

Type exit to quit:

You: exit
Goodbye!

💾 Where Results Go

All HTML results are saved in the folder search_results

Filenames follow this format:

<OriginalFileName>_<SheetName>_results.html


Example: file1_Sheet1_results.html

💡 Tips

Only rows containing your search keyword are shown

Works with multiple sheets in Excel

Search multiple times — each query opens a new HTML result

Original files never get changed
