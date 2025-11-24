📊 Excel/CSV Search Chatbot
A Python script that enables fast, case-insensitive keyword searches across multiple Excel (.xlsx, .xls) and CSV (.csv) files in a folder. Matching rows are saved as HTML tables and opened automatically in your browser — without modifying original files.

✅ Features
🔍 Case-insensitive full-row search across all sheets and files
📁 Supports multiple Excel files (multi-sheet) and CSV files
🧹 Automatically detects headers and cleans empty rows/columns
🖥️ Results rendered as HTML tables and opened in your default browser
🛡️ Original files remain completely unchanged
💬 Interactive terminal-based chat interface
📁 Results saved in search_results/ folder for review
📦 Requirements
Python 3.7 or higher
Required libraries:
pandas
webbrowser (standard library)
os (standard library)
Install Dependencies
bash


1
pip install pandas
(No additional installation needed for os or webbrowser.)

🗂️ Folder Setup
Create a folder (e.g., Excel_Files) to store all your data files.
Place your .xlsx, .xls, and .csv files inside it.
Excel files may contain multiple sheets — all will be scanned.
CSV files are treated as a single sheet.
Update the folder path in the script if needed:
python


1
folder_path = "Excel_Files"  # ✅ Change this to your folder path
⚙️ How It Works
Scan: The script scans the specified folder for all Excel and CSV files.
Load & Clean:
Reads each file (auto-detects headers).
Removes empty rows and columns for cleaner searching.
Loads all data into memory for fast querying.
Search Loop:
Terminal-based chat interface starts.
Enter a keyword → script searches all rows across all files/sheets.
Only rows containing the keyword (anywhere, case-insensitively) are retained.
Output:
Matching rows saved as HTML tables in search_results/.
Files named as:
<OriginalFileName>_<SheetName>_results.html
(e.g., file1_Sheet1_results.html)
HTML files auto-open in your default browser.
▶️ How to Use
Run the script:
bash


1
python chatbot_excel_search.py
You’ll see a list of detected files:


1
2
3
4
Found 3 file(s) in 'Excel_Files':
1) file1.xlsx
2) file2.xlsx
3) file3.csv
Enter a search term:


1
You: critical
If matches are found:


1
📄 Matches found in File: file1.xlsx | Sheet: Sheet1 (3 row(s))
→ Matching rows open instantly in your browser 🌐.

To exit:


1
2
You: exit
Goodbye!
📁 Output Format
All results are saved in a subfolder named search_results/.

📄 Example filename:
file1_Sheet1_results.html

Each HTML file contains a clean, readable table of matching rows — perfect for reporting or sharing.

💡 Notes & Tips
✅ Search is case-insensitive (Critical, CRITICAL, critical → all match).
✅ Supports multiple keywords — just run separate searches (e.g., search "urgent", then "high-priority").
🚫 Original Excel/CSV files are never modified — safe for production data.
🚀 Data is loaded once at startup → subsequent searches are fast.
🧩 Works with messy data: skips blank rows/columns and infers headers.
