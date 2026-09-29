# Py-Script-for-USE-TO-GENERATE-SEATING-ARRANGEMENT-FOR-SEMESTER-END-EXAMINATION
USE TO GENERATE SEATING ARRANGEMENT FOR SEMESTER END EXAMINATION, CAN SKIP ROWS / COLUMNS/AND EXPORT TO EXCEL
USER GUIDE & INSTALLATION MANUAL
**Register Number Based Multi-Allotment Seating Arrangement Application**
**STEP 1:** Install Python on Your Computer
Go to the official website python.org/downloads and click the download button for your operating system
(Windows or macOS).
CRITICAL WINDOWS REQUIREMENT: When the installer opens, you MUST check the box that says
'Add python.exe to PATH' at the bottom of the window before clicking 'Install Now'. If skipped, the app cannot launch.
**STEP 2:** Install Required Libraries
Open your computer's terminal or command line prompt:
• Windows: Press the Windows Key, type cmd, and hit Enter.
• macOS: Press Cmd+Space, type Terminal, and hit Enter.
Copy, paste, and run the following exact command, then hit Enter:
pip install pandas openpyxl pdfplumber
**STEP 3:** Run the Application Script
Locate your downloaded .py script file on your machine.
• Windows: Double-click the file to open the interface directly.
• Alternative (All Systems): Drag your terminal window to the script path or type python followed by a space, drag-and-drop the file into the window, and hit Enter.
**STEP 4:** Load Register Numbers
Click the Select Excel / PDF Files button in Section 1 of the app. Select your registration spreadsheets or PDF documents. The software automatically scans and extracts numerical strings strictly under detected headings like Register Number or Reg No, protecting against invalid data or text overlaps.
**STEP 5:** Configure the Room Layout
In Section 3, declare your room size parameters using the Rows ( up to 20) and Columns ( up to 10) adjusters. Select your preferred seat flow orientation from the dropdown menus into the Skip Rows or Skip Columns inputs (e.g., 3, 7-9).
**STEP 6:** Generate and Block Seats
Click Create / Refresh Grid to dynamically build your seating configuration layout. Click directly on any individual seat box within the visual window to lock or block that desk. The system instantly recalculates values, pushing records downward chronologically.
**STEP 7:** Export All Sheets to Excel
Click EXPORT ALL PAGES TO EXCEL, select a destination folder path on your computer, and name your report. The script automatically compiles a locked workbook featuring complete visual maps for every row/column arrangement alongside unified master candidate index sheets.
This is for informational purposes only.  AI responses may include mistakes.
