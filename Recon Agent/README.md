



A practical setup guide for installing, launching, and using the AI Data Reconciliation Agent to compare DGW source data with Workday target extracts. 




Document Version 	1.0 
Application 	AI Data Reconciliation Agent Web App 
Primary Use 	Reconciliation of Legacy data & Workday post-load data 




Quick outcome: After setup, users can upload a source DGW file and a Workday extract, run reconciliation, view discrepancy counts, and export an Excel reconciliation report. 




































Contents
•	1. Purpose
•	2. System Requirements
•	3. Install and Verify Python
•	4. Install Application Dependencies
•	5. Start the Web Application
•	6. Use the Application
•	7. Discrepancy Classifications
•	8. Dashboard and Report Output
•	9. Performance Guidelines
•	10. Troubleshooting
•	11. Recommended Use in Workday Projects



1. Purpose

The AI Data Reconciliation Agent is a web application that automates comparison between source data files, such as DGW extracts, and target data files, such as Workday validation reports. It helps project teams identify missing records, field-level mismatches, formatting issues, duplicate keys, and other exceptions so reviewers can focus on resolution instead of manual file-by-file checking.


2. System Requirements
Requirement Area 	Recommended Setup 
Operating System 	Windows 10 or Windows 11 
Python 	Python 3.11 or higher and Visual Studio 
Browser 	Microsoft Edge or Google Chrome 
Review Tool 	Microsoft Excel for output report review 
Input Files 	Excel .xlsx  
Python Libraries
 	flask, pandas, openpyxl, xlsxwriter  


3. Install and Verify Python

1.	Check whether Python is installed: Open PowerShell and run: python --version
2.	If Python is not installed: Install Python from the official Python website and select Add Python to PATH during installation.
3.	Restart PowerShell: Close and reopen PowerShell after installation so PATH changes are refreshed. 
4.	Validate again: Run python --version. You should see a Python version number instead of a Microsof Store alias or command-not-found message.
5.	Official download link: Python downloads




4. Install Application Dependencies
Navigate to the folder where the application file is saved, then install the required packages. Either install it from CMD prompts or run Visual studio > New terminal and use the commands - pip install <package_name>

Task 	PowerShell Command 
Navigate to the app folder 	cd "C:\Users\<username>\OneDrive - Accenture\Workday\AI Recon Agent" 
Install required libraries 	pip install flask pandas openpyxl xlsxwriter numpy 
Optional requirements file command 	pip install -r requirements.txt 
Validate installed packages 	pip list 


5. Start the Web Application
Launch the app locally from PowerShell using Visual Studio or similar tool. Open new terminal and enter below commands. Ensure to save the Python code file in the right folder and use the path accordingly.

•	cd "C:\Users\<username>\OneDrive - Accenture\Workday\AI Recon Agent"
•	python "Recon Agent.py"
•	
Expected browser URL: http://127.0.0.1:5000

Open the URL to use the tool.


6. Use the Application
Step 	Action 	What to Do 
1 	Upload Source File 	Upload the DGW or legacy source extract. 
2 	Upload Workday File 	Upload the Workday target extract or validation report. 
4 	Run Reconciliation 	Start processing. The app follows: Load Data -> Normalize -> Match Keys -> Compare Fields -> Detect Issues -> Generate Reports. 
5 	Review Output 	Use the dashboard and downloadable Excel report to review exceptions. 








7. Discrepancy Classifications
Classification 	Meaning 
Value Mismatch 	Values differ between source and target. 
Missing Record 	Record exists in source but not in Workday target extract. 
Extra Record 	Record exists in Workday target extract but not in source. 
Case Difference 	Text differs only by capitalization. 
Format Difference 	Value is equivalent but appears in a different date, number, or text format. 
Trim Issue 	Leading or trailing spaces are detected. 
Data Type Issue 	Source and target values are stored or interpreted as different data types. 
Duplicate Key 	Primary key appears more than once in one or both files. 
Column Structure Issue 	Expected columns are missing, renamed, or misaligned. 

8. Dashboard and Report Output
The dashboard should allow the reviewer to confirm whether processing is complete and quickly understand the reconciliation result.
•	Total source records
•	Total target records
•	Matched records
•	Mismatched records
•	Missing records
•	Extra records
•	Total discrepancies
•	Total fields compared

The downloadable Excel reconciliation report should include at minimum:
•	Summary Reconciliation Report
•	Detailed Discrepancy Report
•	Full Record Dump, if enabled
•	Column Mismatch Analytics, if enabled

9. Performance Guidelines
•	Use a clean, unique primary key wherever possible.
•	Remove unnecessary columns and sheets (if possible) before upload when validating very large files.
•	Close files in Excel before uploading to the app.
•	Validate duplicate keys before running final reconciliation.
•	Keep source and target column names consistent where practical.








10. Troubleshooting

Issue 	Recommended Fix 
Python was not found 	Install Python from the official website, enable Add Python to PATH, then restart PowerShell. 
py is not recognized 	Use python instead of py, or reinstall Python and ensure launcher/path options are selected. 
PowerShell opens Microsoft Store alias 	Disable App execution aliases for Python in Windows settings or install Python properly and verify PATH. 
ModuleNotFoundError 	Install the missing package using pip install <package_name>. 
Port already in use 	Change the Flask port in the app run command or code, for example app.run(port=5001). 
File upload fails 	Confirm the file is .xlsx or .csv, is not password protected, and is closed in Excel. 
Report output missing expected fields 	Check whether the selected fields exist in both source and Workday extracts and whether column names are aligned. 

11. Recommended Use in Workday Projects
•	Data migration validation
•	SIT validation
•	UAT validation
•	Data conversion cycles
•	Production data load validation
•	Post-load audit review



Best practice: Before final project use, run one small pilot reconciliation file first to confirm keys, field mapping, discrepancy classification, and report format are working as expected. 

