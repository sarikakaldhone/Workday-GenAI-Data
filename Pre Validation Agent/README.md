DGW Excel Data Pre-Validation Agent
Setup Guide
1. Prerequisites
Ensure the following are available:
•	Microsoft Excel-compatible .xlsx files
•	Python 3.x
•	Python package for Excel processing:
pip install openpyxl
•	Read access to the input location
•	Write access to the output location
Input workbooks must not be password protected.
________________________________________
2. Create the Folder Structure
Create the following folders:
DGW_Validation/
├── input/
├── output/
└── run_validator.py
•	input: Stores the two source workbooks.
•	output: Stores the generated validation report.
•	run_validator.py: Contains the validation agent implementation.
________________________________________
3. Add the Input Workbooks
Place the following two files in the input folder:
input/
├── DGW_Input.xlsx
└── Validation_Rules.xlsx
Only two input workbooks are required for a run:
1.	DGW workbook containing the DGW data sheets.
2.	Validation Rules workbook containing validation rules and Reference ID sections.
The implementation should accept file paths dynamically and must not depend on these example file names.
________________________________________
4. Prepare the DGW Workbook
Each DGW data sheet should contain:
•	A row with recognizable business field names
•	At least one populated business record below the header
•	Meaningful business columns
The header may appear below template or metadata rows.
Example:
Row 1: Section label
Row 2: Integration mapping
Row 3: Data type
Row 4: Required/Optional indicator
Row 5: Business field names
Row 6 onward: Business data
In this example, row 5 is the actual header.
The workbook may also contain instruction, walkthrough, reference, code-list, DQ-rule, blank, or metadata sheets. These do not need to be removed because the agent should detect and ignore them dynamically.
________________________________________

5. Prepare the Validation Rules Workbook
Include validation rule sections with identifiable labels such as:
Validation Rule
Validation Rules
Required Fields
Each validation section should include:
•	Section or process name
•	Required fields
•	Rule text
•	Business field names referenced by the rules
Use clear rule descriptions, for example:
Worker ID is required.
Worker ID must be unique.
Hire Date Format must be YYYY-MM-DD.
Annual Salary must be numeric.
Annual Salary cannot be zero.
Do not depend on exact validation sheet names. The agent should discover and match rule sections dynamically.
________________________________________
6. Prepare the Reference ID Sections
Reference sections should contain:
•	Reference ID in the section title or sheet name
•	One header row with reference field names
•	Allowed values listed below the corresponding headers


Example:
Worker Type	Location ID	Pay Rate Type
Employee	BLR01	HOURLY
Contingent Worker	HYD01	SALARIED
Intern	MUM01	COMMISSION
Keep allowed values field-specific.
Do not include the following within the allowed-value rows:
•	Data types
•	Required/Optional indicators
•	Instructions
•	Notes
•	Repeated headers
If metadata is required, place it above the reference header row.
________________________________________
7. Configure the Input and Output Paths
The validator should receive three paths:
DGW workbook path
Validation Rules workbook path
Output report path
Example:
python run_validator.py \
  --dgw "input/DGW_Input.xlsx" \
  --rules "input/Validation_Rules.xlsx" \
  --output "output/DGW_Validation_Report.xlsx"
The file paths, workbook names, sheet names, process names, and column names must not be hardcoded.
________________________________________
8. Run the Validator
From the DGW_Validation folder, execute:
python run_validator.py \
  --dgw "input/DGW_Input.xlsx" \
  --rules "input/Validation_Rules.xlsx" \
  --output "output/DGW_Validation_Report.xlsx"
Before execution:
•	Close the input workbooks in Excel if the environment restricts access to open files.
•	Close any existing output report with the same name.
•	Confirm that the output folder is writable.
•	Confirm that both input files open without Excel repair warnings.
________________________________________
9. Verify the Setup
Confirm the following before production execution:
•	Python 3.x is installed.
•	openpyxl is installed.
•	The DGW workbook is in the input folder.
•	The Validation Rules workbook is in the input folder.
•	The DGW workbook contains recognizable business headers.
•	Intended DGW sheets contain at least one data row.
•	Validation sections contain clear rule text.
•	Reference sections contain field-specific allowed values.
•	Input paths are passed dynamically.
•	The output folder exists and is writable.
•	The output workbook is not already open.
•	No process, sheet, or column names are hardcoded.

