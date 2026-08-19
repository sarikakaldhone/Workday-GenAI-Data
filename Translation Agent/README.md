Translation Agent
Local Setup & Installation Guide
For: Anyone running the agent on a new machine

Overview
The Translation Agent is a self-contained Python web application that translates Excel workbooks into multiple languages. It runs entirely on your local machine — no cloud account, no server, and (optionally) no internet connection required.

Key facts:
•	Single Python file — copy it anywhere and run it
•	Browser-based UI powered by Streamlit
•	Three translation engines to choose from
•	Auto-installs all required Python libraries on first run

Prerequisites
1. Python 3.10 or higher
Python is the only thing you need to install manually. Everything else is handled automatically.

Step	Action
Download	Go to python.org/downloads and click Download Python 3.x
Install	Run the installer — IMPORTANT: check 'Add Python to PATH'
Verify	Open Command Prompt and run:   python --version

Important:  During installation, tick the checkbox 'Add Python to PATH'. Without this, the python command will not be recognised in the terminal.

2. Internet connection
Required for:
•	Downloading Python libraries on first run (one-time only)
•	Google Translate engine (every translation)
•	LibreTranslate — only for first-time model download
•	Argos Translate — only for first-time language model download (~100 MB per language)

Note:  Once libraries and models are downloaded, Argos Translate works completely offline. LibreTranslate also works offline once models are installed.

3. The agent file
You need one file:

translation_agent.py

Copy this file to any folder on your machine. That folder is your working directory.

Running the Application
Step 1 — Open a terminal
Open Command Prompt (Windows) or Terminal (Mac/Linux).

Navigate to the folder containing translation_agent.py:

cd C:\path\to\your\folder

Step 2 — Run the file
python translation_agent.py

On first run, the script will:
•	Detect any missing Python libraries
•	Automatically install them via pip (requires internet)
•	Launch the Streamlit web app
•	Open your browser at http://localhost:8501

Tip:  Subsequent runs are faster — libraries are already installed and the browser opens immediately.

Step 3 — Use the app
•	Select a translation engine from the radio buttons
•	Upload your .xlsx Excel file
•	Review the estimated time and row count
•	Click Translate
•	Download the translated file when complete

Python Libraries (Auto-installed)
These libraries are installed automatically the first time you run the agent. You do not need to install them manually.

Library	Version	Purpose
streamlit	1.35.0 or higher	Runs the browser-based web interface
openpyxl	3.1.2 or higher	Reads and writes Excel .xlsx files
deep-translator	1.11.4 or higher	Google Translate integration (no API key)
pandas	2.0.0 or higher	File preview table in the UI
requests	Latest	HTTP calls for LibreTranslate engine

Additional libraries installed on demand (only when selected):
•	argostranslate — installed automatically when Argos engine is selected
•	libretranslate — installed separately if using the self-hosted server option

Translation Engines
The agent supports three translation engines. Select the one that suits your data sensitivity needs.

Option 1 — Google Translate (default)
Data privacy	Text is sent to Google over HTTPS
Internet required	Yes — for every translation
Languages	130+
Extra setup	None
API key	Not required
Best for	Non-sensitive data: UI labels, system messages, field names

Option 2 — LibreTranslate (self-hosted, private)
Requires a local LibreTranslate server running on the same machine or network.

Data privacy	Fully private — data never leaves your network
Internet required	Only for initial model download
Languages	30 languages
RAM required	4 GB minimum, 8 GB recommended
Disk space	~300 MB per language pair
Best for	Sensitive data: PII, financials, client names, contracts

Setup commands:
pip install libretranslate
libretranslate --host 0.0.0.0 --port 5000

Keep the terminal window open while using the app. The agent will show a green tick when the server is detected.

Option 3 — Argos Translate (fully offline)
Pure Python library — no separate server needed. Works completely offline after initial setup.

Data privacy	Fully private — no internet needed after setup
Internet required	Only for first-time language model download
Languages	30 languages
RAM required	2 GB minimum
Disk space	~100 MB per language pair
Best for	Air-gapped machines or travel without reliable internet

Note:  Argos Translate downloads language models automatically the first time you translate to a new language. Subsequent runs use the cached model and work without internet.

Supported Languages
Code	Language	Code	Language	Google
ar	Arabic	ko	Korean	130+
az	Azerbaijani	fa	Persian	130+
zh	Chinese	pl	Polish	130+
cs	Czech	pt	Portuguese	130+
da	Danish	ru	Russian	130+
nl	Dutch	sk	Slovak	130+
fi	Finnish	es	Spanish	130+
fr	French	sv	Swedish	130+
de	German	tl	Tagalog	130+
el	Greek	tr	Turkish	130+
he	Hebrew	uk	Ukrainian	130+
hi	Hindi	vi	Vietnamese	130+
hu	Hungarian	it	Italian	130+
id	Indonesian	ga	Irish	130+
ja	Japanese			130+

Note:  Google Translate supports 130+ languages including regional variants (fr-CA, pt-BR, zh-TW). LibreTranslate and Argos Translate support the 30 languages listed above and map regional variants to the base language (e.g. fr-CA → fr).

Excel File Requirements
Your workbook must contain the following columns. Exact names are not required — the agent uses fuzzy matching to detect them.

Column	Example name	Description
User Language	User Language / Locale / Lang	BCP-47 locale code, e.g. fr-CA, de-DE, ja-JP
Base Value	Base Value / Source / English Value	English source text to translate
Rich Base Value	Rich Base Value / Rich Source	Extended or formatted English text (optional)
Translated Value	Translated Value / Translation	Populated automatically by the agent
Translated Rich Value	Translated Rich Value / Rich Translation	Populated automatically by the agent
Status	Status / Translation Status	Set to 'Done' by the agent after translation

Additional rules:
•	The header row can be anywhere in the first 20 rows — auto-detected
•	Blank User Language cells inherit the last filled value above (fill-down)
•	Placeholders like {0}, %s, and HTML tags are preserved untouched
•	System codes, IDs, and proper nouns are not translated
•	Blank source cells are left blank in the output

Troubleshooting
Problem	Solution
'python' not recognised	Python is not in PATH. Reinstall Python and tick 'Add to PATH'
Browser does not open	Manually open http://localhost:8501 in your browser
LibreTranslate not found	Start the server first: libretranslate --host 0.0.0.0 --port 5000
Translation takes too long	Large files with many unique strings take longer. Progress bar and remaining time are shown
Port 8501 already in use	Run: python translation_agent.py --server.port 8502
pip install fails	Check internet connection. On corporate networks, try: pip install --trusted-host pypi.org <package>

Quick Start Checklist
Checklist:  Complete these steps in order to get the agent running on a new machine.

•	Install Python 3.10 or higher from python.org — tick 'Add Python to PATH'
•	Copy translation_agent.py to any folder on your machine
•	Open Command Prompt and navigate to that folder
•	Run:   python translation_agent.py
•	Wait for libraries to install (first run only)
•	Browser opens at http://localhost:8501
•	Select your translation engine
•	Upload your .xlsx file and click Translate
•	Click Download when complete

Summary:  That's it. No cloud account, no API key, no database, no Docker. Just Python and one file.

