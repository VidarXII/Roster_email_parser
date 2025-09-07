# Roster_email_parser

📧 Roster Email → Structured Excel Extractor

This project automates the process of parsing .eml emails (provider roster updates) and extracting structured information into a standardized Excel template using an LLM (Large Language Model).

It is designed for healthcare roster management hackathons where you receive unstructured email updates about providers and must normalize them into a fixed schema.

🚀 Features

Extracts provider and transaction details (add/update/term, NPIs, TINs, specialties, addresses, etc.) from unstructured roster emails.

Uses Microsoft Phi-3-mini-4k-instruct for fast, small-footprint LLM extraction.

Handles plain-text and HTML emails.

Normalizes missing data to "Information not found".

Exports structured results directly into a provided Excel template.

Supports both single file and batch processing.

📂 Project Structure
.
├── extractor.py        # Main script
├── requirements.txt    # Dependencies
├── README.md           # Project documentation
└── /samples            # Example .eml files (optional)

🛠️ Requirements

Python 3.9+

PyTorch
 (CPU or GPU)

Hugging Face transformers

openpyxl for Excel manipulation

beautifulsoup4 + lxml for HTML email parsing

Install dependencies:

pip install -r requirements.txt


Contents of requirements.txt:

transformers
torch
openpyxl
beautifulsoup4
lxml
pandas

⚡ Usage
1. Single .eml file
python extractor.py \
    /path/to/email.eml \
    /path/to/template.xlsx \
    /path/to/output.xlsx

2. Folder of .eml files
python extractor.py \
    /path/to/email/folder \
    /path/to/template.xlsx \
    /path/to/output.xlsx \
    -b 5 -v


-b sets batch size (default = 1).

-v enables verbose logging.

📑 Output Format

The extracted Excel follows the template headers:

Transaction Type (Add/Update/Term)	Transaction Attribute	Effective Date	Term Date	Term Reason	Provider Name	Provider NPI	Provider Specialty	State License	Organization Name	TIN	Group NPI	Complete Address	Phone Number	Fax Number	PPG ID	Line Of Business (Medicare/Commercial/Medical)

Missing fields will be filled with:

"Information not found"

🧠 How It Works

Parse Email
Extracts plain text + cleaned HTML text from .eml using email and BeautifulSoup.

Build JSON Prompt
Constructs a strict schema prompt for the LLM.

LLM Extraction
Phi-3-mini is used to generate structured JSON output.

Excel Mapping
JSON keys are mapped to template headers and appended row-wise.

🔍 Example

Input email snippet:

Please terminate Dr. John Smith, NPI 1234567890, effective 09/01/2024. 
Term reason: Retired. Group NPI: 9876543210.


Output row in Excel:

Term | Provider | 09/01/2024 | 09/01/2024 | Retired | John Smith | 1234567890 | Information not found | Information not found | Information not found | Information not found | 9876543210 | Information not found | Information not found | Information not found | Information not found | Information not found

⚠️ Notes

Ensure your Excel template has correct headers (they map directly to JSON keys).

LLM performance may vary; post-processing fills missing values with "Information not found".

For hackathon speed: run on GPU if available, otherwise CPU works (slower).

📜 License

MIT License – free to use, modify, and distribute.
