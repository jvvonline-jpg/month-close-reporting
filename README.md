# Month-End Close Agent

**Cornerstones, Inc. | Finance Department**

An intelligent Streamlit application for automated month-end close diagnostics. Processes QuickBooks GL exports and generates six advanced analytical reports with Apple-branded Word document output.

## Repository Structure

```
16.06 Month-End Close App/
├── app.py                      # Main Streamlit application
├── requirements.txt            # Python dependencies
├── .gitignore
├── .streamlit/
│   └── config.toml             # Streamlit theme (Apple color palette)
└── README.md
```

## Setup

### Prerequisites

- Python 3.9+
- Node.js 18+ (for Word document export via docx-js)
- npm

### Installation

```bash
# Clone the repository
git clone <your-repo-url>
cd "16.06 Month-End Close App"

# Install Python dependencies
pip install -r requirements.txt

# Install Node.js docx package (for Word export)
npm install -g docx

# Run the application
streamlit run app.py
```

### Deploying to Streamlit Cloud

1. Push this folder to a GitHub repository.
2. Go to [share.streamlit.io](https://share.streamlit.io) and connect your repo.
3. Set the main file path to `app.py`.
4. Note: Word document export requires Node.js. For Streamlit Cloud, the IIF export and on-screen reports work without Node.js.

## Reports

| # | Report | Purpose |
|---|--------|---------|
| 1 | Flux Narrative | Month-over-month variance analysis with AI-generated explanations |
| 2 | Missing Bill | Identifies recurring vendors absent from the current period |
| 3 | Suspense Reclass | Flags misc/suspense transactions with COA-based reclassification suggestions |
| 4 | Materiality | Isolates transactions above a configurable dollar threshold |
| 5 | IIF Pre-Flight | Validates journal entries before QuickBooks import |
| 6 | Reconciliation | Three-way match between GL, bank data, and PDF invoices |

## Input Files

- **General Ledger**: CSV or Excel (.xlsx) export from QuickBooks Desktop
- **Chart of Accounts**: IIF or CSV file with account names and classes
- **PDF Invoices/Statements**: Optional, for multi-source reconciliation

## Export Formats

- **Word Documents (.docx)**: Apple-branded reports with Cornerstones headers/footers
- **IIF File**: Tab-delimited QuickBooks import file with TRNS/SPL/ENDTRNS structure
