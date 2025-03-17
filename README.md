# PDF Table Extractor

This Python script extracts tables from PDF bank statements and saves them as Excel files. It is designed to handle various bank statement formats, such as those from Punjab and Sind Bank and Standard Chartered, by leveraging spatial text analysis with `pdfminer.six`.

## Features
- Extracts text lines with coordinates from multi-page PDFs.
- Groups text into rows and detects column boundaries based on spatial alignment.
- Filters transaction rows using date patterns.
- Saves extracted tables to Excel files using pandas.
- Processes all PDFs in a specified folder.

## Prerequisites
- Python 3.6 or higher
- Required Python packages (see `requirements.txt`)

## Installation

1. **Clone the Repository** (or download the script):
   ```bash
   git clone <repository-url>
   cd pdf-table-extractor
