# Legislative Monitor – Argentina

Python pipeline for monitoring legislative initiatives related to freedom of expression, digital rights, and internet regulation in the Argentine Chamber of Deputies.

## Overview

This project automatically:

1. Scrapes newly introduced legislative projects from the Argentine Chamber of Deputies.
2. Downloads the associated PDF files.
3. Extracts text from each document.
4. Uses a Large Language Model to determine whether the project affects freedom of expression or digital regulation.
5. Generates a structured Excel dataset for monitoring legislative developments.

## Research Context

This tool was developed as part of research on digital regulation and freedom of expression in Latin America.

## Requirements


requests
beautifulsoup4
pandas
openpyxl
pymupdf
anthropic


## Usage

1. Create a config file:


config.json


Example:


{
"anthropic_api_key": "YOUR_API_KEY",
"llm_model": "claude-sonnet"
}


2. Run:


python monitor_arg.py


## Output

The script produces:

- Excel dataset of relevant projects
- downloaded PDFs
- log files

## License

MIT License
