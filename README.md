# Resume ATS Formatter for Autofill

A Python script that converts ALL CAPS words in Microsoft Word (.docx) files to Title Case (as well as other fixes) to prepare resumes for autofill functionality in common Applicant Tracking Systems (ATS).

## Features

- Converts ALL CAPS words (5+ characters) to Title Case
- Preserves formatting and removes brackets after ALL CAPS words
- Handles multi-word ALL CAPS sequences
- Supports custom exceptions via .exceptions file

## Installation

1. Clone or download this repository
2. Create virtual environment: `python3 -m venv .venv`
3. Activate virtual environment: `source .venv/bin/activate`
4. Install dependencies: `pip install -r requirements.txt`

## Usage

### Basic Usage

```bash
# Process a single file (creates ATS_filename.docx)
python resume_ATS_fmt.py input.docx

# Process with custom output filename
python resume_ATS_fmt.py input.docx -o output.docx

# Process with custom exceptions file
python resume_ATS_fmt.py input.docx -e my_exceptions.txt
```

### Advanced Usage

```bash
# Add specific exceptions via command line
python resume_ATS_fmt.py input.docx --add-exception "JAVASCRIPT" --add-exception "MACOS"

# Combine multiple options
python resume_ATS_fmt.py input.docx -o processed.docx -e exceptions.txt --add-exception "GITHUB"
```

## Exceptions File

Create a `.exceptions` file in the project directory to define custom word formatting:

```
# Format: EXCEPTION=FORMATTED
JAVASCRIPT=JavaScript
MACOS=macOS
IOS=iOS
GITHUB=GitHub
LINKEDIN=LinkedIn

# Format: EXCEPTION (uses title case)
PYTHON
```

## Command Line Options

- `input_file`: Path to input .docx file (required)
- `-o, --output`: Output file path (default: ATS_filename.docx)
- `-e, --exceptions`: Path to exceptions file (default: .exceptions)
- `--add-exception`: Add specific exception word (can be used multiple times)
- `-h, --help`: Show help message
