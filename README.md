# DataUtils

**DataUtils** provides a set of Python utilities for common data-handling tasks. In this repository you will find scripts to convert Markdown files into beautifully styled HTML and well-formatted DOCX documents.

## Repository Structure

```
DataUtils/
├── .gitignore                  # Specifies files and directories to be ignored by Git
├── LICENSE                     # MIT License
├── README.md                   # This file provides an overview and instructions
├── pyproject.toml             # Modern Python project configuration and dependencies
├── md2html.py                 # Script to convert Markdown to styled HTML
├── md2docx.py                 # Script to convert Markdown to DOCX with formatting
└── markdownresp.txt           # Example markdown input file
```

## Features

- Convert Markdown to beautifully styled HTML with support for:
  - Syntax highlighted code blocks
  - Tables
  - Emojis
  - Custom styling
- Convert Markdown to well-formatted DOCX documents with:
  - Proper heading hierarchy
  - Code formatting
  - Hyperlinks
  - Lists and tables

## Getting Started

### Prerequisites

- Python 3.8+
- [uv](https://github.com/astral-sh/uv) - Fast Python package installer and resolver
  ```bash
  # Install uv using pip
  pip install uv
  # Or on Windows using the installer
  winget install astral.uv
  ```

### Installation

1. **Clone the Repository**

   ```bash
   git clone https://github.com/Dvsh-Srma/DataUtils.git
   cd DataUtils
   ```

2. **Create and Activate a Virtual Environment**

   ```bash
   uv venv
   .venv/Scripts/activate    # On Windows
   # source .venv/bin/activate    # On Unix/MacOS
   ```

3. **Install the Dependencies**

   ```bash
   uv add markdown beautifulsoup4 html2docx python-docx pymdown-extensions
   ```

## Usage

### Convert Markdown to HTML

The script will use `markdownresp.txt` by default if no input file is specified:

```bash
python md2html.py
```

Or specify input and output files:

```bash
python md2html.py --input_markdown input.md --html_output result.html
```

### Convert Markdown to DOCX

Use default input/output files:

```bash
python md2docx.py
```

Or specify custom files:

```bash
python md2docx.py -i input.md -o result.docx
```

## Dependencies

This project uses modern Python packaging with pyproject.toml and the following main dependencies:

- `markdown>=3.3.0` - For Markdown parsing and conversion
- `beautifulsoup4>=4.9.0` - For HTML processing and formatting
- `html2docx` - For HTML to DOCX conversion
- `python-docx>=0.8.11` - For creating and manipulating DOCX files
- `pymdown-extensions>=10.3` - For enhanced Markdown features including emoji support

## Contributing

Feel free to submit issues or pull requests if you'd like to contribute!

## License

Distributed under the MIT License. See the LICENSE file for more information.