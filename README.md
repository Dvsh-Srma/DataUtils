# DataUtils

**DataUtils** provides a set of Python utilities for common data-handling tasks. In this repository you will find scripts to convert Markdown files into beautifully styled HTML and well-formatted DOCX documents.

## Repository Structure

```
DataUtils/
├── .gitignore                  # Specifies files and directories to be ignored by Git
├── README.md                   # This file provides an overview and instructions
├── requirements.txt            # Python dependencies required for the project
├── md2html.py                  # Script to convert Markdown to styled HTML
├── md2docx.py                  # Script to convert Markdown to DOCX with formatting
├── markdownresp.txt            # Example markdown input file (optional)
└── output.html / output.docx   # Generated after script execution (ignored by Git)
```

## Getting Started

### Prerequisites

- Python 3.6+
- (Optional) Virtual environment tools: `venv` or `virtualenv`

### Installation

1. **Clone the Repository**

   ```bash
   git clone https://github.com/Dvsh-Srma/DataUtils.git
   cd DataUtils
   ```

2. **Create and Activate a Virtual Environment**

   ```bash
   python -m venv venv
   source venv/bin/activate    # On Windows use: venv\Scripts\activate
   ```

3. **Install the Dependencies**

   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Convert Markdown to HTML

Run the conversion script with:

```bash
python md2html.py --input_markdown markdownresp.txt --html_output output.html
```

### Convert Markdown to DOCX

Run the conversion script with:

```bash
python md2docx.py -i markdownresp.txt -o output.docx
```

## Dependencies

This project uses the following libraries:

- `markdown`
- `beautifulsoup4`
- `html2docx`
- `python-docx`

All the dependencies are listed in the [requirements.txt](requirements.txt) file.

## Contributing

Feel free to submit issues or pull requests if you'd like to contribute!

## License

Distributed under the MIT License. See the [LICENSE](LICENSE) file for more information.