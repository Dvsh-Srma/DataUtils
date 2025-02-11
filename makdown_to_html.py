
import argparse
import sys
import textwrap
import markdown
from bs4 import BeautifulSoup
from html2docx import html2docx

def convert_markdown_to_html(markdown_text, prettify=True):
    try:
        raw_html = markdown.markdown(
            markdown_text,
            extensions=['pymdownx.superfences', 'tables', 'codehilite', 'pymdownx.emoji'],
            extension_configs={
                'pymdownx.emoji': {
                    # This simple lambda wraps emojis in a span with class "emoji"
                    'emoji_generator': lambda emoji, options, md: f'<span class="emoji">{emoji}</span>',
                },
                'codehilite': {
                    'guess_lang': False,
                    'css_class': 'codehilite'
                }
            }
        )
        if prettify:
            soup = BeautifulSoup(raw_html, "html.parser")
            formatted_html = soup.prettify()
            return formatted_html
        return raw_html
    except Exception as e:
        raise RuntimeError(f"Failed converting Markdown to HTML: {e}")

def wrap_html_with_style(html_content):
    css = """
    <style>
      body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: #f8f9fa;
        color: #333;
        padding: 20px;
        line-height: 1.6;
      }
      h1, h2, h3, h4, h5, h6 {
        color: #0a0a23;
        margin-bottom: 15px;
      }
      p {
        margin-bottom: 15px;
      }
      code {
        background: #f4f4f4;
        padding: 2px 4px;
        border-radius: 4px;
      }
      pre {
        background: #f4f4f4;
        padding: 10px;
        overflow: auto;
        border: 1px solid #ddd;
        border-radius: 4px;
      }
      .codehilite {
        background: #f4f4f4;
        border: 1px solid #ddd;
        border-radius: 4px;
        overflow: auto;
        margin-bottom: 15px;
        padding: 10px;
      }
      table {
        border-collapse: collapse;
        width: 100%;
        margin-bottom: 20px;
      }
      th, td {
        border: 1px solid #ddd;
        padding: 8px;
        text-align: left;
      }
      th {
        background-color: #f2f2f2;
      }
      .emoji {
        font-size: 1.2em;
      }
    </style>
    """
    html_header = f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>Converted Markdown</title>{css}</head><body>"
    html_footer = "</body></html>"
    return html_header + html_content + html_footer

def save_html(html_content, output_html_file):
    try:
        with open(output_html_file, "w", encoding="utf-8") as f:
            f.write(html_content)
        print(f"HTML output saved to: {output_html_file}")
    except Exception as e:
        raise RuntimeError(f"Failed saving HTML to {output_html_file}: {e}")

def main():
    parser = argparse.ArgumentParser(
        description="Convert Markdown into beautifully styled HTML with support for code, tables, emoji, and more."
    )
    parser.add_argument("--input_markdown", type=str,
                        help="Path to the input markdown file. If not provided, a default sample is used.")
    parser.add_argument("--html_output", type=str, default="output.html",
                        help="Path for HTML output file (default: output.html)")
    parser.add_argument("--docx_output", type=str, default="output.docx",
                        help="Path for DOCX output file (default: output.docx)")
    
    args = parser.parse_args()

    # Load markdown text from file or use a default sample.
    if args.input_markdown:
        try:
            with open(args.input_markdown, "r", encoding="utf-8") as f:
                markdown_text = f.read()
        except Exception as e:
            print(f"Error reading input markdown file: {e}")
            sys.exit(1)
    else:
        try:
            with open("markdownresp.txt", "r", encoding="utf-8") as f:
                markdown_text = f.read()
                #markdown_text = textwrap.dedent(markdown_text)# markdown_text.replace("```", "```python")
        except Exception as e:
            print(f"Error reading default markdown file: {e}")
            sys.exit(1)

    # Step 1: Convert Markdown to HTML.
    try:
        raw_html = convert_markdown_to_html(markdown_text)
        # Wrap the raw HTML with a full HTML structure including CSS.
        full_html = wrap_html_with_style(raw_html)
        save_html(full_html, args.html_output)
    except Exception as e:
        print(f"Error during Markdown to HTML conversion: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()