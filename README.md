# Super Tiny2Docx

[![PyPI](https://img.shields.io/pypi/v/super_tiny2docx)](https://pypi.org/project/super_tiny2docx/)
[![Python Version](https://img.shields.io/pypi/pyversions/super_tiny2docx)](https://pypi.org/project/super_tiny2docx/)
[![License: MIT](https://img.shields.io/github/license/cement-hools/super_tiny2docx)](https://github.com/cement-hools/super_tiny2docx/blob/main/LICENSE)

Convert HTML from TinyMCE (or other WYSIWYG editors) into .docx documents easily. Designed for developers looking to integrate rich text export functionality into their Python applications. Supports tables, lists, styles, and more.

**Contributions welcome!**

## ✨ Features

- Convert HTML to `.docx` with accurate formatting.
- Supports paragraphs, headings, lists (ordered & unordered), tables, and inline styles.
- Handles CSS-like styling (font, size, color, alignment, margins, borders, etc.).
- Proper style inheritance for nested elements.
- Preserves spaces and line breaks.
- Extensible architecture for adding new features.

## 📦 Installation

```bash
pip install super_tiny2docx
```

## 🚀 Quick Start

```python
from super_tiny2docx import SuperTiny2Docx

# Example HTML content
html_content = """
<h1>Welcome to Super Tiny2Docx</h1>
<p>This is a <strong>bold</strong> paragraph with <em>emphasis</em>.</p>
<ol>
    <li>First ordered list item</li>
    <li>Second ordered list item</li>
</ol>
<table border="1">
    <tr>
        <th>Header 1</th>
        <th>Header 2</th>
    </tr>
    <tr>
        <td>Data 1</td>
        <td>Data 2</td>
    </tr>
</table>
"""

# Convert to DOCX
converter = SuperTiny2Docx(html_content)
docx_bytes = converter.convert()

# Save to file
with open("output.docx", "wb") as f:
    f.write(docx_bytes.read())

print("Document saved as output.docx")
```

## 🛠️ Advanced Usage
### Super Tiny2Docx automatically parses inline styles and <style> tags to apply formatting. 
You can customize the default styles by extending the ComputedStyle class.


## Contributing
We welcome contributions! 
Whether it's bug reports, feature suggestions, code improvements, 
or documentation fixes — your help is appreciated.
Fork the repository.
Create a new branch for your feature or bug fix.
Add your changes and write tests if applicable.
Submit a pull request.

## 🙏 Acknowledgements
Thanks to the community for support and inspiration. 
Special thanks to the maintainers of `python-docx` and `beautifulsoup4`.


