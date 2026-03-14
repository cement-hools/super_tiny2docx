
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