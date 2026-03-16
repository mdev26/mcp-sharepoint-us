import sys
sys.path.insert(0, 'src')

from mcp_sharepoint.docx_builder import create_word_document

# Test content
content = [
    {"type": "heading", "level": 1, "text": "Test Document"},
    {"type": "paragraph", "text": "This is a test paragraph.", "bold": True},
    {"type": "table", 
     "headers": ["Column 1", "Column 2"],
     "rows": [
       ["Data A", "Data B"],
       ["Data C", "Data D"]
     ]},
    {"type": "list", "style": "bullet", "items": ["Item 1", "Item 2", "Item 3"]}
]

# Create the document
doc_bytes = create_word_document(content)

# Save it locally to check
with open("test_output.docx", "wb") as f:
    f.write(doc_bytes)

print("✓ Document created: test_output.docx")
print("Open it to verify formatting!")
