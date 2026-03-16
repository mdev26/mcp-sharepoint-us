import sys
sys.path.insert(0, 'src')

from mcp_sharepoint.docx_builder import create_word_document

def test_create_simple_document():
    content = [
        {"type": "heading", "level": 1, "text": "Test"},
        {"type": "paragraph", "text": "Hello World"}
    ]
    
    doc_bytes = create_word_document(content)
    
    assert doc_bytes is not None
    assert len(doc_bytes) > 0
    print("✓ Simple document test passed")

def test_create_complex_document():
    content = [
        {"type": "heading", "level": 1, "text": "Monthly Report"},
        {"type": "paragraph", "text": "Summary", "bold": True},
        {"type": "table", 
         "headers": ["Task", "Status"],
         "rows": [["Task 1", "Done"], ["Task 2", "Pending"]]},
        {"type": "list", "style": "bullet", "items": ["Item 1", "Item 2"]}
    ]
    
    doc_bytes = create_word_document(content)
    
    assert doc_bytes is not None
    assert len(doc_bytes) > 0
    print("✓ Complex document test passed")

if __name__ == "__main__":
    test_create_simple_document()
    test_create_complex_document()
    print("\n✓ All tests passed!")
