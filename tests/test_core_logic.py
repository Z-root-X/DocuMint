import pytest

def test_manual_string_replacement():
    """Verify that manual string replacement works as expected."""
    replacements = {"<Name>": "Zihad", "<ID>": "123"}
    body = "Hello <Name>, your ID is <ID>."
    
    final_body = body
    for k, v in replacements.items():
        final_body = final_body.replace(k, v)
        
    assert final_body == "Hello Zihad, your ID is 123."

def test_missing_placeholder():
    """Verify behavior when a placeholder is missing in replacements."""
    replacements = {"<Name>": "Zihad"}
    body = "Hello <Name>, your ID is <ID>."
    
    final_body = body
    for k, v in replacements.items():
        final_body = final_body.replace(k, v)
        
    # <ID> should remain untouched
    assert final_body == "Hello Zihad, your ID is <ID>."
