try:
    replacements = {"<Name>": "Zihad", "<ID>": "123"}
    body = "Hello <Name>, your ID is <ID>."
    
    # Current logic in core.py
    formatted = body.format(**replacements)
    print(f"Result: {formatted}")
except Exception as e:
    print(f"Error: {e}")

# Alternative logic I plan to check
body2 = "Hello <Name>."
for k, v in replacements.items():
    body2 = body2.replace(k, v)
print(f"Manual Replace: {body2}")
