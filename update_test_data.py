import pandas as pd

# Define user's test data
data = [
    {'Name': 'Zihad Connects', 'ID': 'TEST-001', 'Department': 'QA Team', 'Email': 'zihad.connects@gmail.com'},
    {'Name': 'Zihad Hasan', 'ID': 'TEST-002', 'Department': 'Dev Team', 'Email': 'zihadhasan.sbmc1124@gmail.com'}
]

df = pd.DataFrame(data)
df.to_excel(r"C:\Users\X000\Downloads\Test\template_data.xlsx", index=False)
print("Updated test data with user emails.")
