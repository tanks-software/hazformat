import json

# Replace 'hazformat-56ea389dba51.json' with your actual filename
with open('hazformat-56ea389dba51.json', 'r') as f:
    data = json.load(f)

# Replace actual newlines in private_key with escaped newlines
data['private_key'] = data['private_key'].replace('\n', '\\n')

# Print pretty JSON with escaped newlines in private_key
print(json.dumps(data, indent=2))
