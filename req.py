required_modules = [
    "pandas",
    "openpyxl",
    "requests"
]

with open('requirements.txt', 'w') as file:
    for module in required_modules:
        file.write(module + '\n')