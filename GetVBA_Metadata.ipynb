import os
import re
import csv

def analyze_vba_file(file_path):
    with open(file_path, 'r') as file:
        content = file.read()

    result = []

    pattern = re.compile(r'(Sub|Function)\s+(\w+)\s*(\([^)]*\))?\s*(\'[^\n]*|)', re.IGNORECASE)
    for match in pattern.finditer(content):
        procedure_type, procedure_name, params, summary = match.groups()
        referenced_objects = extract_referenced_objects(content, procedure_name)

        result.append({
            'file_name': os.path.basename(file_path),
            'procedure_name': procedure_name,
            'referenced_objects': ', '.join(referenced_objects),
            'summary': summary.strip()
        })

    return result

def extract_referenced_objects(content, procedure_name):
    procedure_pattern = re.compile(r'(Sub|Function)\s+' + procedure_name + r'\s*(\([^)]*\))?\s*(\'[^\n]*|)', re.IGNORECASE)
    procedure_match = procedure_pattern.search(content)
    if not procedure_match:
        return []

    start = procedure_match.end()
    end_pattern = re.compile(r'End (Sub|Function)', re.IGNORECASE)
    end_match = end_pattern.search(content, pos=start)
    end = end_match.start()

    procedure_content = content[start:end]

    # Add regex patterns for the objects you want to extract
    object_patterns = [
        re.compile(r'Set\s+(\w+)\s*=', re.IGNORECASE),
        re.compile(r'Dim\s+(\w+)\s+As\s+(\w+)', re.IGNORECASE)
    ]

    objects = set()
    for pattern in object_patterns:
        for match in pattern.finditer(procedure_content):
            objects.add(match.group(1))

    return list(objects)

def main():
    folder_path = "path/to/your/vba/files"
    output_csv = "output.csv"

    result = []

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".bas") or file_name.endswith(".frm"):
            file_path = os.path.join(folder_path, file_name)
            result.extend(analyze_vba_file(file_path))

    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['file_name', 'procedure_name', 'referenced_objects', 'summary']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for row in result:
            writer.writerow(row)

if __name__ == "__main__":
    main()
