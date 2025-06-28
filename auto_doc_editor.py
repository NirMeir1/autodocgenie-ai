"""Generate Word documents by filling placeholders using data from an Excel file."""
import os
import pandas as pd
from docx import Document

OUTPUT_DIR = "AutomaticDocEditor"

def replace_placeholders(doc, replacements):
    """Replace '____' placeholders in the document with provided values sequentially."""
    placeholder = "____"
    index = 0
    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            while placeholder in run.text and index < len(replacements):
                run.text = run.text.replace(placeholder, str(replacements[index]), 1)
                index += 1
    # Replace inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        while placeholder in run.text and index < len(replacements):
                            run.text = run.text.replace(placeholder, str(replacements[index]), 1)
                            index += 1
    return doc

def process_documents(excel_path, template_path):
    """Generate Word documents by replacing placeholders based on Excel rows."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.read_excel(excel_path)

    # Expect at least four columns: Business Name, Year, Field 3, and Field 4
    for _, row in df.iterrows():
        business_name = row[df.columns[0]]
        year = row[df.columns[1]]
        field3 = row[df.columns[2]]
        field4 = row[df.columns[3]]

        replacements = [business_name, year, field3, field4]

        doc = Document(template_path)
        replace_placeholders(doc, replacements)

        output_file = os.path.join(OUTPUT_DIR, f"{business_name}.docx")
        doc.save(output_file)
        print(f"Created: {output_file}")

if __name__ == "__main__":
    # Example usage; modify paths as needed
    EXCEL_FILE = "input.xlsx"
    TEMPLATE_FILE = "template.docx"

    process_documents(EXCEL_FILE, TEMPLATE_FILE)
