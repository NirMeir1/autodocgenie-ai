# autodocgenie-ai

This project generates Word documents by filling a template with data from
an Excel spreadsheet. Each row of the spreadsheet results in a new document
named after the business in that row.

## Requirements

- Python 3.11+
- [python-docx](https://python-docx.readthedocs.io/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- pandas (for Excel helpers)

You can install the libraries with:

```bash
python -m pip install python-docx openpyxl pandas
```

## Usage

1. **Prepare your files**
   - Excel file (`.xlsx`) with columns: Business Name, Year, Field 3, Field 4.
   - Word template (`.docx`) containing `____` placeholders to be replaced.

2. **Run the script**

   ```bash
   python auto_doc_editor.py <excel_path> <template_path> [--workers N]
   ```

   - `excel_path` – path to the Excel file.
   - `template_path` – path to the Word template.
   - `--workers` – optional number of parallel workers. Defaults to the CPU
     count if omitted.

3. **Results**
   - Output documents are created inside the `AutomaticDocEditor` directory.
   - Each document is named after the corresponding Business Name.

Example:

```bash
python auto_doc_editor.py data.xlsx template.docx --workers 4
```

This will populate `template.docx` with the data from `data.xlsx` and place
the resulting files in the `AutomaticDocEditor` folder.

