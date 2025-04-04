import pandas as pd
import os

try:
    import pdfplumber
except ModuleNotFoundError:
    pdfplumber = None
    print("Warning: pdfplumber is not installed. Please install it using 'pip install pdfplumber' before running this script.")

def extract_tables_from_pdf(pdf_path, output_excel):
    if not pdfplumber:
        print("Error: pdfplumber module not available. Cannot extract tables.")
        return

    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"Processing Page {page_num + 1}")
            page_tables = page.extract_tables()
            for table_index, table in enumerate(page_tables):
                if table:
                    cleaned_table = [[cell.strip() if isinstance(cell, str) else cell for cell in row] for row in table if any(cell is not None for cell in row)]
                    df = pd.DataFrame(cleaned_table)
                    tables.append((f"Page_{page_num + 1}_Table_{table_index + 1}", df))

    if not tables:
        print("No tables found.")
        return

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        for sheet_name, df in tables:
            sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    print(f"Extraction complete. Tables saved to {output_excel}")

if __name__ == "__main__":
    input_folder = "pdf_inputs"
    output_folder = "excel_outputs"

    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print(f"Input folder '{input_folder}' was missing and has been created. Please add PDFs to this folder and rerun the script.")
    else:
        os.makedirs(output_folder, exist_ok=True)

        for filename in os.listdir(input_folder):
            if filename.lower().endswith(".pdf"):
                input_path = os.path.join(input_folder, filename)
                output_filename = os.path.splitext(filename)[0] + ".xlsx"
                output_path = os.path.join(output_folder, output_filename)
                extract_tables_from_pdf(input_path, output_path)
