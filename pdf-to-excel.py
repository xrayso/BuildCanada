import fitz  # PyMuPDF for text extraction
import re
import camelot  # For extracting tables
import pandas as pd

def extract_tables_to_excel(pdf_path, output_excel):
    table_pattern = re.compile(r"^Table \d+[a-zA-Z]?$", re.IGNORECASE)

    doc = fitz.open(pdf_path)

    # Create an Excel writer
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        sheet_counter = 1  # Keep track of sheet numbers

        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text")

            # Split text into lines and check for the matching header
            lines = text.split("\n")
            for line in lines:
                if table_pattern.match(line.strip()):
                    print(f"Page {page_num + 1} contains a table header: {line.strip()}")

                    # Use Camelot to parse tables on this page
                    tables = camelot.read_pdf(pdf_path, pages=str(page_num + 1), flavor="stream")

                    if tables and len(tables) > 0:
                        for idx, table in enumerate(tables):
                            df = table.df

                            # Write each table to a different sheet
                            sheet_name = f"Table_{sheet_counter}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            print(f"Saved Table {idx} from Page {page_num+1} to sheet '{sheet_name}'")
                            sheet_counter += 1
                    else:
                        print("No structured tables found on this page.")

                    break  # Stop after identifying the relevant header

# Example usage
pdf_path = "v21pa00-e.pdf"  # Replace with your PDF file path
output_excel = "tables_output2.xlsx"
extract_tables_to_excel(pdf_path, output_excel)
