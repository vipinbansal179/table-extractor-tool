from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextLine
import pandas as pd
from collections import Counter
import os
import glob

from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBox, LTTextLine

def extract_text_lines(pdf_path):
    text_lines = []
    for page_layout in extract_pages(pdf_path):
        for element in page_layout:
            if isinstance(element, LTTextBox):
                for text_line in element:
                    if isinstance(text_line, LTTextLine):
                        text_lines.append({
                            'text': text_line.get_text().strip(),
                            'x0': text_line.x0,
                            'y0': text_line.y0,
                            'x1': text_line.x1,
                            'y1': text_line.y1
                        })
    return text_lines

def group_into_rows(text_lines, y_threshold=10):
    """Group text lines into rows based on vertical (y0) proximity."""
    text_lines.sort(key=lambda x: x['y0'], reverse=True)
    rows = []
    current_row = [text_lines[0]]
    
    for line in text_lines[1:]:
        if abs(line['y0'] - current_row[-1]['y0']) <= y_threshold:
            current_row.append(line)
        else:
            rows.append(current_row)
            current_row = [line]
    rows.append(current_row)
    return rows

def detect_column_boundaries(rows, x_threshold=5):
    """Detect column boundaries based on frequent x0 and x1 positions."""
    x_positions = []
    for row in rows:
        for line in row:
            x_positions.append(round(line['x0'], 1))
            x_positions.append(round(line['x1'], 1))
    
    counter = Counter(x_positions)
    frequent_xs = [x for x, count in counter.items() if count > 1]
    frequent_xs.sort()
    
    columns = []
    for i in range(len(frequent_xs) - 1):
        columns.append((frequent_xs[i], frequent_xs[i + 1]))
    return columns

def assign_to_cells(rows, columns):
    """Assign text lines to table cells based on column boundaries."""
    table = []
    for row in rows:
        row_cells = [''] * len(columns)
        for line in row:
            for i, (col_start, col_end) in enumerate(columns):
                if (line['x0'] <= col_end and line['x1'] >= col_start):
                    if not row_cells[i]:
                        row_cells[i] = line['text']
                    else:
                        row_cells[i] += ' ' + line['text']
                    break
        table.append(row_cells)
    return table

def extract_table_to_excel(pdf_path, output_excel_path):
    """Extract a table from a PDF and save it to an Excel file."""
    text_lines = extract_text_lines(pdf_path)
    if not text_lines:
        print(f"No text lines found in {pdf_path}.")
        return
    
    rows = group_into_rows(text_lines)
    columns = detect_column_boundaries(rows)
    if not columns:
        print(f"Could not detect column boundaries in {pdf_path}.")
        return
    
    table = assign_to_cells(rows, columns)
    df = pd.DataFrame(table, columns=[f"Col{i+1}" for i in range(len(columns))])
    df.to_excel(output_excel_path, index=False)
    print(f"Table extracted from {pdf_path} and saved to {output_excel_path}")

def process_pdfs_in_folder(input_folder, output_folder):
    """Process all PDFs in the input folder and save tables as Excel files in the output folder."""
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Get all PDF files from the input folder
    pdf_files = glob.glob(os.path.join(input_folder, "*.pdf"))
    
    if not pdf_files:
        print(f"No PDF files found in {input_folder}.")
        return
    
    # Process each PDF file
    for pdf_path in pdf_files:
        # Generate the output Excel file path using the PDF's base name
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_excel_path = os.path.join(output_folder, f"{base_name}.xlsx")
        
        # Extract the table and save it to Excel
        extract_table_to_excel(pdf_path, output_excel_path)

# Example usage
if __name__ == "__main__":
    input_folder = "sample_pdfs"  # Replace with your input folder path
    output_folder = "output_excels"  # Replace with your output folder path
    process_pdfs_in_folder(input_folder, output_folder)