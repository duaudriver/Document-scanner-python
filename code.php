## document_scanner.py ##
import os
import re
import zipfile
import fitz
import pandas as pd
import phonenumbers
import chardet
import logging
from docx import Document
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import spacy

# Setup Logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Load spaCy model
nlp = spacy.load("en_core_web_sm")

# Regex patterns
EMAIL_REGEX = re.compile(r'\b[\w\.-]+@[\w\.-]+\.\w+\b')
CREDIT_CARD_REGEX = re.compile(r'\b(?:\d[ -]*?){13,16}\b')
TFN_REGEX = re.compile(r'\b\d{8,9}\b')
MEDICARE_REGEX = re.compile(r'\b\d{4}\s?\d{5}\s?\d\b')
CRN_REGEX = re.compile(r'\b[A-Z]\d{8}\b')

# Supportde file extensions
SUPPORTED_EXTENSIONS = {'.txt', '.docx', '.pdf', '.xlsx', '.zip'}

# Normalise text encoding
def normalise_text(text):
    try:
        return text.encode('utf-8', errors='ignore').decode('utf-8')
    except Exception:
        return text

#file extraction functions
def extract_text_from_txt(file_path):
    with open(file_path, 'rb') as f:
        raw = f.read()
        encoding = chardet.detect(raw)['encoding']
        return raw.decode(encoding or 'utf-8', errors='ignore')

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    return normalise_text('\n'.join(para.text for para in doc.paragraphs))

def extract_text_from_pdf(file_path):
    text = ''
    with fitz.open(file_path) as doc:
        for page in doc:
            text += page.get_text()
    return normalise_text(text)

def extract_text_from_xlsx(file_path):
    text = ''
    try:
        wb = load_workbook(file_path, data_only=True)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        text += f"{cell.value}\n"
        return normalise_text(text)
    except Exception as e:
        logging.error(f"Error reading excel file {file_path}: {e}")
        return ''

# TFN Checksum
def is_valid_tfn(tfn):
    weights_9 = [1, 4, 3, 7, 5, 8, 6, 9, 10]
    weights_8 = [10, 7, 8, 4, 6, 3, 5, 1]
    digits = [int(d) for d in tfn]
    if len(digits) == 9:
        total = sum(d * w for d, w in zip(digits, weights_9))
    elif len(digits) == 8:
        total = sum(d * w for d, w in zip(digits, weights_8))
    else:
        return False
    return total % 11 == 0

# Sensitive information extraction
def extract_sensitive_info(text):
    emails = EMAIL_REGEX.findall(text)
    credit_cards = CREDIT_CARD_REGEX.findall(text)
    raw_tfns = TFN_REGEX.findall(text)
    tfns = [tfn for tfn in raw_tfns if is_valid_tfn(tfn)]
    medicare = MEDICARE_REGEX.findall(text)
    crns = CRN_REGEX.findall(text)

    phone_numbers = [
        phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
        for match in phonenumbers.PhoneNumberMatcher(text, "AU")
    ]

    doc = nlp(text)
    names = [ent.text for ent in doc.ents if ent.label_ == 'PERSON']

    return {
        'emails': list(set(emails)),
        'phone_numbers': list(set(phone_numbers)),
        'credit_cards': list(set(credit_cards)),
        'tax_file_numbers': list(set(tfns)),
        'medicare_numbers': list(set(medicare)),
        'centrelink_crns': list(set(crns)),
        'names': list(set(names))
    }

# File processor
def process_file(file_path, results, base_path=''):
    ext = os.path.splitext(file_path)[1].lower()
    full_path = os.path.join(base_path, file_path) if base_path else file_path
    try:
        if ext == '.txt':
            text = extract_text_from_txt(full_path)
        elif ext == '.docx':
            text = extract_text_from_docx(full_path)
        elif ext == '.pdf':
            text = extract_text_from_pdf(full_path)
        elif ext == '.xlsx':
            text = extract_text_from_xlsx(full_path)
        elif ext == '.zip':
            with zipfile.ZipFile(full_path, 'r') as zip_ref:
                extract_dir = os.path.join(base_path, f"{os.path.splitext(os.path.basename(file_path))[0]}_extracted")
                os.makedirs(extract_dir, exist_ok=True)
                zip_ref.extractall(extract_dir)
                for root, _, files in os.walk(extract_dir):
                    for f in files:
                        if os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS:
                            logging.info(f"Processing extracted file: {f}")
                            process_file(f, results, base_path=root)
                        else:
                            logging.warning(f"Skipped unsupported file: {f}")
            return
        else:
            logging.warning(f"Unsupported file type: {file_path}")
            return
        
        results[full_path] = extract_sensitive_info(text)
    except Exception as e:
        logging.error(f"Failed to process {full_path}: {e}")
        results[full_path] = {'error': str(e)}

# Directory scanner
def scan_uploaded_files(upload_dir):
    results = {}
    for root, _, files in os.walk(upload_dir):
        for file in files:
            if os.path.splitext(file)[1].lower() in SUPPORTED_EXTENSIONS:
                process_file(file, results, base_path=root)
    return results

# Save results to a formatted excel
def save_results_to_excel(results, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Scan Results"
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = "Sensitive Data Scan Report"
    title_cell.font = Font(bold=True, size=24)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40
    headers = ['File', 'Emails', 'Phone Numbers', 'Names', 'Credit Cards', 'Tax File Numbers', 'Medicare Numbers', 'Centrelink CRNs']
    ws.append(headers)

    # Header format
    header_font = Font(bold=True, size=18)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=2, column=col)
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 30

    
    # Rows
    for file, info in results.items():
        if 'error' in info:
            row = [file] + ['Error'] * (len(headers) - 1)
        else:
            row = [
                file,
                normalise_text('\n'.join(info['emails'])),
                normalise_text('\n'.join(info['phone_numbers'])),
                normalise_text('\n'.join(info['names'])),
                normalise_text('\n'.join(info['credit_cards'])),
                normalise_text('\n'.join(info['tax_file_numbers'])),
                normalise_text('\n'.join(info['medicare_numbers'])),
                normalise_text('\n'.join(info['centrelink_crns']))
            ]
        ws.append(row)
    
    # Formatting of cells
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    fill_gray = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

    for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        max_lines = max(str(cell.value).count('\n') + 1 if cell.value else 1 for cell in row)
        ws.row_dimensions[row[0].row].height = (max_lines * 20) + 5 if max_lines > 1 else 30

        for cell in row:
            if cell.row == 2:
                cell.font = Font(bold=True, size=18)
            else:
                cell.font = Font(size=18)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
            if idx % 2 == 0 and cell.row > 2:
                cell.fill = fill_gray
    
    # Column widths
    MAX_WIDTH = 50
    MIN_WIDTH = 20
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = max(min(max_length + 2, MAX_WIDTH), MIN_WIDTH)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    custom_widths = {
        'F': 40,
        'G': 40,
        'H': 40
    }
    for col_letter, width in custom_widths.items():
        ws.column_dimensions[col_letter].width = width
    
    # Freeze header row
    ws.freeze_panes = "A3"

    wb.save(output_file)

# Main execution
if __name__ == "__main__":
    upload_directory = "./uploads"
    scan_results = scan_uploaded_files(upload_directory)
    save_results_to_excel(scan_results, "scan_results.xlsx")
    print("Scan results saved to scan_results.csv")
