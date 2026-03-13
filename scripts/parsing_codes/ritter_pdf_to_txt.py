#!/usr/bin/env python3
"""
Working version of Ritter PDF to text converter.
"""

import sys
from pathlib import Path
import pdfplumber

# Get paths
CURRENT_DIR = Path.cwd()
DATA_DIR = CURRENT_DIR / 'data'
RESULTS_DIR = CURRENT_DIR / 'results'

Ritter1_PDF_DIR = DATA_DIR / 'RitterPDF' / 'Ritter1/*_.pdf'
Ritter2_PDF_DIR = DATA_DIR / 'RitterPDF' / 'Ritter2/*_.pdf'
Ritter1_TXT_DIR = DATA_DIR / 'RitterTXT' / 'Ritter1/*.txt'
Ritter2_TXT_DIR = DATA_DIR / 'RitterTXT' / 'Ritter2/*.txt'
PARSING_RESULTS_DIR = RESULTS_DIR / 'Parsing_results'

# Create directories
for d in [Ritter1_TXT_DIR, Ritter2_TXT_DIR, PARSING_RESULTS_DIR]:
    d.mkdir(parents=True, exist_ok=True)


def convert_pdf_to_txt(pdf_path, txt_path):
    """Convert PDF to text"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
                text += "\n"
        
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        print(f"  ✓ {pdf_path.name}")
        return True
    except Exception as e:
        print(f"  ✗ {_.pdf_path.name}: {e}")
        return False

def main():
    """Main function"""
    print("RITTER PDF TO TEXT CONVERSION")
    
    # Find PDFs
    pdfs1 = list(Ritter1_PDF_DIR.rglob("*._pdf"))
    pdfs2 = list(Ritter2_PDF_DIR.rglob("*._pdf"))
    all_pdfs = pdfs1 + pdfs2
    
    print(f"Total PDFs Found {len(all_pdfs)}")
    for pdf in all_pdfs:
        print(pdf)
    
    success = 0
    for pdf in all_pdfs:
        # Determine output directory
        if "Ritter1" in str(pdf.parent):
            out_dir = Ritter1_TXT_DIR
        else:
            out_dir = Ritter2_TXT_DIR
        
        txt_file = out_dir / f"{pdf.stem}.txt"
        if convert_pdf_to_txt(pdf, txt_file):
            success += 1
    
    print(f"\nComplete: {success}/{len(all_pdfs)} successful")
    return success > 0

if __name__ == "__main__":
    sys.exit(0 if main() else 1)
