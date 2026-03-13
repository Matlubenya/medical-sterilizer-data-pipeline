#!/usr/bin/env python3
"""
Working version of Statim parser.
"""

import sys
from pathlib import Path
import pandas as pd

# Get paths
CURRENT_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = CURRENT_DIR / 'data'
RESULTS_DIR = CURRENT_DIR / 'results'

STATIMA_DIR = DATA_DIR / 'Statim' / 'StatimA'
STATIMB_DIR = DATA_DIR / 'Statim' / 'StatimB'
PARSING_RESULTS_DIR = RESULTS_DIR / 'Parsing_results'

# Create directory
PARSING_RESULTS_DIR.mkdir(parents=True, exist_ok=True)

def parse_statim_file(txt_path):
    """Parse Statim file - ADD YOUR PARSING LOGIC HERE"""
    try:
        with open(txt_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Your parsing logic here
        data = {
            'file': txt_path.name,
            'source': 'StatimA' if 'StatimA' in str(txt_path.parent) else 'StatimB',
            'size': len(content),
            'lines': content.count('\n') + 1
        }
        
        return data
    except Exception as e:
        print(f"  Error: {txt_path.name} - {e}")
        return None

def main():
    """Main function"""
    print("STATIM PARSING")
    
    # Find files
    filesA = list(STATIMA_DIR.glob("*.txt"))
    filesB = list(STATIMB_DIR.glob("*.txt"))
    all_files = filesA + filesB
    
    print(f"Found {len(all_files)} text files")
    
    # Parse files
    all_data = []
    for txt in all_files:
        data = parse_statim_file(txt)
        if data:
            all_data.append(data)
    
    # Save to CSV
    if all_data:
        df = pd.DataFrame(all_data)
        output_file = PARSING_RESULTS_DIR / "statim.csv"
        df.to_csv(output_file, index=False)
        print(f"\nSaved {len(df)} records to: {output_file}")
        return True
    else:
        print("\nNo data extracted")
        return False

if __name__ == "__main__":
    sys.exit(0 if main() else 1)
