#!/usr/bin/env python3
"""
Simple pipeline runner.
"""

import sys
import subprocess
from pathlib import Path

# Get current directory
CURRENT_DIR = Path.cwd()

# Define all paths directly
SCRIPTS_DIR = CURRENT_DIR / 'scripts'
PARSING_CODES_DIR = SCRIPTS_DIR / 'parsing_codes'
STATISTICS_CODES_DIR = SCRIPTS_DIR / 'statistics_codes'
RESULTS_DIR = CURRENT_DIR / 'results'
PARSING_RESULTS_DIR = RESULTS_DIR / 'Parsing_results'

print(f"Pipeline directory: {CURRENT_DIR}")

def run_script(script_path, description):
    """Run a script"""
    print(f"\n▶ {description}")
    
    if not script_path.exists():
        print(f"  ✗ Not found: {script_path.name}")
        return False
    
    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0:
            print(f"  ✓ Success")
            return True
        else:
            print(f"  ✗ Failed (code: {result.returncode})")
            if result.stderr:
                print(f"    Error: {result.stderr[:100]}...")
            return False
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False

# Create results directory
PARSING_RESULTS_DIR.mkdir(parents=True, exist_ok=True)

# Run scripts
scripts = [
    (PARSING_CODES_DIR / 'ritter_pdf_to_txt.py', "Ritter PDF to Text"),
    (PARSING_CODES_DIR / 'ritter_txt_to_csv.py', "Ritter Text to CSV"),
    (PARSING_CODES_DIR / 'statim_parser.py', "Statim Parsing"),
]

# Add statistics scripts if they exist
for script_name, description in [('merged2.py', 'Analysis'), ('report2.py', 'Report')]:
    script_path = STATISTICS_CODES_DIR / script_name
    if script_path.exists():
        scripts.append((script_path, description))

results = []
for script_path, description in scripts:
    success = run_script(script_path, description)
    results.append((description, success))

# Summary
print("\n" + "="*60)
print("SUMMARY")
print("="*60)

success_count = sum(1 for _, success in results if success)
print(f"Success: {success_count}/{len(results)}")

# Check results
csv_files = list(PARSING_RESULTS_DIR.glob("*.csv"))
if csv_files:
    print(f"\nGenerated {len(csv_files)} CSV file(s):")
    for csv in csv_files:
        print(f"  • {csv.name}")

sys.exit(0 if success_count == len(results) else 1)
