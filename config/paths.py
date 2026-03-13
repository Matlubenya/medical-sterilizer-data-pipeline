import sys
from pathlib import Path

# Add config to Python path
config_path = Path(__file__).resolve().parents[2] / 'config'
sys.path.insert(0, str(config_path))

from paths import *

"""
Path configuration for medical data pipeline.
Matches EXACT structure you provided.
"""

from pathlib import Path

# Main pipeline directory - where this config file is located
# config/paths.py → parent = config → parent = medical-pipeline
MEDICAL_PIPELINE_DIR = Path(__file__).resolve().parent.parent

# ==================== DATA PATHS ====================
DATA_DIR = MEDICAL_PIPELINE_DIR / 'data'

# Ritter PDF data (raw PDF files)
RITTER_PDF_DIR = DATA_DIR / 'RitterPDF'
RITTER1_PDF_DIR = RITTER_PDF_DIR / 'Ritter1'
RITTER2_PDF_DIR = RITTER_PDF_DIR / 'Ritter2'

# Ritter TXT data (intermediate text files)
RITTER_TXT_DIR = DATA_DIR / 'RitterTXT'
RITTER1_TXT_DIR = RITTER_TXT_DIR / 'Ritter1'
RITTER2_TXT_DIR = RITTER_TXT_DIR / 'Ritter2'

# Statim data (raw text files)
STATIM_DIR = DATA_DIR / 'Statim'
STATIMA_DIR = STATIM_DIR / 'StatimA'
STATIMB_DIR = STATIM_DIR / 'StatimB'

# ==================== RESULTS PATHS ====================
RESULTS_DIR = MEDICAL_PIPELINE_DIR / 'results'

# Parsing results
PARSING_RESULTS_DIR = RESULTS_DIR / 'Parsing_results'
STATIM_CSV = PARSING_RESULTS_DIR / 'statim.csv'
RITTER_COMBINED_CSV = PARSING_RESULTS_DIR / 'ritter_combined_summary.csv'

# Analysis results
ANALYSIS_RESULTS_DIR = RESULTS_DIR / 'analysis_results'
JSON_RESULTS_DIR = ANALYSIS_RESULTS_DIR / 'json'
NUMERICAL_RESULTS_DIR = ANALYSIS_RESULTS_DIR / 'numerical'
VISUAL_RESULTS_DIR = ANALYSIS_RESULTS_DIR / 'visual'
COMPLETE_ANALYSIS_PKL = ANALYSIS_RESULTS_DIR / 'complete_analysis.pkl'

# ==================== SCRIPTS PATHS ====================
SCRIPTS_DIR = MEDICAL_PIPELINE_DIR / 'scripts'

# Parsing codes
PARSING_CODES_DIR = SCRIPTS_DIR / 'parsing_codes'
RITTER_PDF_TO_TXT = PARSING_CODES_DIR / 'ritter_pdf_to_txt.py'
RITTER_TXT_TO_CSV = PARSING_CODES_DIR / 'ritter_txt_to_csv.py'
STATIM_PARSER = PARSING_CODES_DIR / 'statim_parser.py'

# Statistics codes
STATISTICS_CODES_DIR = SCRIPTS_DIR / 'statistics_codes'
MERGED2_PY = STATISTICS_CODES_DIR / 'merged2.py'
REPORT2_PY = STATISTICS_CODES_DIR / 'report2.py'

# ==================== OTHER FILES ====================
REQUIREMENTS_TXT = MEDICAL_PIPELINE_DIR / 'requirements.txt'
RUN_PIPELINE_PY = MEDICAL_PIPELINE_DIR / 'run_pipeline.py'
SETUP_COLAB_PY = MEDICAL_PIPELINE_DIR / 'setup_colab.py'
README_MD = MEDICAL_PIPELINE_DIR / 'README.md'

# ==================== CREATE DIRECTORIES ====================
def create_directories():
    """Create all necessary directories if they don't exist."""
    dirs_to_create = [
        RITTER_TXT_DIR,  # For intermediate text files
        RITTER1_TXT_DIR,
        RITTER2_TXT_DIR,
        PARSING_RESULTS_DIR,
        ANALYSIS_RESULTS_DIR,
        JSON_RESULTS_DIR,
        NUMERICAL_RESULTS_DIR,
        VISUAL_RESULTS_DIR,
    ]
    
    for directory in dirs_to_create:
        directory.mkdir(parents=True, exist_ok=True)

# Create directories on import
create_directories()

# ==================== VERIFICATION ====================
if __name__ == "__main__":
    print("=" * 60)
    print("MEDICAL PIPELINE PATHS CONFIGURATION")
    print("=" * 60)
    print(f"Pipeline directory: {MEDICAL_PIPELINE_DIR}")
    print(f"✓ Data directory exists: {DATA_DIR.exists()}")
    print(f"✓ Scripts directory exists: {SCRIPTS_DIR.exists()}")
    print(f"✓ Results directory exists: {RESULTS_DIR.exists()}")
    print("=" * 60)
