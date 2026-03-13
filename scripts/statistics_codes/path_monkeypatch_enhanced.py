#!/usr/bin/env python3
"""
Enhanced monkey-patch that handles multiple path patterns.
"""

import sys
import os
import builtins
import pandas as pd
from pathlib import Path
import re

# Store originals
_original_open = builtins.open
_original_pd_read_csv = pd.read_csv

# Multiple possible original bases (add all that your scripts use)
ORIGINAL_BASES = [
    Path("/home/Ben/medical-pipeline"),
    Path("/home/Ben/sterilizers_project"),  # If some scripts still reference this
    Path("/home/user/old_path"),  # Add any others
]

# Detect environment
if 'google.colab' in sys.modules:
    NEW_BASE = Path("/content/medical-pipeline")
else:
    NEW_BASE = Path.cwd()

def translate_path(path):
    """Translate any hardcoded path to current environment"""
    if isinstance(path, Path):
        path_str = str(path)
    else:
        path_str = str(path)
    
    # If it's already a valid path in current environment, return as-is
    if Path(path_str).exists():
        return Path(path_str)
    
    # Try each original base
    for original_base in ORIGINAL_BASES:
        if path_str.startswith(str(original_base)):
            rel_path = Path(path_str).relative_to(original_base)
            translated = NEW_BASE / rel_path
            return translated
    
    # Pattern-based translation for any /home/.../medical-pipeline/ path
    path_str = re.sub(r'^/home/[^/]+/medical-pipeline/', str(NEW_BASE) + '/', path_str)
    path_str = re.sub(r'^/home/[^/]+/sterilizers_project/', str(NEW_BASE) + '/', path_str)
    
    return Path(path_str)

# Monkey-patch functions
def patched_open(file, *args, **kwargs):
    translated = translate_path(file)
    return _original_open(translated, *args, **kwargs)

def patched_read_csv(filepath, *args, **kwargs):
    translated = translate_path(filepath)
    return _original_pd_read_csv(translated, *args, **kwargs)

# Apply patches
builtins.open = patched_open
pd.read_csv = patched_read_csv

print(f"Enhanced monkey-patch activated")
print(f"  New base directory: {NEW_BASE}")
print(f"  Handling paths from: {[str(b) for b in ORIGINAL_BASES]}")

# Common path variables
DATA_DIR = NEW_BASE / 'data'
RESULTS_DIR = NEW_BASE / 'results'
SCRIPTS_DIR = NEW_BASE / 'scripts'

# Create directories
(NEW_BASE / 'results' / 'Parsing_results').mkdir(parents=True, exist_ok=True)
(NEW_BASE / 'results' / 'analysis_results').mkdir(parents=True, exist_ok=True)
