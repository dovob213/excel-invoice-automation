from datetime import datetime, timedelta
import re

def parse_date_from_sheet_name(sheet_name):
    """
    Parses date from sheet name (e.g., '1.1', '12.25').
    Assumes current year or handles year logic if needed.
    Returns datetime object.
    """
    try:
        # User example: '1.1', '2.23'
        # Let's assume format is M.D
        parts = sheet_name.split('.')
        if len(parts) == 2:
            month, day = int(parts[0]), int(parts[1])
            current_year = datetime.now().year # Or user might want specific year?
            # User example file: "1-3월_xx하우스.xlsx". Catalog: "2026년..." 
            # Safest is to use current year or config. 
            # Given the prompt mentions "2026년_1분기...", let's try to detect year or default to current.
            # Using current execution time year (2026) is safe based on User Metadata.
            return datetime(2026, month, day)
    except Exception as e:
        print(f"Error parsing date from {sheet_name}: {e}")
        return None

def get_sending_date(receiving_date):
    """
    Returns receiving_date - 1 day.
    """
    if receiving_date:
        return receiving_date - timedelta(days=1)
    return None

import sys
import os

def get_resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def normalize_string(s):
    """
    Removes spaces, hidden characters, lowercases, and handles basic normalization.
    """
    if not s:
        return ""
    # Unicode normalize?
    import unicodedata
    s = str(s).lower()
    s = unicodedata.normalize('NFKC', s) # Normalize chars
    # Remove all whitespace
    return "".join(s.split())
