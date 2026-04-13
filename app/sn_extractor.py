"""
sn_extractor.py
---------------
Reads the M&V Review Sheet Template at startup and returns the authoritative
list of question SNs from the "1. M&V Report Compliance Check" sheet.

Scans column B from row 24 onwards. Skips blank cells and whole-integer rows
(section headers). Returns SNs like ["0.1", "0.2"].
"""

from typing import List

import openpyxl


def _is_whole_integer(value) -> bool:
    """Return True if the value is a whole-integer section header or blank."""
    if value is None:
        return True
    s = str(value).strip()
    if not s:
        return True
    try:
        f = float(s)
        return f == int(f) and "." not in s
    except (ValueError, TypeError):
        return False


def extract_expected_sns(template_path: str) -> List[str]:
    """
    Open the Excel template and return the list of question SNs.

    Rules:
    - Scan column B from row 24 onwards in sheet "1. M&V Report Compliance Check".
    - Skip blank cells and whole-integer rows (section headers).

    Returns a list like ["0.1", "0.2"].
    """
    wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    ws = wb["1. M&V Report Compliance Check"]

    sns: List[str] = []
    for row in ws.iter_rows(min_row=24, min_col=2, max_col=2, values_only=True):
        value = row[0]
        if _is_whole_integer(value):
            continue
        sns.append(str(value).strip())

    wb.close()
    return sns
