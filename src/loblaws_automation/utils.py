"""
Utility functions for date parsing, filename extraction, and openxml decoding.
"""

import re
from datetime import datetime, timedelta


def excel_serial_to_iso(serial_str: str) -> str | None:
    """
    Convert an Excel serial date number to an ISO date string (YYYY-MM-DD).

    Excel epoch is 1899-12-30 (accounting for the Lotus 1-2-3 Feb 29 1900 bug).
    Returns None if conversion fails.
    """
    try:
        serial = int(float(serial_str))
        if serial < 1 or serial > 200000:
            return None
        epoch = datetime(1899, 12, 30)
        dt = epoch + timedelta(days=serial)
        return dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError, OverflowError):
        return None


def extract_date_from_filename(source_name: str) -> str | None:
    """
    Extract a date from a workbook filename like LOBLAWS_NAC2020094821_090826.xlsx.

    The trailing 6-digit group is MMDDYY format.
    """
    match = re.search(r'_(\d{6})\.xlsx', source_name, re.IGNORECASE)
    if match:
        mmddyy = match.group(1)
        try:
            mm, dd, yy = mmddyy[:2], mmddyy[2:4], mmddyy[4:6]
            year = int(yy)
            full_year = 2000 + year if year < 80 else 1900 + year
            dt = datetime(full_year, int(mm), int(dd))
            return dt.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            return None
    return None
