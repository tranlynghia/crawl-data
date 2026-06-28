import re
from typing import List, Tuple

FINANCIAL_KEYWORDS = [
    "ngân hàng", "chứng khoán", "bảo hiểm", "quỹ đầu tư", "tài chính"
]

def clean_stock_symbols(input_text: str) -> List[str]:
    """
    Cleans stock symbols input text.
    Handles commas, spaces, newlines.
    Returns uppercase list of unique symbols.
    """
    if not input_text:
        return []
    
    # Replace commas with spaces, then split by whitespace
    raw_symbols = re.sub(r'[,;]', ' ', input_text).split()
    
    # Strip, uppercase, filter empty
    symbols = [s.strip().upper() for s in raw_symbols if s.strip()]
    
    # Remove non-alphanumeric (some stocks might have numbers like FPT, but just basic clean)
    symbols = [re.sub(r'[^A-Z0-9]', '', s) for s in symbols]
    symbols = [s for s in symbols if s]
    
    # Unique while preserving order
    seen = set()
    unique_symbols = []
    for s in symbols:
        if s not in seen:
            seen.add(s)
            unique_symbols.append(s)
            
    return unique_symbols

def is_financial_stock(industry_name: str) -> bool:
    """
    Checks if an industry name suggests it's a financial stock
    (Bank, Insurance, Securities, Fund).
    """
    if not industry_name:
        return False
        
    industry_lower = industry_name.lower()
    for kw in FINANCIAL_KEYWORDS:
        if kw in industry_lower:
            return True
    return False

def check_duplicates(new_symbols: List[str], existing_symbols: List[str]) -> Tuple[List[str], List[str]]:
    """
    Checks which symbols are already in the existing symbols list.
    Returns:
        - List of new symbols (not in existing)
        - List of duplicated symbols
    """
    existing_set = set([s.upper() for s in existing_symbols])
    new_only = []
    dups = []
    for s in new_symbols:
        if s in existing_set:
            dups.append(s)
        else:
            new_only.append(s)
    return new_only, dups
