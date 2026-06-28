import re
import requests
import pandas as pd
from io import StringIO
import unicodedata
from typing import Dict, List, Optional
import time

COMPANY_JSON_URL = "https://cafefnew.mediacdn.vn/Search/company.json"
OLD_REPORT_URL = (
    "https://cafef.vn/du-lieu/bao-cao-tai-chinh/"
    "{ticker}/{statement}/{year}/{quarter}/0/{type}/0/{slug}.chn"
)

STATEMENTS = {
    "income": ("incsta", "ket-qua-hoat-dong-kinh-doanh"),
    "balance": ("bsheet", "can-doi-ke-toan"),
    "cashflow": ("cashflow", "luu-chuyen-tien-te-tructiep"),
    "cashflow_indirect": ("cashflowin", "luu-chuyen-tien-te-giantiep")
}

def create_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
    })
    return session

def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"^\s*[-+()]?\s*[A-Za-z]?\d+([.)])?\s*", "", text)
    text = re.sub(r"^\s*[-+]\s*", "", text)
    text = re.sub(r"\s+", " ", text.lower()).strip()
    return text

def parse_number(value: object, unit_multiplier: float) -> Optional[float]:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan" or text == "-":
        return None
    negative = text.startswith("(") and text.endswith(")")
    digits = re.sub(r"[^\d-]", "", text)
    if not digits or digits == "-":
        return None
    number = float(digits) * unit_multiplier
    return -abs(number) if negative else number

def fetch_html(session, url: str) -> Optional[str]:
    try:
        response = session.get(url, timeout=30)
        response.raise_for_status()
        return response.text
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return None

def parse_table_html(html: str, target_years: List[int], unit_multiplier: float) -> Dict[int, Dict[str, float]]:
    try:
        tables = pd.read_html(StringIO(html))
    except ValueError:
        return {}
    
    if len(tables) < 5:
        return {}

    header = tables[3]
    data = tables[4]
    
    # Try to find years in header
    years = []
    for col in range(1, min(5, header.shape[1])):
        val = str(header.iat[0, col])
        if val.isdigit() and len(val) == 4:
            years.append(int(val))
    
    if not years:
        # Fallback if parsing header failed, sometimes CafeF changes format
        return {}

    rows = {y: {} for y in years}
    asset_context = ""
    for _, source_row in data.iterrows():
        label_raw = source_row.iloc[0]
        if pd.isna(label_raw):
            continue
        label = normalize_text(label_raw)
        
        # Keep context for generic names like "nguyen gia"
        if label in {
            "tai san co dinh huu hinh",
            "tai san co dinh thue tai chinh",
            "tai san co dinh vo hinh",
            "bat dong san dau tu",
        }:
            asset_context = label
            
        compound_label = f"{asset_context} {label}" if label in ("nguyen gia", "gia tri hao mon luy ke") and asset_context else label
        
        for offset, year in enumerate(years, start=1):
            if offset < len(source_row):
                val = parse_number(source_row.iloc[offset], unit_multiplier)
                if val is not None:
                    rows[year][compound_label] = val
                    
    return rows

def crawl_company_data(
    session: requests.Session, 
    ticker: str, 
    start_year: int, 
    end_year: int, 
    is_consolidated: bool = True,
    unit_multiplier: float = 1.0
) -> Dict[str, Dict[int, Dict[str, float]]]:
    """
    Crawls data for a ticker over a range of years.
    Returns dict:
    {
        "income": {2021: {"doanh thu": ...}, ...},
        "balance": {...},
        "cashflow": {...}
    }
    """
    results = {"income": {}, "balance": {}, "cashflow": {}}
    report_type = "0" if is_consolidated else "1"
    
    # We will try to fetch year by year, but CafeF usually returns 4 years per request
    # So we step backwards from end_year in chunks of 4.
    
    current_end = end_year
    while current_end >= start_year:
        for stmt_key in STATEMENTS:
            statement, slug = STATEMENTS[stmt_key]
            url = OLD_REPORT_URL.format(
                ticker=ticker.lower(),
                statement=statement,
                year=current_end,
                quarter=0,
                type=report_type,
                slug=slug
            )
            
            html = fetch_html(session, url)
            if html:
                parsed_rows = parse_table_html(html, list(range(start_year, end_year+1)), unit_multiplier)
                for y, row_data in parsed_rows.items():
                    if y not in results[stmt_key]:
                        results[stmt_key][y] = {}
                    results[stmt_key][y].update(row_data)
            time.sleep(1) # Polite crawling
        
        current_end -= 4
        
    return results
