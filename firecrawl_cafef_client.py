import time
import re
from firecrawl import FirecrawlApp
from schemas import financial_statement_schema, EXTRACT_PROMPT

OLD_REPORT_URL = (
    "https://cafef.vn/du-lieu/bao-cao-tai-chinh/"
    "{ticker}/{statement}/{year}/0/{type}/0/{slug}.chn"
)

STATEMENTS = {
    "income": ("incsta", "ket-qua-hoat-dong-kinh-doanh"),
    "balance": ("bsheet", "can-doi-ke-toan"),
    "cashflow": ("cashflow", "luu-chuyen-tien-te-tructiep"),
    "cashflow_indirect": ("cashflowin", "luu-chuyen-tien-te-giantiep")
}

def clean_number(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    # Convert string to float
    s = str(val).strip()
    if not s or s.lower() == 'nan' or s == '-':
        return None
    # Handle negative in parenthesis
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative = True
        s = s[1:-1]
    
    # Remove all non-digits except dot and minus
    s = re.sub(r'[^\d.-]', '', s)
    try:
        if not s or s == '-' or s == '.':
            return None
        num = float(s)
        return -abs(num) if negative else num
    except ValueError:
        return None

def scrape_company_data_firecrawl(ticker: str, start_year: int, end_year: int, is_consolidated: bool, api_key: str, unit_multiplier: float = 1.0):
    app = FirecrawlApp(api_key=api_key)
    results = {"income": {}, "balance": {}, "cashflow": {}, "cashflow_indirect": {}}
    report_type = "0" if is_consolidated else "1"
    
    current_end = end_year
    while current_end >= start_year:
        for stmt_key, (statement, slug) in STATEMENTS.items():
            url = OLD_REPORT_URL.format(
                ticker=ticker.lower(),
                statement=statement,
                year=current_end,
                type=report_type,
                slug=slug
            )
            
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    print(f"Scraping {url} via Firecrawl (attempt {attempt+1})...")
                    res = app.scrape_url(url, params={
                        'formats': ['extract'],
                        'extract': {
                            'schema': financial_statement_schema,
                            'prompt': EXTRACT_PROMPT
                        }
                    })
                    
                    if res and 'extract' in res and res['extract']:
                        extracted_data = res['extract']
                        for row in extracted_data.get('rows', []):
                            item_name = row.get('item_name')
                            if not item_name: continue
                            
                            item_name_lower = str(item_name).lower()
                            
                            values = row.get('values', {})
                            for y_str, val in values.items():
                                try:
                                    y = int(y_str)
                                    num_val = clean_number(val)
                                    if num_val is not None:
                                        if y not in results[stmt_key]:
                                            results[stmt_key][y] = {}
                                        results[stmt_key][y][item_name_lower] = num_val * unit_multiplier
                                except ValueError:
                                    pass
                    break # Success
                except Exception as e:
                    print(f"Error scraping {url}: {e}")
                    if attempt < max_retries - 1:
                        time.sleep(3)
                    else:
                        print(f"Failed to scrape {url} after {max_retries} attempts.")
            time.sleep(2) # Polite sleep
            
        current_end -= 4
        
    return results
