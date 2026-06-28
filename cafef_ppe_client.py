import requests
from bs4 import BeautifulSoup
import re
from typing import Dict, Any, Optional

def clean_number(text: str) -> Optional[float]:
    if not text: return None
    s = text.strip()
    if not s or s.lower() in ['nan', '-']: return None
    
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative = True
        s = s[1:-1]
        
    # Remove all non-digits, commas, dots, and minus
    cleaned = re.sub(r'[^\d,.\-]', '', s)
    if not cleaned or cleaned == '-' or cleaned == '.': return None
    
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(",", "")
    elif cleaned.count(",") > 0:
        cleaned = cleaned.replace(",", "")
    elif cleaned.count(".") > 1 or (cleaned.count(".") == 1 and len(cleaned.rsplit(".", 1)[1]) == 3):
        cleaned = cleaned.replace(".", "")
        
    try:
        num = float(cleaned)
        return -abs(num) if negative else num
    except ValueError:
        return None

class CafeFPPEClient:
    URL_TEMPLATE = "https://cafef.vn/du-lieu/bao-cao-tai-chinh/{ticker}/bsheet/{year}/0/0/0/can-doi-ke-toan.chn"
    
    def __init__(self):
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        
    def get_ppe(self, ticker: str, end_year: int) -> Dict[str, Any]:
        """
        Lấy Nguyên giá TSCĐ hữu hình (PPE) từ Cân đối kế toán CafeF.
        Trả về dict: { "status": "OK" / "MISSING" / "ERROR", "data": { year_int: ppe_value_in_millions } }
        """
        url = self.URL_TEMPLATE.format(ticker=ticker.lower(), year=end_year)
        try:
            res = requests.get(url, headers=self.headers, timeout=10)
            if res.status_code != 200:
                return {"status": "ERROR", "data": {}}
                
            html = res.text
            if "Không tìm thấy dữ liệu" in html:
                return {"status": "MISSING", "data": {}}
                
            soup = BeautifulSoup(html, 'html.parser')
            table = soup.find('table', id='tableContent')
            if not table:
                return {"status": "MISSING", "data": {}}
                
            tbl_grid = soup.find('table', id='tblGridData')
            if not tbl_grid:
                return {"status": "MISSING", "data": {}}
                
            rows = table.find_all('tr')
            if not rows:
                return {"status": "MISSING", "data": {}}
                
            # Parse Header for years from tblGridData
            years = []
            grid_rows = tbl_grid.find_all('tr')
            if not grid_rows:
                return {"status": "MISSING", "data": {}}
                
            headers = grid_rows[0].find_all(['th', 'td'])
            
            # Note: the first column in headers usually contains 'Trước/Sau' navigation
            for th in headers[1:]:
                text = th.get_text(strip=True)
                year_match = re.search(r'\b(20\d{2})\b', text)
                if year_match:
                    years.append(int(year_match.group(1)))
                else:
                    years.append(None)
                    
            valid_years = [y for y in years if y is not None]
            if not valid_years:
                return {"status": "MISSING", "data": {}}
                
            # Find PPE row
            ppe_data = {}
            found_tscd = False
            for tr in rows[1:]:
                tds = tr.find_all('td')
                if not tds: continue
                
                name = tds[0].get_text(strip=True).lower()
                
                if "tài sản cố định hữu hình" in name:
                    found_tscd = True
                    
                if found_tscd and "nguyên giá" in name:
                    # Extract values
                    # CafeF uses 'number' or 'number-item' class
                    number_tds = [td for td in tds if 'number' in td.get('class', []) or 'number-item' in td.get('class', [])]
                    
                    if not number_tds:
                        # Fallback
                        number_tds = [td for td in tds[1:] if not td.find(class_=re.compile('echarts|chart|growth', re.I))]
                        
                    for i, td in enumerate(number_tds):
                        if i < len(valid_years):
                            val = clean_number(td.get_text(strip=True))
                            if val is not None:
                                # Normalize from VND to Million VND
                                ppe_data[valid_years[i]] = val / 1_000_000
                    break # Stop after finding PPE
                    
            if ppe_data:
                return {"status": "OK", "data": ppe_data}
            return {"status": "MISSING", "data": {}}
            
        except Exception as e:
            print(f"CafeF PPE Error for {ticker}: {e}")
            return {"status": "ERROR", "data": {}}
