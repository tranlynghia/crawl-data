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
        
    def _fetch_page(self, ticker: str, year: int) -> Dict[str, Any]:
        url = self.URL_TEMPLATE.format(ticker=ticker.lower(), year=year)
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
                
            years = []
            grid_rows = tbl_grid.find_all('tr')
            if not grid_rows:
                return {"status": "MISSING", "data": {}}
                
            headers = grid_rows[0].find_all(['th', 'td'])
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
                
            ppe_data = {}
            found_tscd = False
            for tr in rows[1:]:
                tds = tr.find_all('td')
                if not tds: continue
                
                name = tds[0].get_text(strip=True).lower()
                if "tài sản cố định hữu hình" in name:
                    found_tscd = True
                    
                if found_tscd and "nguyên giá" in name:
                    number_tds = [td for td in tds if 'number' in td.get('class', []) or 'number-item' in td.get('class', [])]
                    if not number_tds:
                        number_tds = [td for td in tds[1:] if not td.find(class_=re.compile('echarts|chart|growth', re.I))]
                        
                    for i, td in enumerate(number_tds):
                        if i < len(valid_years):
                            val = clean_number(td.get_text(strip=True))
                            if val is not None:
                                ppe_data[valid_years[i]] = val / 1_000_000
                    break
                    
            if ppe_data:
                return {"status": "OK", "data": ppe_data}
            return {"status": "MISSING", "data": {}}
        except Exception as e:
            print(f"CafeF PPE Error for {ticker}: {e}")
            return {"status": "ERROR", "data": {}}

    def get_ppe(self, ticker: str, start_year: int, end_year: int) -> Dict[str, Any]:
        """
        Lấy Nguyên giá TSCĐ hữu hình (PPE) từ Cân đối kế toán CafeF.
        Lặp qua nhiều trang nếu cần để lấy đủ từ start_year đến end_year.
        """
        all_data = {}
        current_year = end_year
        
        # Max loop to avoid infinite loop
        for _ in range(5):
            res = self._fetch_page(ticker, current_year)
            if res["status"] != "OK" or not res["data"]:
                break
                
            page_data = res["data"]
            all_data.update(page_data)
            
            min_year_fetched = min(page_data.keys())
            if min_year_fetched <= start_year:
                break
            
            # Step back
            current_year = min_year_fetched - 1
            if current_year < start_year:
                break
                
        if all_data:
            return {"status": "OK", "data": all_data}
        return {"status": "MISSING", "data": {}}
