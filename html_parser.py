import re
from bs4 import BeautifulSoup
from typing import List, Dict, Any
from fireant_mapper import parse_financial_number

def clean_text(text: str) -> str:
    if not text: return ""
    return re.sub(r'\s+', ' ', text).strip()

class CafeFParser:
    @staticmethod
    def parse(html_content: str) -> List[Dict[str, Any]]:
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Thử tìm bảng KQKD, CDKT của giao diện mới hoặc cũ
        table = soup.find('table', id='tableContent') or soup.find('table', id='table_KQKD') or soup.find('table')
        if not table:
            return []
            
        rows = table.find_all('tr')
        if not rows:
            return []
            
        # Lấy năm từ Header
        years = []
        headers = rows[0].find_all(['th', 'td'])
        for th in headers:
            text = clean_text(th.get_text())
            # Regex tìm số năm (ví dụ: 2021, 2022)
            year_match = re.search(r'\b(20\d{2})\b', text)
            if year_match:
                years.append(int(year_match.group(1)))
            else:
                years.append(None)
                
        # Nếu ko tìm thấy năm trong header, thử fallback dựa trên số cột
        valid_years = [y for y in years if y is not None]
        
        result = []
        stack = [] # Dùng stack để lưu các node cha theo level
        
        for tr in rows[1:]:
            classes = tr.get('class', [])
            tds = tr.find_all('td')
            if not tds: continue
            
            # Tên chỉ tiêu
            name = clean_text(tds[0].get_text())
            if not name: continue
            
            # Xử lý level từ class của tr (CafeF thường có level1, level2, level3...)
            level = 1
            for cls in classes:
                if cls.startswith('level'):
                    try:
                        level = int(cls.replace('level', ''))
                    except: pass
            
            # Nếu không có class level, thử tìm qua thuộc tính treegrid
            # ví dụ treegrid-1, treegrid-parent-1
            tree_classes = [c for c in classes if c.startswith('treegrid-')]
            is_parent = False
            for c in tree_classes:
                if c.startswith('treegrid-parent-'):
                    pass
                elif c.startswith('treegrid-'):
                    is_parent = True
            
            # Lấy số liệu từ các cột
            values = {}
            # Các cột chứa số của CafeF mới thường có class "number" hoặc "number-item"
            number_tds = [td for td in tds if 'number' in td.get('class', []) or 'number-item' in td.get('class', [])]
            
            if number_tds:
                for i, td in enumerate(number_tds):
                    if i < len(valid_years):
                        values[str(valid_years[i])] = parse_financial_number(td.get_text())
            else:
                # Nếu không có class number, fallback lấy tất cả td trừ cột đầu tiên
                val_tds = [td for td in tds[1:] if not td.find(class_=re.compile('echarts|chart|growth', re.I))]
                for i, td in enumerate(val_tds):
                    if i < len(valid_years):
                        values[str(valid_years[i])] = parse_financial_number(td.get_text())
            
            node = {
                "name": name,
                "values": values,
                "rows": [],
                "_level": level
            }
            
            # Xây dựng cây
            while stack and stack[-1]['_level'] >= level:
                stack.pop()
                
            if stack:
                stack[-1]['rows'].append(node)
            else:
                result.append(node)
                
            stack.append(node)
            
        return result

class FireAntParser:
    @staticmethod
    def parse(html_content: str) -> List[Dict[str, Any]]:
        soup = BeautifulSoup(html_content, 'html.parser')
        table = soup.find('table')
        if not table:
            return []
            
        rows = table.find_all('tr')
        if not rows:
            return []
            
        years = []
        headers = rows[0].find_all(['th', 'td'])
        for th in headers:
            text = clean_text(th.get_text())
            year_match = re.search(r'\b(20\d{2})\b', text)
            if year_match:
                years.append(int(year_match.group(1)))
            else:
                years.append(None)
                
        valid_years = [y for y in years if y is not None]
        
        result = []
        stack = []
        
        for tr in rows[1:]:
            tds = tr.find_all('td')
            if not tds: continue
            
            name = clean_text(tds[0].get_text())
            if not name: continue
            
            # FireAnt: Xác định cấp bậc dựa trên CSS padding-left
            # style="padding-left: 29px;"
            level = 1
            style = tds[0].get('style', '')
            padding_match = re.search(r'padding-left:\s*(\d+)px', style)
            if padding_match:
                padding = int(padding_match.group(1))
                # Giả sử padding cơ bản là 8px, mỗi cấp cộng 20px
                # 8px -> level 1, 29px -> level 2, 49px -> level 3...
                if padding < 15: level = 1
                elif padding < 35: level = 2
                elif padding < 55: level = 3
                elif padding < 75: level = 4
                else: level = 5
            
            values = {}
            # FireAnt: các td giá trị thường bắt đầu từ index 1
            val_tds = [td for td in tds[1:] if not td.find(class_=re.compile('echarts|chart', re.I))]
            for i, td in enumerate(val_tds):
                if i < len(valid_years):
                    values[str(valid_years[i])] = parse_financial_number(td.get_text())
                    
            node = {
                "name": name,
                "values": values,
                "rows": [],
                "_level": level
            }
            
            while stack and stack[-1]['_level'] >= level:
                stack.pop()
                
            if stack:
                stack[-1]['rows'].append(node)
            else:
                result.append(node)
                
            stack.append(node)
            
        return result

def build_raw_data_from_htmls(htmls: List[str], ticker: str, source: str) -> Dict[str, Any]:
    raw_data = {
        "status": "OK",
        "ticker": ticker,
        "income": [],
        "balance": [],
        "cashflow": []
    }
    
    # Chọn parser dựa theo source
    parser = CafeFParser if source == "CafeF" else FireAntParser
    
    for html in htmls:
        html_lower = html.lower()
        json_tree = parser.parse(html)
        
        # Xác định bảng dựa trên từ khóa trong 500 ký tự đầu tiên
        content_sample = clean_text(BeautifulSoup(html, 'html.parser').get_text())[:500].lower()
        
        if "cân đối kế toán" in content_sample or "tài sản" in content_sample:
            if not raw_data["balance"]: raw_data["balance"] = json_tree
        elif "kết quả hoạt động" in content_sample or "kết quả kinh doanh" in content_sample or "doanh thu" in content_sample:
            if not raw_data["income"]: raw_data["income"] = json_tree
        elif "lưu chuyển tiền" in content_sample:
            if not raw_data["cashflow"]: raw_data["cashflow"] = json_tree
            
    return raw_data
