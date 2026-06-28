import openpyxl
from typing import List, Dict, Any, Tuple

class ExcelInspector:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.wb = openpyxl.load_workbook(file_path, data_only=True)
    
    def get_sheet_names(self) -> List[str]:
        return self.wb.sheetnames
        
    def find_header_row(self, sheet_name: str) -> Tuple[int, Dict[str, int]]:
        """
        Finds the header row by looking for 'MÃ CK' or 'TÊN CÔNG TY'.
        Returns (row_idx, {col_name: col_idx (1-based)})
        """
        if sheet_name not in self.wb.sheetnames:
            return -1, {}
            
        sheet = self.wb[sheet_name]
        for r in range(1, 10): # Check first 10 rows
            headers = {}
            for c in range(1, sheet.max_column + 1):
                val = sheet.cell(row=r, column=c).value
                if val and isinstance(val, str):
                    val = val.strip()
                    # Some headers might have linebreaks or extra spaces
                    headers[val] = c
            
            # Simple heuristic
            if any("MÃ CK" in k for k in headers.keys()) or any("ID" in k for k in headers.keys()):
                # Normalize keys slightly for easier lookup later
                norm_headers = {}
                for k, v in headers.items():
                    if "Unnamed" in k:
                        continue
                    if "Năm (t)" in k or k == "Năm":
                        norm_headers["Năm"] = v
                    elif "ID" == k:
                        norm_headers["ID"] = v
                    elif "TÊN CÔNG TY" in k:
                        norm_headers["TÊN CÔNG TY"] = v
                    elif "MÃ CK" in k:
                        norm_headers["MÃ CK"] = v
                    elif "NGÀNH" in k.upper():
                        norm_headers["NGÀNH"] = v
                    elif "Tổng tài sản đầu kì" in k:
                        norm_headers["Tổng tài sản đầu kì"] = v
                    elif "Tổng tài sản năm t" in k:
                        norm_headers["Tổng tài sản năm t"] = v
                    elif "Nợ phải trả năm t" in k:
                        norm_headers["Nợ phải trả năm t"] = v
                    elif "Dòng tiền từ hoạt động kinh doanh" in k or "CFO (Dòng tiền" in k and "CFO" not in norm_headers:
                        if "CFO" in k:
                            norm_headers["CFO"] = v
                        else:
                            norm_headers["Dòng tiền từ hoạt động kinh doanh"] = v
                    elif "Lợi nhuận sau thuế" in k:
                        norm_headers["Lợi nhuận sau thuế"] = v
                    elif "Khoản phải thu ngắn hạn năm t-1" in k:
                        norm_headers["Khoản phải thu ngắn hạn năm t-1"] = v
                    elif "Khoản phải thu ngắn hạn năm t" in k and "t-1" not in k:
                        norm_headers["Khoản phải thu ngắn hạn năm t"] = v
                    elif "Nguyên giá tài sản cố định" in k:
                        norm_headers["Nguyên giá tài sản cố định"] = v
                    elif "ROA" in k:
                        norm_headers["ROA"] = v
                    elif "Doanh thu t-1" in k:
                        norm_headers["Doanh thu t-1"] = v
                    elif "Doanh thu t" in k and "t-1" not in k:
                        norm_headers["Doanh thu t"] = v
                    elif "DA" in k and "DATA" not in k:
                        norm_headers["DA"] = v
                    elif "Big4" in k:
                        norm_headers["Big4"] = v
                    elif "DIGI" in k:
                        norm_headers["DIGI"] = v
                    elif "Size" in k:
                        norm_headers["Size"] = v
                    elif "LEV" in k:
                        norm_headers["LEV"] = v
                    elif "TA" in k:
                        norm_headers["TA"] = v
                    elif "REC" in k:
                        norm_headers["REC"] = v
                    elif "GROWTH" in k:
                        norm_headers["GROWTH"] = v
                    else:
                        norm_headers[k] = v
                        
                # Fix for the empty header being Year
                if 1 not in norm_headers.values() and sheet.cell(row=r+1, column=1).value in [2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]:
                    norm_headers["Năm"] = 1
                elif "Năm (t)" not in norm_headers and "Năm" not in norm_headers:
                     # Check if cell(r, 1) is empty but next row has year
                     if sheet.cell(row=r, column=1).value is None:
                         norm_headers["Năm"] = 1
                         
                return r, norm_headers
        return -1, {}

    def get_existing_symbols_and_max_id(self) -> Tuple[List[str], int]:
        symbols = set()
        max_id = 0
        
        for sheet_name in ["DATA ", "DATA DA", "DATA"]:
            if sheet_name not in self.wb.sheetnames:
                continue
            r, headers = self.find_header_row(sheet_name)
            if r == -1 or "MÃ CK" not in headers:
                continue
                
            ticker_col = headers["MÃ CK"]
            id_col = headers.get("ID")
            sheet = self.wb[sheet_name]
            
            for row_idx in range(r + 1, sheet.max_row + 1):
                val = sheet.cell(row=row_idx, column=ticker_col).value
                if val and isinstance(val, str):
                    symbols.add(val.strip().upper())
                
                if id_col:
                    id_val = sheet.cell(row=row_idx, column=id_col).value
                    try:
                        if id_val is not None:
                            id_num = int(id_val)
                            if id_num > max_id:
                                max_id = id_num
                    except (ValueError, TypeError):
                        pass
                        
        return list(symbols), max_id
