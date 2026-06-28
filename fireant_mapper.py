from typing import Dict, Any, Tuple, Optional, Iterable
import unicodedata
import re
from decimal import Decimal, InvalidOperation

METRIC_ALIASES = {
    "total_assets": ("Tổng cộng tài sản", "TỔNG CỘNG TÀI SẢN", "Tổng tài sản"),
    "liabilities": ("Nợ phải trả", "Tổng nợ phải trả", "Total liabilities", "Tổng nợ", "Liability"),
    "net_income": (
        "Lợi nhuận sau thuế",
    ),
    "cfo": (
        "Lưu chuyển tiền thuần từ hoạt động kinh doanh",
        "Lưu chuyển tiền thuần từ HĐKD",
        "Dòng tiền thuần từ hoạt động kinh doanh",
    ),
    "revenue": ("Doanh thu thuần về bán hàng và cung cấp dịch vụ", "Doanh thu thuần", "3. Doanh thu thuần (1)-(2)"),
    "short_term_receivables": ("Các khoản phải thu ngắn hạn",),
    "inventory": ("Hàng tồn kho", "Tổng hàng tồn kho", "IV. Hàng tồn kho"),
    "ppe": ("- Nguyên giá", "Nguyên giá tài sản cố định hữu hình", "Nguyên giá TSCĐ"),
}

def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    text = text.upper().replace("Đ", "D")
    return re.sub(r"\s+", " ", text).strip()

def compact_text(value: object) -> str:
    return re.sub(r"[^A-Z0-9]+", "", normalize_text(value))

def parse_financial_number(value: object) -> float | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, (int, float, Decimal)):
        return float(value)
    text = str(value).strip()
    if text in {"-", "--", "N/A", "null", "None"}:
        return None
    negative = text.startswith("(") and text.endswith(")")
    cleaned = re.sub(r"[^\d,.\-]", "", text.strip("()"))
    if not cleaned:
        return None
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(",", "")
    elif cleaned.count(",") > 0:
        cleaned = cleaned.replace(",", "")
    elif cleaned.count(".") > 1 or (
        cleaned.count(".") == 1 and len(cleaned.rsplit(".", maxsplit=1)[1]) == 3
    ):
        cleaned = cleaned.replace(".", "")
    try:
        number = Decimal(cleaned)
    except InvalidOperation:
        return None
    return float(-number if negative else number)

def first_alias_match(label: object, aliases: Iterable[str]) -> bool:
    normalized_label = normalize_text(label)
    compact_label = compact_text(label)
    for alias in aliases:
        if compact_text(alias) in compact_label:
            return True
        alias_words = normalize_text(alias).split()
        if alias_words and all(word in normalized_label.split() for word in alias_words):
            return True
    return False

def _year_column_index(columns: object, year: int) -> int | None:
    if not isinstance(columns, list):
        return None
    for index, column in enumerate(columns):
        if str(year) in str(column):
            return index
    return None

def _read_value(values: object, year: int, column_index: int | None) -> object:
    if isinstance(values, dict):
        return values.get(str(year), values.get(year))
    if isinstance(values, list):
        for item in values:
            if (
                isinstance(item, dict)
                and int(item.get("year", 0) or 0) == year
                and int(item.get("quarter", 0) or 0) == 0
            ):
                return item.get("value")
        if column_index is not None and len(values) > column_index:
            item = values[column_index]
            if isinstance(item, dict):
                return item.get("value") or item.get(str(year))
            return item
    return None

def _iter_row_values(rows: object, year: int, column_index: int | None, parent_name: str = ""):
    if not isinstance(rows, list):
        return
    for row in rows:
        if isinstance(row, dict):
            label = next(
                (row.get(key) for key in ("name", "title", "label", "itemName", "normName", "accountName", "rowName") if row.get(key)),
                None,
            )
            values = row.get("values") or row.get("financialValues")
            value = _read_value(values, year, column_index)
            if label is not None:
                yield label, value, parent_name
            nested = row.get("rows") or row.get("children")
            if nested:
                yield from _iter_row_values(nested, year, column_index, parent_name=str(label) if label else parent_name)
        elif isinstance(row, list) and row:
            label = next((cell for cell in row if isinstance(cell, str) and not cell.isdigit()), row[0])
            if column_index is None:
                value = None
            elif len(row) > column_index and row[column_index] != label:
                value = row[column_index]
            elif len(row) > column_index + 1:
                value = row[column_index + 1]
            else:
                value = None
            yield label, value, parent_name

class FireantMapper:
    def __init__(self, raw_data: Dict[str, Any]):
        self.profile = raw_data.get("profile", {})
        self.income = raw_data.get("income", [])
        self.balance = raw_data.get("balance", [])
        self.cashflow = raw_data.get("cashflow", [])
        self.cashflow_in = raw_data.get("cashflow_in", [])

    def _find_value(self, statement: Any, year: int, metric_key: str) -> Tuple[Optional[float], Optional[str]]:
        aliases = METRIC_ALIASES.get(metric_key, [])
        columns = statement.get("columns", []) if isinstance(statement, dict) else []
        rows = statement.get("rows", []) if isinstance(statement, dict) else statement
        column_index = _year_column_index(columns, year)
        
        for label, value, parent_name in _iter_row_values(rows, year, column_index):
            if value is not None and first_alias_match(label, aliases):
                # Exclusions
                norm_label = normalize_text(label)
                norm_parent = normalize_text(parent_name)
                
                if metric_key == "ppe":
                    if "- NGUYEN GIA" in norm_label or "NGUYEN GIA" == norm_label:
                        if "TAI SAN CO DINH HUU HINH" not in norm_parent:
                            continue
                
                if metric_key == "total_assets":
                    if any(excl in norm_label for excl in ["KHAC", "NGAN HAN", "DAI HAN", "DO DANG"]):
                        continue

                if metric_key == "liabilities":
                    if any(excl in norm_label for excl in ["NGAN HAN", "DAI HAN", "KHAC", "NGUOI BAN", "NO THUE", "TRUOC"]):
                        continue
                        
                if metric_key == "inventory":
                    if "DU PHONG" in norm_label:
                        continue
                        
                if metric_key == "revenue" and any(excl in norm_label for excl in ["TAI CHINH", "KHAC", "HOAT DONG"]):
                    continue
                    
                val = parse_financial_number(value)
                if val is not None:
                    return val, str(label)
        return None, None

    def get_total_assets(self, year: int) -> Tuple[Optional[float], Optional[str]]:
        return self._find_value(self.balance, year, "total_assets")

    def get_liabilities(self, year: int) -> Tuple[Optional[float], Optional[str]]:
        return self._find_value(self.balance, year, "liabilities")
        
    def get_cfo(self, year: int) -> Tuple[Optional[float], Optional[str]]:
        val, label = self._find_value(self.cashflow, year, "cfo")
        if val is not None:
            return val, label
        return self._find_value(self.cashflow_in, year, "cfo")
        
    def get_net_income(self, year: int) -> Tuple[Optional[float], Optional[str]]:
        return self._find_value(self.income, year, "net_income")
        
    def get_revenue(self, year: int) -> Tuple[Optional[float], Optional[str]]:
        return self._find_value(self.income, year, "revenue")
        
    def get_short_term_receivables(self, year: int) -> Tuple[Optional[float], Optional[str]]:
        return self._find_value(self.balance, year, "short_term_receivables")
        
    def get_inventory(self, year: int) -> Tuple[Optional[float], Optional[str]]:
        return self._find_value(self.balance, year, "inventory")
        
    def get_ppe(self, year: int) -> Tuple[Optional[float], Optional[str]]:
        return self._find_value(self.balance, year, "ppe")

def build_unified_data(mapper: FireantMapper, ticker: str, start_year: int, end_year: int, unit_multiplier: float = 1.0) -> Dict[str, Any]:
    company_name = mapper.profile.get("companyName", ticker)
    industry = mapper.profile.get("industry", "")
    
    def scale(v):
        return v * unit_multiplier if v is not None else None

    years_data = {}
    for year in range(start_year - 1, end_year + 1):
        ta, ta_src = mapper.get_total_assets(year)
        ta_prev, _ = mapper.get_total_assets(year - 1)
        liab, liab_src = mapper.get_liabilities(year)
        cfo, cfo_src = mapper.get_cfo(year)
        ni, ni_src = mapper.get_net_income(year)
        rev, rev_src = mapper.get_revenue(year)
        rec, rec_src = mapper.get_short_term_receivables(year)
        inv, inv_src = mapper.get_inventory(year)
        ppe, ppe_src = mapper.get_ppe(year)
        
        # calculate ROA
        roa, roa_src = None, None
        if ni is not None and ta is not None and ta_prev is not None and (ta_prev + ta) != 0:
            roa = ni / ((ta_prev + ta) / 2)
            roa_src = "Tự tính (LNST / TA trung bình)"
        
        years_data[year] = {
            "total_assets": {"value": scale(ta), "source": ta_src},
            "liabilities": {"value": scale(liab), "source": liab_src},
            "operating_cash_flow": {"value": scale(cfo), "source": cfo_src},
            "net_profit_after_tax": {"value": scale(ni), "source": ni_src},
            "short_term_receivables": {"value": scale(rec), "source": rec_src},
            "inventory": {"value": scale(inv), "source": inv_src},
            "ppe_gross": {"value": scale(ppe), "source": ppe_src},
            "revenue": {"value": scale(rev), "source": rev_src},
            "roa": {"value": roa, "source": roa_src},
        }
        
    return {
        "ticker": ticker,
        "company_name": company_name,
        "industry": industry,
        "years": years_data
    }
