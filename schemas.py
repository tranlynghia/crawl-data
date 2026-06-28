financial_statement_schema = {
    "type": "object",
    "properties": {
        "ticker": {"type": "string"},
        "year": {"type": "integer"},
        "report_type": {"type": "string"},
        "company_name": {"type": "string"},
        "unit": {"type": "string"},
        "rows": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "item_name": {"type": "string"},
                    "values": {
                        "type": "object",
                        "additionalProperties": {"type": ["number", "string", "null"]}
                    }
                },
                "required": ["item_name", "values"]
            }
        }
    },
    "required": ["ticker", "year", "report_type", "rows"]
}

# Prompt to extract the financial statement table
EXTRACT_PROMPT = """
Bạn đang đọc một trang báo cáo tài chính của CafeF. Hãy trích xuất toàn bộ bảng tài chính chính trên trang. 
Không tóm tắt. Không bỏ dòng. Giữ nguyên tên chỉ tiêu tiếng Việt. 
Trả về JSON theo schema, gồm ticker, year, report_type, company_name, unit và danh sách rows. 
Mỗi row gồm item_name và values. Nếu có nhiều cột kỳ báo cáo (nhiều năm), hãy giữ đầy đủ giá trị theo nhãn cột (key là năm). 
Nếu không tìm thấy bảng, trả về rows rỗng.
"""
