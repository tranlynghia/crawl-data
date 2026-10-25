from ownership_core.validator import parse_vietnamese_number

def test_failing_terminal_cases():
    cases = [
        "-Tir(.6)",
        "FPT cũng đặt mục tiêu đầy tham vọng - Mục tiêu 135. Đó",
        "BIDIPHAR - Bidiphar.com",
        "theo.”",
        "công ty gỗ. Sản",
    ]
    for c in cases:
        ratio, count = parse_vietnamese_number(c)
        assert ratio is None, f"Expected None for {c}, got {ratio}"
        assert count is None, f"Expected None for {c}, got {count}"

def test_valid_numbers():
    assert parse_vietnamese_number("51,08%") == (51.08, None)
    assert parse_vietnamese_number("51.08") == (51.08, None)
    assert parse_vietnamese_number("6.562.361") == (None, 6562361)
    assert parse_vietnamese_number("6,562,361") == (None, 6562361)
    assert parse_vietnamese_number("-") == (None, None)
    assert parse_vietnamese_number(None) == (None, None)
