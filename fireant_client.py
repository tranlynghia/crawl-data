import json
import asyncio
import requests
from typing import Dict, Any, List
from playwright.async_api import async_playwright

class FireAntClient:
    def __init__(self, headless=True, user_data_dir=None):
        self.headless = headless
        self.user_data_dir = user_data_dir
        self._bearer_token = None
        self._playwright_context = None
        self._playwright_browser = None

    async def _get_token(self):
        if self._bearer_token:
            return self._bearer_token
            
        async with async_playwright() as p:
            if self.user_data_dir:
                context = await p.chromium.launch_persistent_context(
                    user_data_dir=self.user_data_dir,
                    headless=self.headless,
                    args=["--disable-blink-features=AutomationControlled"]
                )
                browser = None
            else:
                browser = await p.chromium.launch(headless=self.headless, args=["--disable-blink-features=AutomationControlled"])
                context = await browser.new_context()
                
            page = await context.new_page()
            
            bearer_token = None
            
            def handle_request(request):
                nonlocal bearer_token
                if "restv2.fireant.vn" in request.url:
                    auth = request.headers.get("authorization")
                    if auth:
                        bearer_token = auth
                        
            page.on("request", handle_request)
            
            url = "https://web.fireant.vn/"
            print("Đang lấy Access Token từ FireAnt...")
            try:
                await page.goto(url, wait_until="networkidle", timeout=30000)
            except Exception as e:
                print(f"Timeout khi mở trang chủ: {e}")
                
            # Wait a bit to ensure we capture a token
            for _ in range(10):
                if bearer_token:
                    break
                await page.wait_for_timeout(500)
                
            self._bearer_token = bearer_token
            
            if self.user_data_dir:
                await context.close()
            else:
                await browser.close()
                
        return self._bearer_token

    def scrape_symbol_financials(self, ticker: str, start_year: int, end_year: int) -> dict:
        try:
            loop = asyncio.get_event_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
        token = loop.run_until_complete(self._get_token())
        
        if not token:
            print(f"[{ticker}] Lỗi: Không thể lấy token.")
            return {"status": "MISSING_OR_NO_PERMISSION"}
            
        print(f"[{ticker}] Lấy dữ liệu từ API...")
        headers = {"Authorization": token, "Accept": "application/json"}
        
        def fetch_api(path: str):
            try:
                resp = requests.get(f"https://restv2.fireant.vn/symbols/{ticker}{path}", headers=headers, timeout=15)
                if resp.status_code == 200:
                    return resp.json()
            except Exception as e:
                print(f"[{ticker}] Lỗi gọi API {path}: {e}")
            return None
            
        profile = fetch_api("") or {}
        # Using industry from base endpoint
        profile["industry"] = profile.get("industry", "")
        profile["company_name"] = profile.get("name", ticker)
        
        bsheet = fetch_api(f"/full-financial-reports?type=1&year={end_year}&quarter=0&limit=10")
        income = fetch_api(f"/full-financial-reports?type=2&year={end_year}&quarter=0&limit=10")
        cashflow = fetch_api(f"/full-financial-reports?type=4&year={end_year}&quarter=0&limit=10")
        
        if not bsheet and not income:
            return {"status": "MISSING_OR_NO_PERMISSION"}
            
        return {
            "status": "OK",
            "profile": profile or {},
            "balance": bsheet,
            "income": income,
            "cashflow": cashflow,
            "logs": []
        }
