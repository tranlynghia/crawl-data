import json
import asyncio
from typing import Dict, Any, List
from playwright.async_api import async_playwright

class FireAntClient:
    def __init__(self, headless=True, user_data_dir=None):
        self.headless = headless
        self.user_data_dir = user_data_dir

    async def _fetch_data(self, ticker: str, end_year: int) -> Dict[str, Any]:
        async with async_playwright() as p:
            if self.user_data_dir:
                context = await p.chromium.launch_persistent_context(
                    user_data_dir=self.user_data_dir,
                    headless=self.headless,
                    args=["--disable-blink-features=AutomationControlled"]
                )
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
            
            url = f"https://web.fireant.vn/ma-chung-khoan/{ticker}"
            print(f"[{ticker}] Mở trang: {url}")
            try:
                await page.goto(url, wait_until="networkidle", timeout=30000)
            except Exception as e:
                print(f"[{ticker}] Timeout khi mở trang: {e}")
                
            # Wait a bit to ensure we capture a token
            for _ in range(10):
                if bearer_token:
                    break
                await page.wait_for_timeout(500)
                
            industry = ""
            try:
                print(f"[{ticker}] Click tab 'Hồ sơ'")
                await page.locator("text='Hồ sơ'").first.click(timeout=5000)
                await page.wait_for_timeout(2000)
                html = await page.content()
                from bs4 import BeautifulSoup
                import re
                soup = BeautifulSoup(html, "lxml")
                nganh_elem = soup.find(string=re.compile(r"Ngành:"))
                if nganh_elem and nganh_elem.parent:
                    industry = nganh_elem.parent.text.replace("Ngành:", "").strip()
                else:
                    for kw in ["Lĩnh vực", "Ngành nghề"]:
                        kw_elem = soup.find(string=re.compile(kw))
                        if kw_elem and kw_elem.parent:
                            industry = kw_elem.parent.text.replace(kw, "").replace(":", "").strip()
                            break
            except Exception as e:
                print(f"[{ticker}] Lỗi khi lấy DOM Hồ sơ: {e}")
                
            try:
                print(f"[{ticker}] Click tab 'Tài chính'")
                await page.locator("text='Tài chính'").first.click(timeout=5000)
                await page.wait_for_timeout(1000)
                print(f"[{ticker}] Click tab 'Dữ liệu'")
                await page.locator("text='Dữ liệu'").first.click(timeout=5000)
                await page.wait_for_timeout(1000)
            except:
                pass

                
            if not bearer_token:
                if self.user_data_dir:
                    await context.close()
                else:
                    await browser.close()
                return {"status": "MISSING_OR_NO_PERMISSION"}
                
            print(f"[{ticker}] Đã bắt được token, tiến hành kéo dữ liệu Năm...")
            headers = {"Authorization": bearer_token, "Accept": "application/json"}
            
            import requests
            def fetch_api(path: str):
                try:
                    resp = requests.get(f"https://restv2.fireant.vn/symbols/{ticker}{path}", headers=headers, timeout=15)
                    if resp.status_code == 200:
                        return resp.json()
                except Exception as e:
                    print(f"[{ticker}] Lỗi gọi API {path}: {e}")
                return None
                
            profile = fetch_api("") or {}
            profile["industry"] = industry
            profile["company_name"] = profile.get("name", ticker)
            
            bsheet = fetch_api(f"/full-financial-reports?type=1&year={end_year}&quarter=0&limit=10")
            income = fetch_api(f"/full-financial-reports?type=2&year={end_year}&quarter=0&limit=10")
            cashflow = fetch_api(f"/full-financial-reports?type=4&year={end_year}&quarter=0&limit=10")
            
            if self.user_data_dir:
                await context.close()
            else:
                await browser.close()
                
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

    def scrape_symbol_financials(self, ticker: str, start_year: int, end_year: int) -> dict:
        try:
            loop = asyncio.get_event_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
        return loop.run_until_complete(self._fetch_data(ticker, end_year))
