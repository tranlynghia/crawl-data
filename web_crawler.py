import asyncio
from playwright.async_api import async_playwright, BrowserContext, Page
from typing import Dict, Any, List
import json

class WebCrawler:
    def __init__(self, headless: bool = True, user_data_dir: str = None, max_concurrent: int = 3):
        self.headless = headless
        self.user_data_dir = user_data_dir
        self.max_concurrent = max_concurrent
        self.semaphore = asyncio.Semaphore(max_concurrent)
        self.playwright = None
        self.browser = None
        self.context = None

    async def init_browser(self):
        if not self.playwright:
            self.playwright = await async_playwright().start()
            
            args = [
                "--disable-blink-features=AutomationControlled",
                "--disable-images",
                "--disable-extensions",
                "--no-sandbox",
                "--disable-setuid-sandbox"
            ]
            
            if self.user_data_dir:
                self.context = await self.playwright.chromium.launch_persistent_context(
                    user_data_dir=self.user_data_dir, 
                    headless=self.headless, 
                    args=args,
                    viewport={"width": 1280, "height": 800}
                )
            else:
                self.browser = await self.playwright.chromium.launch(
                    headless=self.headless, 
                    args=args
                )
                self.context = await self.browser.new_context(viewport={"width": 1280, "height": 800})
                
            # Block images and css for speed
            if self.headless:
                await self.context.route("**/*", self._route_intercept)

    async def _route_intercept(self, route):
        if route.request.resource_type in ["image", "media", "font"]:
            await route.abort()
        else:
            await route.continue_()

    async def close(self):
        if self.context and not self.user_data_dir:
            await self.context.close()
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

    async def _wait_and_click(self, page: Page, text_selector: str, timeout: int = 5000):
        try:
            # wait for selector text
            locator = page.locator(f"text='{text_selector}'").first
            await locator.wait_for(state="visible", timeout=timeout)
            await locator.click()
            await page.wait_for_load_state("networkidle", timeout=5000)
            return True
        except:
            return False

    async def crawl_fireant(self, ticker: str, page: Page) -> Dict[str, Any]:
        htmls = []
        try:
            url = f"https://fireant.vn/ma-chung-khoan/{ticker}"
            print(f"[{ticker}] FireAnt: Đang tải trang...")
            await page.goto(url, wait_until="networkidle", timeout=30000)
            
            # Đóng popup quảng cáo nếu có
            for btn_text in ["Đóng", "Bỏ qua", "Không, cảm ơn", "Close"]:
                try:
                    locator = page.locator(f"text='{btn_text}'").first
                    if await locator.is_visible(timeout=1000):
                        await locator.click()
                except:
                    pass

            clicked = await self._wait_and_click(page, "Tài chính")
            if not clicked:
                return {"status": "FAILED", "htmls": []}

            clicked = await self._wait_and_click(page, "Báo cáo tài chính")
            if not clicked:
                await self._wait_and_click(page, "Dữ liệu")

            # Chuyển sang Theo Năm
            try:
                selects = await page.locator("select").all()
                for select in selects:
                    text = await select.inner_text()
                    if "Theo năm" in text:
                        await select.select_option(label="Theo năm")
                        await page.wait_for_load_state("networkidle", timeout=3000)
                        break
            except Exception as e:
                print(f"[{ticker}] FireAnt Select Năm Lỗi: {e}")

            # Đợi bảng dữ liệu load
            try:
                await page.wait_for_selector("table", timeout=5000)
            except:
                pass

            htmls.append(await page.content()) # Cân đối kế toán mặc định

            if await self._wait_and_click(page, "Kết quả kinh doanh"):
                htmls.append(await page.content())
                
            if await self._wait_and_click(page, "Lưu chuyển tiền tệ trực tiếp") or await self._wait_and_click(page, "Lưu chuyển tiền tệ gián tiếp") or await self._wait_and_click(page, "Lưu chuyển tiền tệ"):
                htmls.append(await page.content())

            return {"status": "OK", "htmls": htmls}
        except Exception as e:
            print(f"[{ticker}] FireAnt Lỗi: {e}")
            return {"status": "FAILED", "htmls": []}

    async def crawl_cafef(self, ticker: str, page: Page) -> Dict[str, Any]:
        htmls = []
        try:
            ticker_lower = ticker.lower()
            found_url = None
            
            # Thay vì thử 3 sàn, thử trực tiếp link tổng quát của CafeF (thường tự redirect)
            url = f"https://s.cafef.vn/bao-cao-tai-chinh/{ticker_lower}/bsheet/2024/0/0/0/can-doi-ke-toan.chn"
            await page.goto(url, wait_until="networkidle", timeout=30000)
            content = await page.content()
            
            if "Không tìm thấy dữ liệu" in content or "tableContent" not in content and "table_KQKD" not in content:
                for san in ['hose', 'hnx', 'upcom']:
                    url = f"https://cafef.vn/du-lieu/{san}/{ticker_lower}-tai-chinh.chn"
                    await page.goto(url, wait_until="networkidle", timeout=20000)
                    content = await page.content()
                    if "Không tìm thấy dữ liệu" not in content and "Cân đối kế toán" in content:
                        found_url = url
                        break
            else:
                found_url = url
                
            if not found_url:
                return {"status": "FAILED", "htmls": []}

            # Nếu ở giao diện cũ (s.cafef.vn), dữ liệu đã hiển thị luôn.
            if "s.cafef.vn" in page.url:
                htmls.append(await page.content()) # Bsheet
                
                # Navigate to incsta
                await page.goto(page.url.replace('bsheet', 'incsta').replace('can-doi-ke-toan', 'ket-qua-hoat-dong-kinh-doanh'), wait_until="networkidle")
                htmls.append(await page.content())
                
                # Navigate to cashflow
                await page.goto(page.url.replace('incsta', 'cashflow').replace('ket-qua-hoat-dong-kinh-doanh', 'luu-chuyen-tien-te-tructiep'), wait_until="networkidle")
                htmls.append(await page.content())
                return {"status": "OK", "htmls": htmls}

            # Chuyển sang Theo Năm ở giao diện mới cafef.vn
            try:
                await page.evaluate("""
                    var el = document.querySelector('[data-time="NAM"]');
                    if(el) el.click();
                """)
                await page.wait_for_load_state("networkidle", timeout=3000)
                htmls.append(await page.content()) # CĐKT
                
                await page.evaluate("""
                    var el = document.querySelector('[data-view="KQKDHDSXKD"]');
                    if(el) el.click();
                """)
                await page.wait_for_load_state("networkidle", timeout=3000)
                htmls.append(await page.content()) # KQKD
                
                await page.evaluate("""
                    var el = document.querySelector('[data-view="LCTT_TT"]');
                    if(el) el.click();
                    else {
                        var el2 = document.querySelector('[data-view="LCTT_GT"]');
                        if(el2) el2.click();
                    }
                """)
                await page.wait_for_load_state("networkidle", timeout=3000)
                htmls.append(await page.content()) # LCTT
            except Exception as e:
                print(f"[{ticker}] CafeF JS Eval Lỗi: {e}")
                
            return {"status": "OK", "htmls": htmls}
        except Exception as e:
            print(f"[{ticker}] CafeF Lỗi: {e}")
            return {"status": "FAILED", "htmls": []}

    async def crawl_ticker(self, ticker: str):
        async with self.semaphore:
            page = await self.context.new_page()
            try:
                print(f"[{ticker}] Bắt đầu crawl FireAnt...")
                fireant_res = await self.crawl_fireant(ticker, page)
                if fireant_res["status"] == "OK" and fireant_res["htmls"]:
                    print(f"[{ticker}] Crawl FireAnt THÀNH CÔNG.")
                    return {"ticker": ticker, "source": "FireAnt", "htmls": fireant_res["htmls"]}
                
                print(f"[{ticker}] FireAnt thất bại. Fallback sang CafeF...")
                cafef_res = await self.crawl_cafef(ticker, page)
                if cafef_res["status"] == "OK" and cafef_res["htmls"]:
                    print(f"[{ticker}] Crawl CafeF THÀNH CÔNG.")
                    return {"ticker": ticker, "source": "CafeF", "htmls": cafef_res["htmls"]}
                    
                return {"ticker": ticker, "source": "None", "htmls": []}
            finally:
                await page.close()

async def run_crawlers_async(tickers: List[str], headless: bool = True, user_data_dir: str = None, max_concurrent: int = 3):
    crawler = WebCrawler(headless, user_data_dir, max_concurrent)
    await crawler.init_browser()
    try:
        tasks = [crawler.crawl_ticker(ticker) for ticker in tickers]
        results = await asyncio.gather(*tasks)
        return results
    finally:
        await crawler.close()

def run_crawlers(tickers: List[str], headless: bool = True, user_data_dir: str = None, max_concurrent: int = 3):
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    return loop.run_until_complete(run_crawlers_async(tickers, headless, user_data_dir, max_concurrent))
