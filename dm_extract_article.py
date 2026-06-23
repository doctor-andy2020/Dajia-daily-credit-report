"""Robust DM article extractor - handles iframe, API interception, and article detail"""
import asyncio
import json
import os
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"
OUT = os.path.dirname(os.path.abspath(__file__))

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context()
        page = await ctx.new_page()

        api_data = []

        async def on_resp(response):
            url = response.url
            if "rest.innodealing.com" in url and response.status == 200:
                try:
                    ct = response.headers.get("content-type", "")
                    if "json" in ct:
                        body = await response.json()
                        if isinstance(body, dict) and "code" in body:
                            api_data.append({"url": url, "data": body})
                except:
                    pass

        page.on("response", on_resp)

        # Login
        print("[1] Login...")
        await page.goto("https://web.innodealing.com/dashboard/", wait_until="networkidle")
        await page.wait_for_timeout(3000)
        await page.fill("#inputUsername", DM_USER)
        await page.fill("#inputPassword", DM_PASS)
        await page.click("button[type=submit]")
        await page.wait_for_timeout(12000)

        # Wait for iframe
        print("[2] Finding iframe...")
        for i in range(10):
            frame = page.frame("new-dashboard-frame")
            if frame:
                print(f"    Found after {i}s")
                break
            await page.wait_for_timeout(1000)
        else:
            print("    ERROR: iframe not found")
            await browser.close()
            return

        # Clear login API data
        api_data.clear()

        # Navigate to important-news
        print("[3] Navigate to important-news...")
        await frame.evaluate("window.location.hash = '#/bond/public-opinion/important-news'")
        await page.wait_for_timeout(12000)

        # Analyze API responses
        print(f"\n[4] Analyzing {len(api_data)} API responses...")
        for i, item in enumerate(api_data):
            d = item["data"]
            ds = json.dumps(d, ensure_ascii=False)
            url = item["url"].split("rest.innodealing.com")[-1]

            has_date_id = "2026061" in ds
            has_article_data = any(
                kw in ds.lower()
                for kw in ["title", "articleid", "newsid", "contenttext", "summarytext"]
            )

            if has_date_id or has_article_data:
                print(f"\n  [{i}] dateId={has_date_id} article={has_article_data}")
                print(f"  URL: {url[:150]}")
                print(f"  Data: {ds[:800]}")
                print()

                # If it contains article IDs, extract them
                if has_date_id:
                    import re
                    ids = re.findall(r"2026061\d{8,}", ds)
                    if ids:
                        unique_ids = list(set(ids))
                        print(f"  ★ Found article IDs: {unique_ids[:10]}")
                        # Save for later use
                        with open(os.path.join(OUT, "dm_article_ids.json"), "w") as f:
                            json.dump(unique_ids, f)

        await browser.close()

asyncio.run(main())
