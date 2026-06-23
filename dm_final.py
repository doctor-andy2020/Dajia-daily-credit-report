"""Final DM article extractor with article ID from API"""
import asyncio
import json
import os
import re
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
                        if isinstance(body, dict):
                            api_data.append({"url": url, "data": body})
                except:
                    pass

        page.on("response", on_resp)

        # Login with retry
        print("[1] Login...")
        await page.goto("https://web.innodealing.com/dashboard/", wait_until="networkidle")
        await page.wait_for_timeout(5000)
        await page.fill("#inputUsername", DM_USER)
        await page.fill("#inputPassword", DM_PASS)
        await page.click("button[type=submit]")

        # Wait for dashboard to fully load (up to 20s)
        for i in range(20):
            await page.wait_for_timeout(1000)
            frame = page.frame("new-dashboard-frame")
            if frame and "quote-web" in frame.url:
                print(f"    Dashboard ready after {i+1}s")
                break
        else:
            print("    ERROR: Dashboard did not load")
            await page.screenshot(path=os.path.join(OUT, "dm_error.png"))
            await browser.close()
            return

        # Clear login responses
        api_data.clear()

        # Navigate to important-news to trigger article list API
        print("[2] Loading article list...")
        await frame.evaluate("window.location.hash = '#/bond/public-opinion/important-news'")
        await page.wait_for_timeout(12000)

        # Find articles from API
        print("[3] Searching API responses for articles...")
        articles = []
        for item in api_data:
            if "sentiment/news/paging" in item["url"]:
                articles = item["data"].get("data", {}).get("list", [])
                print(f"    Found {len(articles)} articles")
                break

        if not articles:
            # Try headline list
            for item in api_data:
                if "headline/list" in item["url"]:
                    articles = item["data"].get("data", [])
                    print(f"    Found {len(articles)} headlines")
                    break

        target = None
        for a in articles:
            title = a.get("title", "") or a.get("sentimentTitle", "")
            sid = a.get("sentimentId", "")
            if sid and ("要闻速览" in title or "信用早报" in title or "DM早报" in title or "早报" in title):
                target = {"id": sid, "title": title}
                print(f"    ★ {sid}: {title}")

        if not target and articles:
            # Print first 10 for debugging
            print("    Available titles (first 10):")
            for a in articles[:10]:
                title = a.get("title", "") or a.get("sentimentTitle", "")
                sid = a.get("sentimentId", "")
                print(f"      {sid}: {title[:80]}")

        if target:
            article_id = target["id"]
            print(f"\n[4] Opening article: {target['title']}")
            print(f"    ID: {article_id}")

            await frame.evaluate(
                f"window.location.hash = '#/bond/sentiment-news-detail/area-news/{article_id}'"
            )
            await page.wait_for_timeout(10000)

            body = await frame.locator("body").inner_text()
            filepath = os.path.join(OUT, "dm_article_output.txt")
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(body)
            print(f"    Body: {len(body)} chars")
            print(f"    Saved: {filepath}")
            print(f"\n    Preview:")
            print(body[:1000])
        else:
            print("\n    No matching article found.")
            # Save API data for debugging
            with open(os.path.join(OUT, "dm_debug_api.json"), "w", encoding="utf-8") as f:
                json.dump([{"url": a["url"], "data": a["data"]} for a in api_data], f, ensure_ascii=False, indent=2)
            print("    Saved API dump to dm_debug_api.json")

        await browser.close()

asyncio.run(main())
