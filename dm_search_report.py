"""保存 API 响应到文件并搜索信用早报"""
import asyncio
import json
import os
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        all_responses = []

        async def on_response(response):
            url = response.url
            if "rest.innodealing.com" in url and response.status == 200:
                try:
                    ct = response.headers.get("content-type", "")
                    if "json" in ct:
                        body = await response.json()
                        all_responses.append({"url": url, "data": body})
                except: pass

        page.on("response", on_response)

        # 登录
        print("[1] 登录...")
        await page.goto("https://web.innodealing.com/dashboard/", wait_until="networkidle")
        await page.wait_for_timeout(3000)
        await page.fill("#inputUsername", DM_USER)
        await page.fill("#inputPassword", DM_PASS)
        try:
            privacy = page.locator("#approve-privacy")
            if await privacy.is_visible(): await privacy.check()
        except: pass
        await page.click("button[type='submit']")
        await page.wait_for_timeout(15000)

        print(f"    收集到 {len(all_responses)} 个响应")

        # 保存到文件（UTF-8，不做任何转换）
        outfile = os.path.join(os.path.dirname(__file__), "dm_api_dump.json")
        with open(outfile, "w", encoding="utf-8") as f:
            json.dump(all_responses, f, ensure_ascii=False, indent=2)
        print(f"    已保存到: {outfile}")

        # 在内存中搜索
        print("\n[2] 全文搜索关键词...")
        full_text = json.dumps(all_responses, ensure_ascii=False)
        for kw in ["信用早报", "DM早报", "早报", "DM信用",
                   "每日信用", "morning report", "daily credit",
                   "sentiment-news-detail", "area-news"]:
            idx = full_text.find(kw)
            if idx >= 0:
                ctx = full_text[max(0,idx-40):idx+200]
                print(f"  ★ '{kw}' 找到:")
                print(f"    ...{ctx}...")
            else:
                print(f"  '{kw}': 未找到")

        await browser.close()

asyncio.run(main())
