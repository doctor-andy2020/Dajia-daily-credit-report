"""在 quote-web iframe 中导航到舆情文章，并读取内容"""
import asyncio
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

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
        await page.wait_for_timeout(10000)

        # 获取 iframe
        frame = page.frame("new-dashboard-frame")
        if not frame:
            print("ERROR: 找不到 iframe")
            await browser.close()
            return

        print(f"    iframe URL: {frame.url}")

        # 方法1: 在 iframe 中设置 hash
        print("\n[2] 在 iframe 中导航到 sentiment-news-detail...")
        # 使用之前 promote API 返回的具体 article URL
        article_hash = "#/bond/sentiment-news-detail/area-news/2026061200010199544"
        await frame.evaluate(f"window.location.hash = '{article_hash}'")
        await page.wait_for_timeout(8000)

        body = await frame.locator("body").inner_text()
        print(f"    页面 body: {len(body)} chars")
        for kw in ["早报", "信用", "DM", "日报", "城投", "地产", "利率"]:
            count = body.count(kw)
            if count > 0:
                print(f"    '{kw}': {count} 次")

        # 方法2: 导航到 sentiment-news 列表页
        print("\n[3] 导航到 sentiment-news 列表...")
        await frame.goto("https://web.innodealing.com/quote-web/#/bond/sentiment-news", wait_until="domcontentloaded", timeout=15000)
        await page.wait_for_timeout(8000)

        body2 = await frame.locator("body").inner_text()
        print(f"    页面 body: {len(body2)} chars")
        for kw in ["早报", "信用", "DM", "日报", "list"]:
            count = body2.count(kw)
            if count > 0:
                print(f"    '{kw}': {count} 次")

        # 方法3: 使用 page.goto 直接进入 iframe URL
        print("\n[4] 直接访问 quote-web 带 sentiment 路由...")
        await page.goto("https://web.innodealing.com/quote-web/", wait_until="domcontentloaded", timeout=15000)
        await page.wait_for_timeout(5000)
        await page.evaluate("window.location.hash = '#/bond/sentiment-news'")
        await page.wait_for_timeout(8000)

        body3 = await page.locator("body").inner_text()
        print(f"    页面 body: {len(body3)} chars")
        for kw in ["早报", "信用", "DM", "日报", "城投"]:
            count = body3.count(kw)
            if count > 0:
                print(f"    '{kw}': {count} 次")

        # 打印所有可见文本
        if body3:
            print(f"\n    body3 内容: {body3[:1500]}")

        await browser.close()

asyncio.run(main())
