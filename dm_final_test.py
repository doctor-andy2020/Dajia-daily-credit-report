"""最终测试：读取文章内容并保存到 UTF-8 文件"""
import asyncio
import os
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

OUT_DIR = os.path.dirname(os.path.abspath(__file__))

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        # 登录
        print("[1] Login...")
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

        frame = page.frame("new-dashboard-frame")
        print(f"    iframe URL: {frame.url}")

        # 导航到 area-news 文章
        print("\n[2] Navigate to article...")
        await frame.evaluate("window.location.hash = '#/bond/sentiment-news-detail/area-news/2026061200010199544'")
        await page.wait_for_timeout(10000)

        # 获取所有文本
        body = await frame.locator("body").inner_text()
        print(f"    body length: {len(body)} chars")

        # 保存到文件
        fpath = os.path.join(OUT_DIR, "dm_article_content.txt")
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(body)
        print(f"[3] Saved: {fpath} ({len(body)} chars)")

        # 同时保存 HTML
        html = await frame.content()
        hpath = os.path.join(OUT_DIR, "dm_article.html")
        with open(hpath, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"[4] Saved HTML: {hpath} ({len(html)} chars)")

        # 搜索 "早报"
        for kw in ["早报", "信用早报", "DM早报", "重要资讯"]:
            idx = body.find(kw)
            if idx >= 0:
                print(f"    ★ FOUND '{kw}' at position {idx}!")
            else:
                print(f"    '{kw}' not found")

        await browser.close()

asyncio.run(main())
