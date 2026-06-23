"""测试 quote-web iframe 中各种舆情路由"""
import asyncio
import os
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

OUT_DIR = os.path.dirname(os.path.abspath(__file__))

async def test_route(frame, route_hash, label, page):
    """测试一个 hash 路由并保存内容"""
    try:
        await frame.evaluate(f"window.location.hash = '{route_hash}'")
        await page.wait_for_timeout(8000)
        body = await frame.locator("body").inner_text()
        fpath = os.path.join(OUT_DIR, f"dm_{label}.txt")
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(body)
        print(f"  {label:30s} → {len(body):5d} chars | saved: dm_{label}.txt")
        # 搜索
        for kw in ["早报", "信用早报", "重要资讯", "DM信用"]:
            if kw in body:
                print(f"      ★ FOUND '{kw}'!")
        return body
    except Exception as e:
        print(f"  {label:30s} → ERROR: {e}")
        return ""

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
        print(f"    iframe: {frame.url}")

        # 测试各种路由
        print("\n[2] Testing routes...")
        routes = [
            ("#/bond/public-opinion/important-news", "important-news"),
            ("#/bond/public-opinion", "public-opinion"),
            ("#/bond/sentiment-news", "sentiment-news"),
            ("#/bond/sentiment-news/important", "sentiment-important"),
            ("#/bond/public-opinion/news-list", "news-list"),
            ("#/bond/public-opinion/sentiment", "opinion-sentiment"),
            ("#/news/important", "news-important"),
            ("#/bond/news", "bond-news"),
            ("#/bond/important-news", "bond-important-news"),
            ("#/credit/news", "credit-news"),
        ]
        for route_hash, label in routes:
            await test_route(frame, route_hash, label, page)

        # 也尝试直接 API 调用抓取文章列表
        print("\n[3] Trying to find article list API...")
        # 收集新的 API 请求
        new_urls = []
        def on_request(request):
            if "rest.innodealing.com" in request.url:
                new_urls.append(request.url)
        page.on("request", on_request)

        # 点击 iframe 中的"舆情"链接
        links = await frame.locator("text=舆情").all()
        print(f"    Found {len(links)} '舆情' links in iframe")
        for i, link in enumerate(links):
            try:
                print(f"    [{i}] visible={await link.is_visible()}, tag={await link.evaluate('el => el.tagName')}")
                if await link.is_visible():
                    await link.click()
                    await page.wait_for_timeout(5000)
                    print(f"        Clicked! New iframe URL: {frame.url}")
                    break
            except Exception as e:
                print(f"    [{i}] error: {e}")

        await browser.close()

asyncio.run(main())
