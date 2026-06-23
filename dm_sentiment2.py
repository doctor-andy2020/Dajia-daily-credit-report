"""捕获 DM 舆情 API 响应 — 通过拦截 SPA 的网络请求"""
import asyncio
import json
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        # 收集所有 API 响应（带有正确 sign 计算）
        api_data = []

        async def on_response(response):
            url = response.url
            if "rest.innodealing.com" in url and response.status == 200:
                try:
                    ct = response.headers.get("content-type", "")
                    if "json" in ct:
                        body = await response.json()
                        api_data.append({"url": url, "data": body})
                except:
                    pass

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
        await page.wait_for_timeout(10000)

        # 导航到 sentiment 页面
        print("\n[2] 导航到 sentiment news 页面...")
        await page.goto("https://web.innodealing.com/quote-web/#/bond/sentiment-news", wait_until="networkidle")
        await page.wait_for_timeout(5000)

        # 搜索"早报"和"信用"
        page_text = await page.locator("body").inner_text()
        print(f"    页面文本长度: {len(page_text)}")
        for kw in ["早报", "信用", "DM", "日报"]:
            positions = [i for i in range(len(page_text)) if page_text[i:i+len(kw)] == kw]
            print(f"    '{kw}': 找到 {len(positions)} 处")

        # 打印所有捕获的 API 响应中包含关键词的
        print("\n[3] 分析 API 响应...")
        for item in api_data:
            url = item['url']
            data_str = json.dumps(item['data'], ensure_ascii=False)
            for kw in ['sentiment', '舆情', '早报', 'credit', 'morning', 'news', 'area','区域']:
                if kw in url.lower() or kw in data_str:
                    print(f"\n  [{kw}] {url[:120]}")
                    print(f"  {data_str[:500]}")
                    break

        # 保存页面 HTML
        html = await page.content()
        with open("dm_sentiment_page.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("\n[4] 页面 HTML 已保存: dm_sentiment_page.html")

        await browser.close()

asyncio.run(main())
