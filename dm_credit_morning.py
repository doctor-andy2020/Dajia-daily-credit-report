"""访问 DM 信用早报——通过 quote-web SPA"""
import asyncio
import json
import re
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        # 收集 API 响应
        api_data = []
        async def on_response(response):
            url = response.url
            if "rest.innodealing.com" in url and response.status == 200:
                try:
                    ct = response.headers.get("content-type", "")
                    if "json" in ct:
                        body = await response.json()
                        api_data.append({"url": url, "data": body})
                except: pass

        page.on("response", on_response)

        # Step 1: 登录 dashboard
        print("[1] 登录 dashboard...")
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

        # Step 2: 打开 quote-web SPA
        print("[2] 打开 quote-web...")
        # 使用 promote API 中的实际新闻 URL
        # 最新一篇新闻 ID 格式: YYYYMMDD + 序号
        from datetime import datetime
        today_str = datetime.now().strftime("%Y%m%d")

        # 尝试今天的新闻
        quote_url = f"https://web.innodealing.com/quote-web/#/bond/sentiment-news"
        await page.goto(quote_url, wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(10000)

        body = await page.locator("body").inner_text()
        print(f"    quote-web body 长度: {len(body)}")

        if len(body) < 50:
            print("    quote-web 加载失败，尝试直接打开新闻详情...")
            # 用 promote API 里拿到的 ID
            await page.goto(
                f"https://web.innodealing.com/quote-web/#/bond/sentiment-news-detail/area-news/{today_str}00010199544",
                wait_until="domcontentloaded", timeout=20000
            )
            await page.wait_for_timeout(10000)
            body = await page.locator("body").inner_text()
            print(f"    新闻详情 body 长度: {len(body)}")

        # Step 3: 搜索早报
        print("\n[3] 搜索关键词...")
        for kw in ["早报", "信用", "DM早报", "信用早报", "日报"]:
            idx = body.find(kw) if body else -1
            if idx >= 0:
                print(f"    ★ '{kw}' 位置 {idx}: {body[max(0,idx-30):idx+200]}")
            else:
                print(f"    '{kw}': 未找到")

        # Step 4: 分析 API 数据
        print(f"\n[4] 分析了 {len(api_data)} 个 API 响应")
        for item in api_data:
            url = item['url']
            data_str = json.dumps(item['data'], ensure_ascii=False)
            for kw in ['sentiment', 'morning', 'area-news', 'daily', 'article', 'news-detail', 'content']:
                if kw in url.lower() or kw.lower() in data_str.lower():
                    print(f"\n  [{kw}] {url[:150]}")
                    if len(data_str) > 2000:
                        # 只打印包含关键词的部分
                        for line in data_str.split(','):
                            if any(k in line.lower() for k in ['title','content','早报','morning','credit']):
                                print(f"    {line[:200]}")
                    else:
                        print(f"  {data_str[:1000]}")
                    break

        await browser.close()

asyncio.run(main())
