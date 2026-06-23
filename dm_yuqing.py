"""导航到 DM 舆情模块，定位信用早报"""
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

        print(f"    登录成功: {page.url}")

        # 导航到舆情页面
        print("\n[2] 导航到舆情模块...")
        yuqing_url = "https://web.innodealing.com/dashboard/#/bond/public-opinion/important-news"
        await page.goto(yuqing_url, wait_until="networkidle")
        await page.wait_for_timeout(8000)

        body = await page.locator("body").inner_text()
        print(f"    页面文本: {len(body)} 字符")
        print(f"    前500字: {body[:500]}")

        # 搜索"早报"
        for kw in ["早报", "信用", "DM", "日报"]:
            count = body.count(kw)
            print(f"    '{kw}': {count} 次")

        # 打印新捕获的 API 响应中包含 sentiment 关键词的
        print("\n[3] 舆情相关 API 响应:")
        for item in api_data:
            url = item['url']
            data_str = json.dumps(item['data'], ensure_ascii=False)
            for kw in ['sentiment', 'public-opinion', 'opinion', 'yuqing', 'area-news', 'important-news', 'article']:
                if kw in url.lower() or kw in data_str.lower():
                    print(f"\n  [{kw}] {url}")
                    print(f"  {data_str[:800]}")
                    break

        # 截图
        await page.screenshot(path="dm_yuqing.png", full_page=True)
        print("\n[4] 截图: dm_yuqing.png")

        await browser.close()

asyncio.run(main())
