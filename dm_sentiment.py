"""直接访问 DM 舆情页面，提取信用早报"""
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

        # 先登录
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
        await page.wait_for_timeout(8000)

        # 从 promote API 获取舆情链接
        print("\n[2] 获取推广位数据（含舆情链接）...")
        resp = await page.evaluate("""async () => {
            const r = await fetch('/onshore-platform-service/api/promote/screen/list/by/userid?' +
                'timestamp=' + Date.now() + '&v=13426&qIntl=0&sign=abc');
            return await r.json();
        }""")
        print(f"    {json.dumps(resp, ensure_ascii=False, indent=2)[:3000]}")

        # 尝试直接导航到舆情页面
        print("\n[3] 尝试导航到舆情相关页面...")
        sentiment_urls = [
            "https://web.innodealing.com/quote-web/#/bond/sentiment-news",
            "https://web.innodealing.com/quote-web/#/bond/sentiment-news-detail/area-news",
            "https://web.innodealing.com/quote-web/#/news/sentiment",
            "https://web.innodealing.com/quote-web/#/bond/yuqing",
            "https://web.innodealing.com/quote-web/#/news/yuqing",
        ]
        for url in sentiment_urls:
            await page.goto(url, wait_until="networkidle", timeout=15000)
            await page.wait_for_timeout(3000)
            title = await page.title()
            body_text = await page.locator("body").inner_text()
            text_preview = body_text[:300] if body_text else "(empty)"
            print(f"    {url.split('#')[-1][:50]:50s} → '{title}' | text: {text_preview[:100]}")

            # 搜索"早报"
            if "早报" in body_text:
                print(f"      ★ 找到'早报'!")
                full_text = body_text
                # 提取早报周围的内容
                idx = full_text.index("早报")
                context_start = max(0, idx - 500)
                context_end = min(len(full_text), idx + 3000)
                print(f"      {full_text[context_start:context_end]}")

        # 尝试通过 API 获取模块列表来找到舆情
        print("\n[4] 获取模块菜单...")
        resp = await page.evaluate("""async () => {
            const r = await fetch('/onshore-auth-service/api/pc/module/menu?' +
                'timestamp=' + Date.now() + '&v=13426&qIntl=0&sign=abc');
            return await r.json();
        }""")
        # 递归搜索"舆情"关键词
        def search_key(obj, path=""):
            results = []
            if isinstance(obj, dict):
                for k, v in obj.items():
                    if isinstance(v, str) and any(kw in v for kw in ['舆情', '早报', '信用', '资讯', '新闻']):
                        results.append(f"{path}.{k} = {v}")
                    results.extend(search_key(v, f"{path}.{k}"))
            elif isinstance(obj, list):
                for i, item in enumerate(obj):
                    results.extend(search_key(item, f"{path}[{i}]"))
            return results

        matches = search_key(resp)
        print(f"    找到 {len(matches)} 个相关项:")
        for m in matches[:30]:
            print(f"    {m[:200]}")

        await browser.close()

asyncio.run(main())
