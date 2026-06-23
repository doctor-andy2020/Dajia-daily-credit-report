"""通过 Playwright 点击 DM 网页导航，触发 SPA 自行加载数据"""
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
                    if "json" in ct and len(response.url) > 30:
                        body = await response.json()
                        api_data.append({"url": url, "data": body})
                except: pass

        page.on("response", on_response)

        # Step 1: 登录
        print("[1] 登录...")
        await page.goto("https://web.innodealing.com/dashboard/", wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
        await page.fill("#inputUsername", DM_USER)
        await page.fill("#inputPassword", DM_PASS)
        try:
            privacy = page.locator("#approve-privacy")
            if await privacy.is_visible(): await privacy.check()
        except: pass
        await page.click("button[type='submit']")
        await page.wait_for_timeout(8000)

        print(f"    登录后 URL: {page.url}")
        print(f"    标题: {await page.title()}")

        # Step 2: 尝试直接 hash 导航，但用不同的 wait_until
        print("\n[2] Hash 导航到舆情...")
        yuqing_url = "https://web.innodealing.com/dashboard/#/bond/public-opinion/important-news"
        try:
            await page.goto(yuqing_url, wait_until="domcontentloaded", timeout=20000)
            await page.wait_for_timeout(5000)
        except Exception as e:
            print(f"    goto 超时: {e}")
            # 继续 — SPA 可能还在加载

        body = await page.locator("body").inner_text()
        print(f"    页面 body 长度: {len(body)}")

        if len(body) > 100:
            print(f"    前300字: {body[:300]}")
        else:
            print("    页面内容为空，尝试其他方式...")

        # Step 3: 查找包含关键词的 API 响应
        print(f"\n[3] 收集到 {len(api_data)} 个新 API 响应")
        for i, item in enumerate(api_data):
            data_str = json.dumps(item['data'], ensure_ascii=False)
            # 搜索文章/新闻相关内容
            for kw in ['早报','信用','DM','title','content','article','news','list','重要']:
                if kw in data_str or kw in item['url']:
                    url_short = item['url'].split('rest.innodealing.com')[-1][:100]
                    print(f"\n  [{i}] [{kw}] {url_short}")
                    # 截断显示，找标题列表
                    data_preview = data_str[:1000]
                    print(f"    {data_preview}")
                    break

        # Step 4: 保存截图
        await page.screenshot(path="dm_after_hashnav.png", full_page=True)
        print("\n[4] 截图: dm_after_hashnav.png")

        # Step 5: 如果还是空的，尝试用 Vue/React 路由
        if len(body) < 100:
            print("\n[5] 尝试通过 SPA 路由...")
            # 打印页面中存在的元素，帮助诊断
            classes = await page.evaluate("""() => {
                const all = document.querySelectorAll('*');
                const classes = new Set();
                all.forEach(el => {
                    if (el.className && typeof el.className === 'string') {
                        el.className.split(' ').forEach(c => { if (c) classes.add(c); });
                    }
                });
                return Array.from(classes).slice(0, 50);
            }""")
            print(f"    页面 CSS 类 (前50): {classes}")

            # 尝试执行 SPA 内部导航
            nav_result = await page.evaluate("""() => {
                // 尝试 React Router
                if (window.__REACT_HISTORY__) {
                    window.__REACT_HISTORY__.push('/bond/public-opinion/important-news');
                    return 'React push done';
                }
                // 尝试 dispatch popstate
                window.dispatchEvent(new PopStateEvent('popstate'));
                return 'popstate dispatched';
            }""")
            print(f"    导航结果: {nav_result}")

            await page.wait_for_timeout(5000)
            body2 = await page.locator("body").inner_text()
            print(f"    新 body 长度: {len(body2)}")

        await browser.close()

asyncio.run(main())
