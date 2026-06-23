"""深度搜索 DM API — 保存所有响应到文件，搜索信用早报"""
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

        print(f"    已登录, 收集到 {len(all_responses)} 个 API 响应")

        # 尝试导航到公共舆情页面
        print("\n[2] 尝试导航...")
        # 方法1: 通过 window.location
        result = await page.evaluate("""
            window.location.hash = '#/bond/public-opinion/important-news';
            return 'done';
        """)
        await page.wait_for_timeout(10000)
        print(f"    新 URL: {page.url}")
        print(f"    总响应数: {len(all_responses)}")

        # 保存所有响应到文件
        with open("dm_api_dump.json", "w", encoding="utf-8") as f:
            json.dump(all_responses, f, ensure_ascii=False, indent=2)
        print(f"\n[3] 已保存 {len(all_responses)} 个响应到 dm_api_dump.json")

        # 搜索关键内容
        print("\n[4] 搜索 API 响应中的关键词...")
        keywords = [
            "信用早报", "DM早报", "早报", "信用日报",
            "morning report", "daily credit",
            "每日信用", "DM信用", "CREDIT MORNING"
        ]
        for kw in keywords:
            found = []
            for i, resp in enumerate(all_responses):
                data_str = json.dumps(resp['data'], ensure_ascii=False)
                if kw.lower() in data_str.lower():
                    found.append((i, resp['url']))
            if found:
                print(f"  ★ '{kw}' 在 {len(found)} 个响应中找到:")
                for idx, url in found[:5]:
                    print(f"      [{idx}] {url[:120]}")
            else:
                print(f"    '{kw}': 未找到")

        # 打印 tabs 配置（包含舆情模块信息）
        print("\n[5] Tabs 配置:")
        for resp in all_responses:
            if 'tabs' in resp['url'] and 'global_tabs_config' in resp['url']:
                tabs_data = json.loads(resp['data']['data'])
                for tab in tabs_data.get('tabs', []):
                    if any(kw in str(tab) for kw in ['舆情', 'PublicOpinion', '信用', '早报', 'DM']):
                        print(f"    {json.dumps(tab, ensure_ascii=False)}")

        # 打印 promote/screen 数据（包含 sentiment 链接）
        print("\n[6] Promote/Screen 数据:")
        for resp in all_responses:
            if 'promote/screen' in resp['url']:
                for item in resp['data'].get('data', []):
                    detail = item.get('detailUrl', '')
                    if 'sentiment' in detail or 'area-news' in detail:
                        print(f"    detailUrl: {detail}")

        await browser.close()

asyncio.run(main())
