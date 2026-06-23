"""在 DM web 中导航到舆情→信用早报，并提取内容"""
import asyncio
import json
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # 捕获 API 响应
        api_responses = []
        async def handle_response(response):
            if "rest.innodealing.com" in response.url and response.status == 200:
                try:
                    body = await response.text()
                    if len(body) < 5000 and len(body) > 20:
                        api_responses.append({
                            "url": response.url,
                            "body": body[:2000]
                        })
                except:
                    pass

        page.on("response", handle_response)

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
        await page.wait_for_timeout(8000)

        print(f"    登录后 URL: {page.url}")

        # 在页面中搜索"舆情"
        print("\n[2] 在页面中搜索'舆情'...")
        yuqing_found = await page.locator("text=舆情").all()
        print(f"    找到 {len(yuqing_found)} 个'舆情'元素")
        for i, el in enumerate(yuqing_found[:10]):
            try:
                visible = await el.is_visible()
                text = await el.inner_text()
                tag = await el.evaluate("el => el.tagName")
                cls = await el.get_attribute("class")
                print(f"    [{i}] tag={tag}, class={cls}, visible={visible}, text='{text[:60]}'")
            except: pass

        # 搜索"信用早报"
        print("\n[3] 搜索'信用早报'或'早报'...")
        for kw in ["信用早报", "早报", "DM早报"]:
            found = await page.locator(f"text={kw}").all()
            print(f"    '{kw}': {len(found)} 个匹配")
        # Also search for "morning" in English
        morning = await page.locator("[class*='morning'], [id*='morning']").all()
        print(f"    [class/id *morning]: {len(morning)} 个匹配")

        # 打印所有匹配文本
        print("\n[4] 打印包含'信用'/'舆情'/'早报'/'credit'的 API 响应...")
        for resp in api_responses:
            url = resp['url']
            body = resp['body']
            for kw in ['信用', '舆情', '早报', 'credit', 'morning', 'news', 'sentiment']:
                if kw in body or kw in url:
                    print(f"  [{kw}] {url[:100]}")
                    print(f"    {body[:300]}")
                    print()
                    break

        # 截图
        await page.screenshot(path="dm_full.png", full_page=True)
        print("[5] 截图: dm_full.png")

        # 获取 cookies 保存
        cookies = await context.cookies()
        sid = next((c['value'] for c in cookies if c['name'] == 'sid'), None)
        print(f"\n[6] SID: {sid}")

        print("\n浏览器保持打开 60 秒，请手动探索...")
        await page.wait_for_timeout(60000)
        await browser.close()

asyncio.run(main())
