"""Headed 模式导航 DM—带窗口渲染，直接点击 tab 元素"""
import asyncio
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

async def main():
    async with async_playwright() as p:
        # HEADED 模式 — 关键改动
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(
            viewport={"width": 1920, "height": 1080}
        )
        page = await context.new_page()

        # 收集 API
        api_urls = []
        def log_req(request):
            if "rest.innodealing.com" in request.url:
                api_urls.append(request.url)
        page.on("request", log_req)

        # 登录
        print("[1] 登录...")
        await page.goto("https://web.innodealing.com/dashboard/", wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(3000)
        await page.fill("#inputUsername", DM_USER)
        await page.fill("#inputPassword", DM_PASS)
        try:
            privacy = page.locator("#approve-privacy")
            if await privacy.is_visible(): await privacy.check()
        except: pass
        await page.click("button[type='submit']")
        await page.wait_for_timeout(10000)

        print(f"    URL: {page.url}")
        body_text = await page.locator("body").inner_text()
        print(f"    body 长度: {len(body_text)}")

        # 查找 tab 元素
        print("\n[2] 查找导航 tab...")
        tabs = await page.locator(".chrome-tab").all()
        print(f"    找到 {len(tabs)} 个 chrome-tab")
        for i, tab in enumerate(tabs):
            try:
                title_el = tab.locator(".chrome-tab-title")
                title = await title_el.inner_text() if await title_el.count() > 0 else "(no title)"
                cls = await tab.get_attribute("class")
                is_current = "current" in cls if cls else False
                print(f"    [{i}] current={is_current} title='{title}' class='{cls}'")
            except Exception as e:
                print(f"    [{i}] error: {e}")

        # 查找所有链接和可点击元素
        print("\n[3] 查找页面中的文本元素...")
        all_text = await page.evaluate("""() => {
            const all = document.querySelectorAll('*');
            const texts = [];
            all.forEach(el => {
                const text = (el.textContent || '').trim();
                if (text.length > 0 && text.length < 100 && el.children.length === 0) {
                    texts.push({
                        tag: el.tagName,
                        text: text,
                        className: el.className,
                        visible: el.offsetParent !== null
                    });
                }
            });
            return texts.slice(0, 100);
        }""")

        for item in all_text:
            if any(kw in item['text'] for kw in ['舆情','新闻','信用','资讯','早报','DM','债券','重要','日报']):
                print(f"    [{item['tag']}] visible={item['visible']} text='{item['text']}' class='{item['className']}'")

        # 截图看页面状态
        await page.screenshot(path="dm_headed.png", full_page=False)
        print("\n[4] 截图: dm_headed.png")

        # 关闭浏览器（让用户看到结果）
        print("\n浏览器保持打开15秒...")
        await page.wait_for_timeout(15000)
        await browser.close()

asyncio.run(main())
