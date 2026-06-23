"""检查 DM 页面 iframe 结构"""
import asyncio
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

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

        # 保存 HTML
        html = await page.content()
        with open("dm_dashboard_full.html", "w", encoding="utf-8") as f:
            f.write(html)

        # 查找 iframe
        iframes = page.frames
        print(f"\n[2] 页面 frames ({len(iframes)}):")
        for i, frame in enumerate(iframes):
            print(f"    [{i}] url={frame.url[:120]}")
            print(f"        name={frame.name}")

        # 如果有嵌套 frame，逐个进入
        for i, frame in enumerate(iframes):
            if i == 0:
                continue  # skip main frame
            try:
                text = await frame.locator("body").inner_text()
                print(f"\n    Frame [{i}] body: {len(text)} chars")
                if text:
                    print(f"        preview: {text[:300]}")
            except Exception as e:
                print(f"        error: {e}")

        # 直接用 JS 查找所有 iframe src
        iframe_srcs = await page.evaluate("""() => {
            return Array.from(document.querySelectorAll('iframe')).map(f => ({
                src: f.src,
                id: f.id,
                className: f.className,
                visible: f.offsetParent !== null
            }));
        }""")
        print(f"\n[3] iframe 元素: {json.dumps(iframe_srcs, indent=2, ensure_ascii=False)}")

        await browser.close()

import json
asyncio.run(main())
