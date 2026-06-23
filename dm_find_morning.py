"""遍历重要资讯页面的子 tab，找到 DM信用早报"""
import asyncio
import os
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"
OUT_DIR = os.path.dirname(os.path.abspath(__file__))

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

        # 导航到 public-opinion/important-news
        print("[2] Navigate to important-news...")
        await frame.evaluate("window.location.hash = '#/bond/public-opinion/important-news'")
        await page.wait_for_timeout(10000)
        body = await frame.locator("body").inner_text()
        print(f"    body: {len(body)} chars")

        # 找到所有子 tab 文本（利率时讯、债事特供、信用观点 等）
        tab_texts = ["重要资讯", "利率时讯", "债事特供", "AI研报", "DM专栏",
                     "债圈热议", "信用观点", "区域舆情", "公告", "公众号"]

        for tab_name in tab_texts:
            print(f"\n[3] Click tab: '{tab_name}'...")
            try:
                # 在 iframe 中找到并点击 tab
                tab = frame.locator(f"text={tab_name}").first
                if await tab.is_visible():
                    await tab.click()
                    await page.wait_for_timeout(5000)
                    new_body = await frame.locator("body").inner_text()

                    # 搜索早报
                    if "早报" in new_body or "DM信用" in new_body:
                        print(f"    ★★ FOUND in '{tab_name}' tab! ★★")
                        fpath = os.path.join(OUT_DIR, f"dm_zaobao_{tab_name}.txt")
                        with open(fpath, "w", encoding="utf-8") as f:
                            f.write(new_body)
                        # 提取早报相关内容
                        for line in new_body.split('\n'):
                            if '早报' in line or 'DM信用' in line:
                                print(f"    >>> {line.strip()}")
                    else:
                        # 检查有没有文章标题
                        lines = [l.strip() for l in new_body.split('\n') if l.strip()]
                        titles = [l for l in lines if 'DM' in l or '日报' in l or '早报' in l or '062' in l]
                        print(f"    Found {len(titles)} DM/日报 related lines (no '早报')")
                        for t in titles[:5]:
                            print(f"    - {t[:80]}")
                else:
                    print(f"    Tab not visible, trying text match...")
                    # 尝试模糊匹配
                    all_visible = await frame.locator("[role='tab'], .tab, .nav-item, li").all()
                    print(f"    Found {len(all_visible)} possible tab elements")
            except Exception as e:
                print(f"    Error: {e}")

        # 保存最终页面
        await browser.close()

asyncio.run(main())
