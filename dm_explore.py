"""探索 DM web 版，找到舆情和信用早报的页面结构和 API"""
import asyncio
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # 1. 打开登录页
        print("[1] 打开登录页...")
        await page.goto("https://web.innodealing.com/dashboard/", wait_until="networkidle")
        await page.wait_for_timeout(2000)

        # 打印页面标题
        print(f"    页面标题: {await page.title()}")

        # 2. 填写登录表单
        print("[2] 填写登录信息...")
        # 等待登录表单加载
        await page.wait_for_selector("#inputUsername", timeout=10000)
        await page.fill("#inputUsername", DM_USER)
        await page.fill("#inputPassword", DM_PASS)

        # 勾选隐私协议 (如果存在)
        try:
            privacy = page.locator("#approve-privacy")
            if await privacy.is_visible():
                await privacy.check()
                print("    已勾选隐私协议")
        except:
            pass

        # 3. 点击登录
        print("[3] 点击登录...")
        await page.click("button[type='submit']")

        # 等待页面跳转
        await page.wait_for_timeout(5000)
        print(f"    登录后 URL: {page.url}")
        print(f"    登录后标题: {await page.title()}")

        # 4. 截图当前页面
        await page.screenshot(path="dm_dashboard.png", full_page=False)
        print("[4] 截图已保存: dm_dashboard.png")

        # 5. 尝试找到舆情入口
        print("[5] 查找导航菜单...")
        # 打印所有导航链接
        nav_links = await page.locator("a, [role='menuitem'], .nav-item, .menu-item, .sidebar-item, li span, .ant-menu-item").all()
        for i, link in enumerate(nav_links[:30]):
            try:
                text = await link.inner_text()
                href = await link.get_attribute("href")
                if text.strip():
                    print(f"    [{i}] {text.strip():30s} | href: {href}")
            except:
                pass

        # 6. 保存页面HTML供分析
        html = await page.content()
        with open("dm_dashboard.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("[6] HTML 已保存: dm_dashboard.html")

        # 保持浏览器打开30秒供观察
        print("\n浏览器保持打开，30秒后关闭...")
        await page.wait_for_timeout(30000)
        await browser.close()

asyncio.run(main())
