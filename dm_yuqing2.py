"""通过 SPA 内部导航到舆情页面"""
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

        # 尝试通过 JS 找到舆情入口
        print("\n[2] 查找页面中的导航元素...")

        # 打印所有可见的文本内容（用于调试）
        all_text = await page.evaluate("""() => {
            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
            const texts = [];
            let node;
            while (node = walker.nextNode()) {
                const text = node.textContent.trim();
                if (text && text.length > 1 && text.length < 50) {
                    texts.push(text);
                }
            }
            return texts.slice(0, 200);
        }""")
        print(f"    页面文本节点 ({len(all_text)} 个):")
        for i, t in enumerate(all_text):
            if any(c in t for c in ['舆情','新闻','信用','资讯','早报','每日','DM','债券','分析','风控','市场','关注','报告']):
                print(f"    [{i}] {t}")

        # 尝试用 JS Hash 路由导航
        print("\n[3] 尝试通过 hash 导航到舆情...")
        await page.evaluate("window.location.hash = '#/bond/public-opinion/important-news'")
        await page.wait_for_timeout(10000)

        body = await page.locator("body").inner_text()
        print(f"    页面文本长度: {len(body)}")

        # 查找"早报"等关键词
        for kw in ["早报", "信用", "DM早报", "DM信用早报", "重要资讯"]:
            idx = body.find(kw)
            if idx >= 0:
                print(f"    ★ '{kw}' 在位置 {idx}")
                print(f"      上下文: {body[max(0,idx-50):idx+200]}")
            else:
                print(f"    '{kw}': 未找到")

        # 截图看实际页面
        await page.screenshot(path="dm_yuqing2.png", full_page=True)
        print("\n[4] 截图: dm_yuqing2.png")

        # 也尝试 promotion 中的 sentiment 链接
        print("\n[5] 尝试 sentiment-news-detail 页面...")
        # 先通过 API 拿到最新的 area-news ID
        resp_text = await page.evaluate("""async () => {
            const r = await fetch('/onshore-platform-service/api/promote/screen/list/by/userid?' +
                'timestamp=' + Date.now() + '&v=13426&qIntl=0&sign=abc');
            // sign 不对但可以看 cookie 是否生效
            return r.status;
        }""")
        print(f"    promote API status: {resp_text}")

        await browser.close()

asyncio.run(main())
