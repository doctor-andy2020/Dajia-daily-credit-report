"""在 Dashboard 中拦截所有 API 响应，定位舆情/信用早报的数据源"""
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

        api_responses = []

        async def on_response(response):
            url = response.url
            if "rest.innodealing.com" in url:
                try:
                    ct = response.headers.get("content-type", "")
                    if "json" in ct:
                        body = await response.json()
                        api_responses.append({"url": url, "data": body})
                except:
                    pass

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
        await page.wait_for_timeout(15000)  # 等久一点，让所有模块加载

        # 搜索 HTML 中的 API 路径和模块信息
        html = await page.content()
        # 找所有 JS 文件
        js_files = re.findall(r'(?:src|href)="([^"]+\.js)"', html)

        # 找 window 变量中的配置
        window_vars = re.findall(r'window\.(\w+)\s*=\s*["\']([^"\']+)["\']', html)
        print("\n[2] window 配置变量:")
        for var, val in window_vars:
            print(f"    window.{var} = {val}")

        # 分析所有 API 响应
        print(f"\n[3] 共捕获 {len(api_responses)} 个 API 响应")

        # 打印所有包含 sentiment/news/舆情/早报 的响应
        print("\n[4] 关键 API 响应:")
        for item in api_responses:
            url = item['url']
            data_str = json.dumps(item['data'], ensure_ascii=False)
            keywords = ['sentiment', 'news', 'morning', 'article', 'yuqing',
                       '舆情', '早报', '信用', 'report', 'daily', 'area']
            for kw in keywords:
                if kw in url.lower() or kw.lower() in data_str.lower():
                    # 截断长响应
                    display = data_str[:800]
                    print(f"\n  [{kw}] {url}")
                    print(f"  {display}")
                    break

        # 找 promote/screen API（这个之前有 sentiment 链接）
        print("\n[5] Promote/Screen API 响应:")
        for item in api_responses:
            if 'promote' in item['url']:
                data_str = json.dumps(item['data'], ensure_ascii=False, indent=2)
                print(f"  {item['url']}")
                print(f"  {data_str[:2000]}")

        # 找菜单/模块 API
        print("\n[6] 模块菜单 API 响应:")
        for item in api_responses:
            if 'module' in item['url'] or 'menu' in item['url']:
                data_str = json.dumps(item['data'], ensure_ascii=False, indent=2)
                print(f"  {item['url']}")
                # 递归搜索舆情相关
                found = []
                def find_kw(obj, path=""):
                    if isinstance(obj, dict):
                        for k, v in obj.items():
                            if isinstance(v, str) and any(x in v for x in ['舆情', '早报', '信用', '新闻', '资讯', 'sentiment', 'news']):
                                found.append(f"{path}.{k} = {v[:100]}")
                            find_kw(v, f"{path}.{k}")
                    elif isinstance(obj, list):
                        for i, v in enumerate(obj):
                            find_kw(v, f"{path}[{i}]")
                find_kw(item['data'])
                for f in found[:20]:
                    print(f"    {f}")

        await browser.close()

asyncio.run(main())
