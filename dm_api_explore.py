"""登录 DM 并探测 REST API 接口"""
import asyncio
import json
from playwright.async_api import async_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

API_BASE = "https://rest.innodealing.com/onshore-bond-management/publicapi"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        # 拦截网络请求，记录 API 调用
        api_calls = []
        def log_request(request):
            if "rest.innodealing.com" in request.url:
                api_calls.append(f"{request.method} {request.url}")

        page.on("request", log_request)

        # 登录
        print("[1] 登录...")
        await page.goto("https://web.innodealing.com/dashboard/", wait_until="networkidle")
        await page.wait_for_timeout(3000)
        await page.fill("#inputUsername", DM_USER)
        await page.fill("#inputPassword", DM_PASS)
        try:
            privacy = page.locator("#approve-privacy")
            if await privacy.is_visible():
                await privacy.check()
        except:
            pass
        await page.click("button[type='submit']")
        await page.wait_for_timeout(8000)

        # 获取 cookies
        cookies = await context.cookies()
        print("\n[2] Cookies:")
        for c in cookies:
            if any(k in c['domain'] for k in ['innodealing', 'dm']):
                print(f"    {c['domain']:30s} {c['name']:20s} = {c['value'][:50]}{'...' if len(c['value'])>50 else ''}")

        # 尝试调用 REST API
        print("\n[3] 探测 REST API 端点...")
        token = None
        for c in cookies:
            if c['name'] in ('sid', 'token', 'access_token'):
                token = c['value']

        # 获取 auth header
        # 从 localStorage 获取 token
        local_storage = await page.evaluate("() => JSON.stringify(window.localStorage)")
        ls_data = json.loads(local_storage)
        for k, v in ls_data.items():
            if any(keyword in k.lower() for keyword in ['token', 'auth', 'user', 'sid']):
                print(f"    localStorage[{k}] = {str(v)[:80]}")

        session_storage = await page.evaluate("() => JSON.stringify(window.sessionStorage)")
        ss_data = json.loads(session_storage)
        for k, v in ss_data.items():
            if any(keyword in k.lower() for keyword in ['token', 'auth', 'user', 'sid']):
                print(f"    sessionStorage[{k}] = {str(v)[:80]}")

        # 打印捕获的 API 调用
        print(f"\n[4] 捕获到的 REST API 调用 ({len(api_calls)} 条):")
        for call in api_calls[:30]:
            print(f"    {call}")

        # 尝试用 cookie 直接调 API
        print("\n[5] 尝试直接调用 REST API...")
        headers = {
            "Content-Type": "application/json",
        }
        # 构建 cookie header
        cookie_header = "; ".join(f"{c['name']}={c['value']}" for c in cookies if 'innodealing' in c['domain'])
        if cookie_header:
            headers["Cookie"] = cookie_header

        # 尝试几个可能的端点
        endpoints = [
            f"{API_BASE}/user/info",
            f"{API_BASE}/auth/token",
            f"{API_BASE}/menu/list",
            f"{API_BASE}/news/list",
            f"{API_BASE}/report/list",
            f"{API_BASE}/sentiment/list",
        ]
        import requests
        for url in endpoints:
            try:
                resp = requests.get(url, headers=headers, timeout=10)
                print(f"    {url.split('/')[-1]:20s} → HTTP {resp.status_code} | {resp.text[:100]}")
            except Exception as e:
                print(f"    {url.split('/')[-1]:20s} → Error: {e}")

        await browser.close()

asyncio.run(main())
