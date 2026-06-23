"""通过 REST API 获取新闻文章内容"""
import requests
import json
from playwright.sync_api import sync_playwright

DM_USER = "18611853878"
DM_PASS = "rbaggio5866058"

# 从 promote API 获取的 article ID
ARTICLE_ID = "2026061200010199544"  # June 12 area-news article

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        # 登录
        print("[1] 登录获取 cookie...")
        page.goto("https://web.innodealing.com/dashboard/", wait_until="networkidle")
        page.wait_for_timeout(3000)
        page.fill("#inputUsername", DM_USER)
        page.fill("#inputPassword", DM_PASS)
        try:
            privacy = page.locator("#approve-privacy")
            if privacy.is_visible(): privacy.check()
        except: pass
        page.click("button[type='submit']")
        page.wait_for_timeout(10000)

        # 获取 cookies
        cookies = page.context.cookies()
        cookie_dict = {c['name']: c['value'] for c in cookies}
        cookie_header = "; ".join(f"{c['name']}={c['value']}" for c in cookies)

        # 获取 localStorage/sessionStorage tokens
        token = page.evaluate("() => localStorage.getItem('token')")
        access_token = page.evaluate("() => localStorage.getItem('access_token')")
        print(f"    token: {token}")
        print(f"    access_token: {access_token}")

        # 尝试各种 API 模式获取文章
        headers = {
            "Cookie": cookie_header,
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        }
        if token:
            headers["Authorization"] = f"Bearer {token}"

        # 多个可能的 API 端点
        api_patterns = [
            # 按文章 ID 获取详情
            f"https://rest.innodealing.com/onshore-bond-management/publicapi/sentiment/news/detail?id={ARTICLE_ID}",
            f"https://rest.innodealing.com/onshore-bond-management/api/sentiment/news/detail?id={ARTICLE_ID}",
            f"https://rest.innodealing.com/onshore-bond-management/publicapi/news/detail/{ARTICLE_ID}",
            f"https://rest.innodealing.com/bond-web/api/sentiment/news/detail?id={ARTICLE_ID}",
            f"https://rest.innodealing.com/onshore-bond-management/api/sentiment/detail?id={ARTICLE_ID}",
            # 文章列表
            f"https://rest.innodealing.com/onshore-bond-management/publicapi/sentiment/news/list",
            f"https://rest.innodealing.com/onshore-bond-management/api/sentiment/news/list",
            # 按模块代码获取
            f"https://rest.innodealing.com/onshore-bond-management/publicapi/news/list?moduleCode=10000046",
            f"https://rest.innodealing.com/onshore-bond-management/api/news/list?moduleCode=10000046",
            # area-news 专用
            f"https://rest.innodealing.com/onshore-bond-management/publicapi/sentiment/area-news/list",
            f"https://rest.innodealing.com/onshore-bond-management/api/sentiment/area-news/list",
        ]

        print("\n[2] 尝试各种 API 端点...")
        for url in api_patterns:
            try:
                resp = requests.get(url, headers=headers, timeout=15)
                status = resp.status_code
                size = len(resp.text)
                preview = resp.text[:200] if resp.text else "(empty)"
                if status == 200 and size > 100:
                    print(f"  ★ HTTP {status} | {len(resp.text)}B | {url.split('/')[-1][:60]}")
                    print(f"    {preview}")
                else:
                    print(f"    HTTP {status} | {size}B | {url.split('/')[-1][:60]}")
            except Exception as e:
                print(f"    ERROR | {url.split('/')[-1][:60]} | {e}")

        # 也尝试通过页面 JS fetch
        print("\n[3] 通过页面 fetch 获取文章...")
        for url_suffix in [
            f"/onshore-bond-management/publicapi/sentiment/news/detail?id={ARTICLE_ID}",
            f"/onshore-bond-management/api/sentiment/news/detail?id={ARTICLE_ID}",
            f"/bond-web/api/sentiment/news/detail?id={ARTICLE_ID}",
        ]:
            result = page.evaluate(f"""async () => {{
                try {{
                    const r = await fetch('https://rest.innodealing.com{url_suffix}', {{
                        headers: {{'Accept': 'application/json'}},
                        credentials: 'include'
                    }});
                    return await r.text();
                }} catch(e) {{ return 'Error: ' + e.message; }}
            }}""")
            print(f"    {url_suffix}")
            print(f"    {result[:300]}")

        browser.close()

main()
