"""Complete DM article extraction pipeline — date-aware for daily automation.

Usage:
    python dm_pipeline.py                          # today's date
    python dm_pipeline.py --date 2026-06-16        # specific date
    python dm_pipeline.py --keyword "信用早报"     # explicit keyword (ignores date)
    python dm_pipeline.py --date 2026-06-16 --output article.txt
"""
import asyncio, json, os, re, sys, argparse
from datetime import date, timedelta
from playwright.async_api import async_playwright

# Credentials from environment or defaults
DM_USER = os.environ.get("DM_USER", "18611853878")
DM_PASS = os.environ.get("DM_PASS", "rbaggio5866058")
OUT = os.path.dirname(os.path.abspath(__file__))


def build_keywords(target_date=None, explicit=None):
    """Build search keywords for the given date.

    DM articles follow naming patterns like:
      - DM信用早报0616
      - 债市要闻速览0612
      - DM早报0616

    Returns list of keywords to try, in priority order.
    """
    if explicit:
        return [explicit]

    if target_date is None:
        target_date = date.today()

    mmdd = target_date.strftime("%m%d")
    md = target_date.strftime("%-m%d") if sys.platform != "win32" else target_date.strftime("%#m%d")

    keywords = [
        f"信用早报{mmdd}",      # DM信用早报0616
        f"早报{mmdd}",           # DM早报0616
        f"要闻速览{mmdd}",       # 债市要闻速览0612
        f"信用早报{md}",         # no zero-padding variant
        f"早报{md}",
        f"要闻速览{md}",
    ]
    # Deduplicate while preserving order
    seen = set()
    unique = []
    for kw in keywords:
        if kw not in seen:
            seen.add(kw)
            unique.append(kw)
    return unique


async def extract_article(keywords, output_file=None, find_all=False, datestr=None):
    """Login to DM, search for article matching any keyword, save to file.

    If find_all=True, extract ALL matching articles and return a list.
    Otherwise, extract only the first match.

    Returns:
        Single result: (success: bool, article_title: str, file_path: str, char_count: int)
        List result:   [(success, title, filepath, char_count), ...]
    """
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context()
        page = await ctx.new_page()

        api_data = []

        async def on_resp(response):
            url = response.url
            if "rest.innodealing.com" in url and response.status == 200:
                try:
                    ct = response.headers.get("content-type", "")
                    if "json" in ct:
                        body = await response.json()
                        if isinstance(body, dict):
                            api_data.append({"url": url, "data": body})
                except:
                    pass

        page.on("response", on_resp)

        # Step 1: Login
        print("[1] Login...")
        await page.goto("https://web.innodealing.com/dashboard/", wait_until="networkidle")
        await page.wait_for_timeout(3000)
        await page.fill("#inputUsername", DM_USER)
        await page.fill("#inputPassword", DM_PASS)
        try:
            privacy = page.locator("#approve-privacy")
            if await privacy.is_visible() and not await privacy.is_checked():
                await privacy.check()
        except:
            pass
        await page.click("button[type=submit]")
        await page.wait_for_timeout(15000)

        frame = page.frame("new-dashboard-frame")
        if not frame:
            print("ERROR: iframe not found")
            await browser.close()
            return False, "", "", 0
        print(f"[2] Iframe OK: {frame.url[:80]}")

        # Step 2: 直接跳转到 important-news 页面（整页加载，触发完整 API）
        print("[3] Navigating to important-news (full page load)...")
        await page.goto("https://web.innodealing.com/quote-web/#/bond/public-opinion/important-news",
                        wait_until="networkidle")
        await page.wait_for_timeout(8000)

        # 重新获取 iframe（整页加载后 iframe 可能已刷新）
        frame = page.frame("new-dashboard-frame")
        if not frame:
            # 尝试不带 iframe 的情况
            print("[3] Iframe not found after reload, trying body...")

        # Step 3: 诊断 — 打印所有截获的 API
        print(f"[3.5] 截获 {len(api_data)} 个 API 响应:")
        for item in api_data:
            url_short = item['url']
            if len(url_short) > 200:
                url_short = url_short[:200] + "..."
            print(f"      {url_short}")

        # Step 4: 在所有 API 响应中搜索文章列表
        articles = []
        for item in api_data:
            url = item["url"]
            data = item["data"]
            # 策略1: 搜索包含新闻列表的响应
            if isinstance(data, dict):
                # 尝试多个可能的数据路径
                for path in [
                    lambda d: d.get("data", {}).get("list", []),
                    lambda d: d.get("data", {}).get("records", []),
                    lambda d: d.get("data", {}).get("items", []),
                    lambda d: d.get("data", []),
                    lambda d: d.get("list", []),
                    lambda d: d.get("records", []),
                    lambda d: d.get("result", {}).get("list", []),
                ]:
                    try:
                        result = path(data)
                        if isinstance(result, list) and len(result) > 0:
                            # 检查第一条是否像文章（有 title 或 sentimentTitle）
                            first = result[0]
                            if isinstance(first, dict) and ("title" in first or "sentimentTitle" in first or "sentimentId" in first):
                                articles = result
                                print(f"      在 {url[:120]} 中找到 {len(articles)} 篇文章")
                                break
                    except:
                        pass
                if articles:
                    break

        print(f"[5] {len(articles)} articles from API")

        # Step 5: Search for articles matching keywords (collect ALL)
        matched = []  # list of (sentimentId, matched_kw, title)
        seen_ids = set()
        for kw in keywords:
            for a in articles:
                title = a.get("title", "") or a.get("sentimentTitle", "")
                sid = a.get("sentimentId", "")
                if sid and sid not in seen_ids and kw in title:
                    matched.append((sid, kw, title))
                    seen_ids.add(sid)
                    print(f"    MATCHED '{kw}': {sid} = {title[:100]}")

        if not matched:
            print("    Target not found. Available titles:")
            for a in articles[:20]:
                title = a.get("title", "") or a.get("sentimentTitle", "")
                sid = a.get("sentimentId", "")
                print(f"      {sid}: {title[:120]}")
            await browser.close()
            if find_all:
                return []
            return False, "", "", 0

        if not find_all:
            matched = matched[:1]

        # Step 6: Extract each matched article
        results = []
        content_target = frame if frame else page
        for idx, (target_id, matched_kw, article_title) in enumerate(matched):
            print(f"\n[6.{idx+1}/{len(matched)}] Opening article {target_id} — {article_title[:80]}")
            if frame:
                await frame.evaluate(
                    f"window.location.hash = '#/bond/sentiment-news-detail/area-news/{target_id}'"
                )
            else:
                await page.goto(
                    f"https://web.innodealing.com/quote-web/#/bond/sentiment-news-detail/area-news/{target_id}",
                    wait_until="networkidle"
                )
            await page.wait_for_timeout(10000)
            body = await content_target.locator("body").inner_text()

            # Determine output filename
            if find_all and len(matched) > 1:
                # Auto-name by article type
                if "要闻速览" in article_title:
                    fname = f"dm_yaowen_{datestr}.txt" if datestr else "dm_yaowen_output.txt"
                elif "信用早报" in article_title or "早报" in article_title:
                    fname = f"dm_zaobao_{datestr}.txt" if datestr else "dm_zaobao_output.txt"
                else:
                    fname = output_file or f"dm_article_{datestr}_{idx}.txt" if datestr else f"dm_article_output_{idx}.txt"
                filepath = os.path.join(OUT, fname)
            elif output_file:
                filepath = output_file
            elif datestr:
                filepath = os.path.join(OUT, f"dm_article_{datestr}.txt")
            else:
                filepath = os.path.join(OUT, "dm_article_output.txt")

            with open(filepath, "w", encoding="utf-8") as f:
                f.write(body)
            print(f"[6.{idx+1}] Saved: {filepath} ({len(body)} chars)")

            results.append((True, article_title, filepath, len(body)))

        await browser.close()

        if find_all:
            return results
        return results[0] if results else (False, "", "", 0)


def main():
    parser = argparse.ArgumentParser(description="Extract DM daily article")
    parser.add_argument("--date", help="Target date (YYYY-MM-DD), default=today")
    parser.add_argument("--keyword", help="Explicit search keyword (overrides date)")
    parser.add_argument("--output", help="Output file path")
    parser.add_argument("--all", dest="find_all", action="store_true",
                        help="Extract ALL matching articles (信用早报 + 要闻速览)")
    args = parser.parse_args()

    target_date = None
    if args.date:
        target_date = date.fromisoformat(args.date)
        date_str = target_date.strftime("%Y-%m-%d")
    else:
        target_date = date.today()
        date_str = target_date.strftime("%Y-%m-%d")

    keywords = build_keywords(target_date, args.keyword)
    datestr = target_date.strftime("%Y%m%d")
    print(f"Date: {date_str}  |  Keywords: {keywords}")

    if args.output:
        output_file = args.output
    else:
        output_file = os.path.join(OUT, f"dm_article_{datestr}.txt")

    result = asyncio.run(extract_article(
        keywords, output_file, find_all=args.find_all, datestr=datestr))

    if args.find_all:
        # result is a list of (success, title, filepath, chars)
        results = result
        if not results:
            print("\n[FAILED] No articles found for the given date/keyword.")
            marker = os.path.join(OUT, "dm_extract_status.txt")
            with open(marker, "w") as f:
                f.write("NOT_FOUND")
            sys.exit(1)

        print(f"\n[DONE] {len(results)} article(s) extracted:")
        marker_lines = []
        for success, title, fpath, chars in results:
            print(f"  {title} ({chars} chars) → {fpath}")
            marker_lines.append(f"OK|{title}|{fpath}|{chars}")
        marker = os.path.join(OUT, "dm_extract_status.txt")
        with open(marker, "w") as f:
            f.write("\n".join(marker_lines))
    else:
        success, title, fpath, chars = result
        if success:
            print(f"\n[DONE] {title} ({chars} chars)")
            print(f"File: {fpath}")
            marker = os.path.join(OUT, "dm_extract_status.txt")
            with open(marker, "w") as f:
                f.write(f"OK|{title}|{fpath}|{chars}")
        else:
            print("\n[FAILED] Article not found for the given date/keyword.")
            marker = os.path.join(OUT, "dm_extract_status.txt")
            with open(marker, "w") as f:
                f.write("NOT_FOUND")
            sys.exit(1)


if __name__ == "__main__":
    main()
