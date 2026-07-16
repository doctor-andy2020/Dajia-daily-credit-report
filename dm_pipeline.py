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

        # Step 2: 先打印登录阶段截获的所有 API
        print(f"[2.5] 登录阶段截获 {len(api_data)} 个 API:")
        for item in api_data:
            print(f"      {item['url'][:200]}")

        # Step 3: Load article list
        print("[3] Loading important-news...")
        await frame.evaluate("window.location.hash = '#/bond/public-opinion/important-news'")
        await page.wait_for_timeout(15000)

        # Step 4: 打印所有截获的 API（不限于特定 URL）
        print(f"[4] 总计截获 {len(api_data)} 个 API 响应")
        for i, item in enumerate(api_data):
            url_short = item['url'][:200]
            print(f"    [{i}] {url_short}")

        # Step 5: 搜索文章列表（优先 sentiment/news/paging，其次 headline/list，最后扫描全部）
        def try_extract(item):
            data = item["data"]
            if not isinstance(data, dict):
                return []
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
                        first = result[0]
                        if isinstance(first, dict) and ("title" in first or "sentimentTitle" in first):
                            return result
                except:
                    pass
            return []

        # ── 从所有 API 源收集文章（反向搜索，hash 导航后的新响应优先）──
        # 登录后 iframe 可能先停在 issuer 详情页（返回 issuer 专属舆情），
        # hash 导航到 important-news 后触发新 API 追加到列表末尾。
        # 反向搜索 + 多源收集 确保始终拿到正确的通用舆情数据。
        all_articles = []  # 所有候选文章
        collected_urls = set()

        for source_label, url_pattern in [
            ("sentiment/news/paging", "sentiment/news/paging"),
            ("headline/list", "headline/list"),
        ]:
            for item in reversed(api_data):
                url_key = item["url"].split("?")[0]
                if url_pattern in item["url"] and url_key not in collected_urls:
                    articles = try_extract(item)
                    if articles:
                        print(f"      [OK] {source_label}: {len(articles)} 篇")
                        all_articles.extend(articles)
                        collected_urls.add(url_key)
                        break  # 每种源只取最新一次响应

        # 去重（按 sentimentId）
        seen_ids = set()
        articles = []
        for a in all_articles:
            sid = a.get("sentimentId", "") or a.get("id", "") or a.get("newsId", "")
            if sid and sid not in seen_ids:
                articles.append(a)
                seen_ids.add(sid)

        print(f"[5] {len(articles)} articles from API (combined from {len(collected_urls)} sources)")

        # ── 诊断：如果文章数为 0，dump API 响应结构以便定位新数据路径 ──
        if len(articles) == 0:
            print("[诊断] 文章列表为空，dump API 响应结构：")
            for i, item in enumerate(api_data):
                url_short = item["url"][:150]
                data = item["data"]
                if isinstance(data, dict):
                    # 递归打印前两层 key 结构
                    def dump_structure(d, depth=0):
                        if not isinstance(d, dict):
                            return f"<{type(d).__name__}> (len={len(d) if isinstance(d, (list,str)) else '?'})"
                        lines = []
                        for k, v in d.items():
                            prefix = "  " * (depth + 1)
                            if isinstance(v, dict):
                                lines.append(f"{prefix}{k}: dict[{', '.join(list(v.keys())[:5])}...]" if len(v)>5 else f"{prefix}{k}: dict[{', '.join(v.keys())}]")
                            elif isinstance(v, list):
                                lines.append(f"{prefix}{k}: list[{len(v)}]")
                                if len(v) > 0 and isinstance(v[0], dict):
                                    lines.append(f"{prefix}  [0]: dict[{', '.join(list(v[0].keys())[:8])}]")
                            else:
                                lines.append(f"{prefix}{k}: {type(v).__name__} = {str(v)[:80]}")
                        return "\n".join(lines)
                    print(f"  [{i}] {url_short}")
                    print(dump_structure(data))
                    print()
                else:
                    print(f"  [{i}] {url_short} — type={type(data).__name__}")
            print("[诊断] 结构 dump 完毕。")

        # Step 5: Search for articles matching keywords (collect ALL)
        matched = []  # list of (sentimentId, matched_kw, title)
        seen_ids = set()
        for kw in keywords:
            for a in articles:
                title = a.get("title", "") or a.get("sentimentTitle", "")
                sid = a.get("sentimentId", "")
                # 忽略空格匹配：DM 标题格式可能变化（"DM信用早报0716" vs "DM信用早报 0716"）
                if sid and sid not in seen_ids and kw.replace(" ", "") in title.replace(" ", ""):
                    matched.append((sid, kw, title))
                    seen_ids.add(sid)
                    print(f"    MATCHED '{kw}': {sid} = {title[:100]}")

        if not matched and articles:
            print("    Target not found. Available titles:")
            for a in articles[:20]:
                title = a.get("title", "") or a.get("sentimentTitle", "")
                sid = a.get("sentimentId", "")
                print(f"      {sid}: {title[:120]}")

        if not matched:
            await browser.close()
            if find_all:
                return []
            return False, "", "", 0

        if not find_all:
            matched = matched[:1]

        # Step 6: Extract each matched article
        results = []
        for idx, (target_id, matched_kw, article_title) in enumerate(matched):
            print(f"\n[6.{idx+1}/{len(matched)}] Opening article {target_id} — {article_title[:80]}")
            await frame.evaluate(
                f"window.location.hash = '#/bond/sentiment-news-detail/area-news/{target_id}'"
            )
            await page.wait_for_timeout(10000)
            body = await frame.locator("body").inner_text()

            # Determine output filename — always use type-based naming with --all
            if find_all:
                # Auto-name by article type (even for single match)
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
