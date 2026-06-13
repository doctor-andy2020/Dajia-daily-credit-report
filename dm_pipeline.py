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


async def extract_article(keywords, output_file=None):
    """Login to DM, search for article matching any keyword, save to file.

    Returns (success: bool, article_title: str, file_path: str, char_count: int)
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
        api_data.clear()

        # Step 2: Load article list
        print("[3] Loading important-news...")
        await frame.evaluate("window.location.hash = '#/bond/public-opinion/important-news'")
        await page.wait_for_timeout(12000)

        # Step 3: Extract articles from API
        articles = []
        for item in api_data:
            if "sentiment/news/paging" in item["url"]:
                articles = item["data"].get("data", {}).get("list", [])
                break
        if not articles:
            for item in api_data:
                if "headline/list" in item["url"]:
                    articles = item["data"].get("data", [])
                    break

        print(f"[4] {len(articles)} articles from API")

        # Step 4: Search for article matching keywords
        target_id = None
        matched_kw = ""
        for kw in keywords:
            for a in articles:
                title = a.get("title", "") or a.get("sentimentTitle", "")
                sid = a.get("sentimentId", "")
                if sid and kw in title:
                    target_id = sid
                    matched_kw = kw
                    print(f"    MATCHED '{kw}': {sid} = {title[:100]}")
                    break
            if target_id:
                break

        if not target_id:
            print("    Target not found. Available titles:")
            for a in articles[:20]:
                title = a.get("title", "") or a.get("sentimentTitle", "")
                sid = a.get("sentimentId", "")
                print(f"      {sid}: {title[:120]}")
            await browser.close()
            return False, "", "", 0

        # Step 5: Open article detail
        print(f"\n[5] Opening article {target_id}...")
        await frame.evaluate(
            f"window.location.hash = '#/bond/sentiment-news-detail/area-news/{target_id}'"
        )
        await page.wait_for_timeout(10000)
        body = await frame.locator("body").inner_text()

        # Step 6: Save
        article_title = ""
        for a in articles:
            sid = a.get("sentimentId", "")
            if sid == target_id:
                article_title = a.get("title", "") or a.get("sentimentTitle", "")
                break

        if output_file is None:
            output_file = os.path.join(OUT, "dm_article_output.txt")

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(body)
        print(f"[6] Saved: {output_file} ({len(body)} chars)")
        print(f"    Title: {article_title}")

        await browser.close()
        return True, article_title, output_file, len(body)


def main():
    parser = argparse.ArgumentParser(description="Extract DM daily article")
    parser.add_argument("--date", help="Target date (YYYY-MM-DD), default=today")
    parser.add_argument("--keyword", help="Explicit search keyword (overrides date)")
    parser.add_argument("--output", help="Output file path")
    args = parser.parse_args()

    target_date = None
    if args.date:
        target_date = date.fromisoformat(args.date)
        date_str = target_date.strftime("%Y-%m-%d")
    else:
        target_date = date.today()
        date_str = target_date.strftime("%Y-%m-%d")

    keywords = build_keywords(target_date, args.keyword)
    print(f"Date: {date_str}  |  Keywords: {keywords}")

    if args.output:
        output_file = args.output
    else:
        output_file = os.path.join(OUT, f"dm_article_{target_date.strftime('%Y%m%d')}.txt")

    success, title, fpath, chars = asyncio.run(extract_article(keywords, output_file))

    if success:
        print(f"\n[DONE] {title} ({chars} chars)")
        print(f"File: {fpath}")
        # Write success marker for CI
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
