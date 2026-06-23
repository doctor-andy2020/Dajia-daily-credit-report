#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
/daily 统一入口脚本
串联：邮件拉取 → HTML 解析 → 报告生成 → 邮件发送
周一拉取周五+周六+周日三天邮件合并，周二至周五拉取前一天邮件
用法：python daily_runner.py
"""

import subprocess
import sys
import os
import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent

# 固定使用北京时间，避免 GitHub Actions UTC 时区导致的日期错位
BEIJING_TZ = datetime.timezone(datetime.timedelta(hours=8))


def beijing_today():
    """返回北京时间今天的 date 对象"""
    return datetime.datetime.now(BEIJING_TZ).date()


def get_target_dates():
    """根据北京时间的星期几，返回需要拉取的邮件日期列表"""
    today = beijing_today()
    weekday = today.weekday()  # 0=Mon, 6=Sun

    if weekday == 0:  # 周一：拉周六、周日、周一（IMAP日期=新闻日期+1）
        return [today - datetime.timedelta(days=i) for i in (2, 1, 0)]
    elif weekday in (5, 6):  # 周六日：不运行
        return []
    else:  # 周二至周五：拉当天（IMAP日期=新闻日期+1）
        return [today]


def run_step(description, cmd):
    print(f"\n{'='*60}")
    print(f">>> {description}")
    print(f"{'='*60}")
    result = subprocess.run(cmd, shell=True, cwd=str(BASE_DIR),
                           capture_output=False)
    if result.returncode != 0:
        print(f"[失败] {description} (exit code: {result.returncode})")
        return False
    return True


def main():
    import datetime as dt
    today = beijing_today()
    target_dates = get_target_dates()

    print("=" * 60)
    print("  大家资产持仓信用主体舆情日报 — 自动化工作流")
    print(f"  启动时间：{dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  今日星期：{['一','二','三','四','五','六','日'][today.weekday()]}")
    print(f"  拉取日期：{', '.join(d.strftime('%Y-%m-%d') for d in target_dates)}")
    print("=" * 60)

    force_run = "--force" in sys.argv or "--force-weekend" in sys.argv

    # ── 时间窗口限制：仅在北京时间 07:00-11:00 执行 ──
    now_bj = dt.datetime.now(BEIJING_TZ)
    hour_bj = now_bj.hour
    if not force_run and not (7 <= hour_bj < 11):
        print(f"[跳过] 当前北京时间 {now_bj.strftime('%H:%M')}，不在执行窗口(07:00-11:00)内。")
        print(f"       使用 --force 可强制运行。")
        sys.exit(0)

    # ── 去重检查：今天已成功生成过报告则跳过 ──
    sentinel = BASE_DIR / f".daily_sentinel_{today.strftime('%Y%m%d')}"
    if sentinel.exists() and not force_run:
        print(f"[跳过] 今日报告已生成（sentinel: {sentinel}），使用 --force 强制重跑。")
        sys.exit(0)

    if not target_dates and not force_run:
        print("[跳过] 周末不执行（使用 --force 强制运行）。")
        sys.exit(0)

    if force_run and not target_dates:
        print("[强制模式] 周末运行，拉取最近一个工作日数据。")
        # Find last weekday
        d = today
        while d.weekday() >= 5:
            d = d - datetime.timedelta(days=1)
        target_dates = [d]

    fetch_script = BASE_DIR / 'fetch_gmail_dm.py'
    if not fetch_script.exists():
        print("[错误] 找不到 fetch_gmail_dm.py")
        sys.exit(1)

    # Step 1: 拉取各日期的邮件
    html_files = []
    for d in target_dates:
        date_str = d.strftime('%Y-%m-%d')
        html_file = BASE_DIR / f'raw_email_body_{date_str}.html'
        desc = f"Step 1: 拉取 {date_str} 的 DM 雷达日报"
        success = run_step(desc,
            f'{sys.executable} "{fetch_script}" --date {date_str}')
        if success:
            html_files.append(html_file)
        elif html_file.exists():
            print(f"[警告] 拉取失败，使用已有缓存: {html_file}")
            html_files.append(html_file)

    if not html_files:
        print("[错误] 没有可用的邮件HTML文件，无法继续。")
        sys.exit(1)

    # Step 2: 解析 HTML（支持多文件合并）
    parse_script = BASE_DIR / 'parse_email_html.py'
    json_file = BASE_DIR / 'raw_email_body_parsed.json'
    html_args = ' '.join(f'"{f}"' for f in html_files)
    if not run_step(f"Step 2: 解析 HTML 提取结构化数据（{len(html_files)} 个文件）",
                    f'{sys.executable} "{parse_script}" {html_args} -o "{json_file}"'):
        print("[错误] HTML 解析失败")
        sys.exit(1)

    # Step 3: 生成报告
    report_script = BASE_DIR / 'generate_qige_report.py'
    if not run_step("Step 3: 生成舆情日报（MD + DOCX）",
                    f'{sys.executable} "{report_script}" "{json_file}"'):
        print("[错误] 报告生成失败")
        sys.exit(1)

    # Step 4: DM 早报提取与生成
    dm_runner = BASE_DIR / 'dm_daily.py'
    dm_files = []
    if dm_runner.exists():
        dm_cmd = f'{sys.executable} "{dm_runner}" --force'
        # DM only publishes on weekdays. Always use today (or today is
        # a weekend, find the most recent weekday) — not target_dates[0]
        # which might be Saturday when running on Monday.
        dm_target_date = today
        while dm_target_date.weekday() >= 5:
            dm_target_date = dm_target_date - dt.timedelta(days=1)
        dm_cmd += f" --date {dm_target_date.strftime('%Y-%m-%d')}"
        if run_step("Step 4: DM 早报提取与 DOCX 生成", dm_cmd):
            # Find ALL generated DM DOCX files (早报 + 要闻速览)
            dm_files = sorted(BASE_DIR.glob('DM早报_*.docx')) + sorted(BASE_DIR.glob('DM要闻速览_*.docx'))
            if dm_files:
                for f in dm_files:
                    print(f"[信息] DM 早报 DOCX: {f}")
            else:
                print("[警告] DM 早报 DOCX 生成完成但未找到文件")
    else:
        print("[警告] 找不到 dm_daily.py，跳过 DM 早报")

    # Step 5: 发送邮件
    md_file = list(BASE_DIR.glob('【大家资产持仓信用主体舆情日报】_*.md'))
    if md_file:
        md_file = sorted(md_file)[-1]
        send_script = BASE_DIR / 'send_report_email.py'
        if send_script.exists():
            cmd = f'{sys.executable} "{send_script}" "{md_file}"'
            # Attach ALL DM DOCX files (早报 + 要闻速览)
            for f in dm_files:
                cmd += f' --attach "{f}"'
            run_step("Step 5: 发送报告邮件", cmd)
        else:
            print("[警告] 找不到 send_report_email.py，跳过邮件发送")

        # ── 去重标记：报告已生成，写入 sentinel 防重复执行 ──
        sentinel = BASE_DIR / f".daily_sentinel_{today.strftime('%Y%m%d')}"
        sentinel.write_text(f"done at {dt.datetime.now().isoformat()}\nmd: {md_file.name}\n")
        print(f"[去重] sentinel 已写入: {sentinel}")
    else:
        print("[警告] 未找到MD报告文件，跳过邮件发送")

    print()
    print("=" * 60)
    print("  完成！报告已生成（MD + DOCX 双格式）并发送邮件。")
    if dm_files:
        for f in dm_files:
            print(f"  DM 早报附件：{f.name}")
    print("=" * 60)


if __name__ == '__main__':
    main()
