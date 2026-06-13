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


def get_target_dates():
    """根据今天是星期几，返回需要拉取的邮件日期列表"""
    today = datetime.date.today()
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
    today = dt.date.today()
    target_dates = get_target_dates()

    print("=" * 60)
    print("  大家资产持仓信用主体舆情日报 — 自动化工作流")
    print(f"  启动时间：{dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  今日星期：{['一','二','三','四','五','六','日'][today.weekday()]}")
    print(f"  拉取日期：{', '.join(d.strftime('%Y-%m-%d') for d in target_dates)}")
    print("=" * 60)

    force_run = "--force" in sys.argv or "--force-weekend" in sys.argv

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

    fetch_script = BASE_DIR / 'fetch_126_email.py'
    if not fetch_script.exists():
        print("[错误] 找不到 fetch_126_email.py")
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
    dm_docx = None
    if dm_runner.exists():
        dm_cmd = f'{sys.executable} "{dm_runner}"'
        if force_run:
            dm_cmd += " --force"
        if run_step("Step 4: DM 早报提取与 DOCX 生成", dm_cmd):
            # Find the generated DM DOCX
            dm_files = sorted(BASE_DIR.glob('DM信用早报_*.docx'))
            if dm_files:
                dm_docx = dm_files[-1]
                print(f"[信息] DM 早报 DOCX: {dm_docx}")
    else:
        print("[警告] 找不到 dm_daily.py，跳过 DM 早报")

    # Step 5: 发送邮件
    md_file = list(BASE_DIR.glob('【大家资产持仓信用主体舆情日报】_*.md'))
    if md_file:
        md_file = sorted(md_file)[-1]
        send_script = BASE_DIR / 'send_report_email.py'
        if send_script.exists():
            cmd = f'{sys.executable} "{send_script}" "{md_file}"'
            if dm_docx:
                cmd += f' --attach "{dm_docx}"'
            run_step("Step 5: 发送报告邮件", cmd)
        else:
            print("[警告] 找不到 send_report_email.py，跳过邮件发送")
    else:
        print("[警告] 未找到MD报告文件，跳过邮件发送")

    print()
    print("=" * 60)
    print("  完成！报告已生成（MD + DOCX 双格式）并发送邮件。")
    if dm_docx:
        print(f"  DM 早报附件：{dm_docx.name}")
    print("=" * 60)


if __name__ == '__main__':
    main()
