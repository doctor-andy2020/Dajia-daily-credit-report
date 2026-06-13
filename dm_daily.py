#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DM 早报每日自动化流水线
Step A: 从 DM 终端提取当日信用早报/债市要闻速览
Step B: 生成格式化 DOCX

用法：
    python dm_daily.py                        # 今天
    python dm_daily.py --date 2026-06-16       # 指定日期
"""
import subprocess, sys, os, argparse
from datetime import date
from pathlib import Path

BASE_DIR = Path(__file__).parent


def run_step(description, cmd):
    print(f"\n{'='*60}")
    print(f">>> {description}")
    print(f"{'='*60}")
    result = subprocess.run(cmd, shell=True, cwd=str(BASE_DIR), capture_output=False)
    return result.returncode == 0


def main():
    parser = argparse.ArgumentParser(description="DM 早报每日自动化")
    parser.add_argument("--date", help="日期 (YYYY-MM-DD)")
    parser.add_argument("--skip-extract", action="store_true",
                        help="跳过提取，仅用已有 txt 生成 DOCX")
    parser.add_argument("--force", action="store_true",
                        help="强制运行，绕过周末检查")
    args = parser.parse_args()

    if args.date:
        target_date = date.fromisoformat(args.date)
    else:
        target_date = date.today()

    # Skip weekends (DM doesn't publish 早报 on weekends)
    weekday = target_date.weekday()
    if weekday >= 5 and not args.force:
        print(f"[跳过] {target_date} 是{'周六' if weekday==5 else '周日'}，DM 不推送早报（使用 --force 强制运行）。")
        sys.exit(0)
    elif weekday >= 5 and args.force:
        print(f"[强制模式] 周末运行，DM 可能没有当日早报。")

    date_str = target_date.strftime("%Y-%m-%d")
    datestr = target_date.strftime("%Y%m%d")
    WEEKDAYS = ['一','二','三','四','五','六','日']
    print(f"DM 早报自动化 — {date_str}（周{WEEKDAYS[weekday]}）")

    txt_file = BASE_DIR / f"dm_article_{datestr}.txt"
    docx_file = BASE_DIR / f"DM信用早报_{datestr}.docx"

    # Step A: Extract article from DM
    if not args.skip_extract:
        ok = run_step(
            f"Step A: 从 DM 提取 {date_str} 早报",
            f'{sys.executable} "{BASE_DIR}/dm_pipeline.py" --date {date_str}'
        )
        if not ok:
            # Check if we have a cached txt file
            if txt_file.exists():
                print(f"[警告] 提取失败，使用缓存文件: {txt_file}")
            else:
                print("[错误] DM 早报提取失败，且无缓存文件。")
                sys.exit(1)

    # Update txt_file reference (dm_pipeline.py may have saved to default name)
    default_txt = BASE_DIR / "dm_article_output.txt"
    pipeline_txt = BASE_DIR / f"dm_article_{datestr}.txt"
    actual_txt = None
    for f in [pipeline_txt, default_txt, txt_file]:
        if f.exists():
            actual_txt = f
            break

    if not actual_txt:
        print("[错误] 找不到 DM 文章文本文件。")
        sys.exit(1)

    print(f"[信息] 使用文章文件: {actual_txt}")

    # Step B: Generate DOCX
    ok = run_step(
        f"Step B: 生成 DOCX",
        f'{sys.executable} "{BASE_DIR}/generate_article_docx.py" "{actual_txt}" -o "{docx_file}"'
    )
    if not ok:
        print("[错误] DOCX 生成失败。")
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"  完成！DM 早报 DOCX: {docx_file}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
