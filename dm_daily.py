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
    docx_file = BASE_DIR / f"DM早报_{datestr}.docx"

    # Step A: Extract ALL articles from DM (信用早报 + 要闻速览)
    if not args.skip_extract:
        ok = run_step(
            f"Step A: 从 DM 提取 {date_str} 早报（含信用早报+要闻速览）",
            f'{sys.executable} "{BASE_DIR}/dm_pipeline.py" --date {date_str} --all'
        )
        if not ok:
            # Check if we have any cached txt files
            has_cache = False
            for pattern in [f"dm_zaobao_{datestr}.txt", f"dm_yaowen_{datestr}.txt",
                            f"dm_article_{datestr}.txt", "dm_article_output.txt"]:
                if (BASE_DIR / pattern).exists():
                    has_cache = True
                    break
            if not has_cache:
                print("[错误] DM 早报提取失败，且无缓存文件。")
                sys.exit(1)
            print(f"[警告] 提取部分失败，使用已有缓存文件。")

    # Find all extracted txt files
    zaobao_txt = BASE_DIR / f"dm_zaobao_{datestr}.txt"
    yaowen_txt = BASE_DIR / f"dm_yaowen_{datestr}.txt"
    article_txt = BASE_DIR / f"dm_article_{datestr}.txt"
    default_txt = BASE_DIR / "dm_article_output.txt"

    txt_files = []  # list of (txt_path, docx_path, label)
    if zaobao_txt.exists():
        txt_files.append((zaobao_txt, BASE_DIR / f"DM早报_{datestr}.docx", "DM信用早报"))
    if yaowen_txt.exists():
        txt_files.append((yaowen_txt, BASE_DIR / f"DM要闻速览_{datestr}.docx", "债市要闻速览"))
    # Fallback: old naming convention
    if not txt_files:
        for f in [article_txt, default_txt]:
            if f.exists():
                txt_files.append((f, docx_file, "DM早报"))
                break

    if not txt_files:
        print("[错误] 找不到 DM 文章文本文件。")
        sys.exit(1)

    print(f"[信息] 找到 {len(txt_files)} 篇文章: {', '.join(t[2] for t in txt_files)}")

    # Step B: Generate DOCX for each article
    generated_files = []
    for txt_path, docx_path, label in txt_files:
        ok = run_step(
            f"Step B: 生成 {label} DOCX",
            f'{sys.executable} "{BASE_DIR}/generate_article_docx.py" "{txt_path}" -o "{docx_path}"'
        )
        if ok:
            generated_files.append(docx_path)
        else:
            print(f"[错误] {label} DOCX 生成失败: {docx_path}")

    if not generated_files:
        print("[错误] 所有 DOCX 生成均失败。")
        sys.exit(1)

    # Print generated files
    print(f"\n{'='*60}")
    for docx_path in generated_files:
        print(f"  完成！DM早报 DOCX: {docx_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
