#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速诊断：从 Gmail 拉取今天邮件 HTML，测试解析器是否产生完整数据
用法：python diagnose_today.py
前提：需要 EMAIL_ACCOUNT 和 EMAIL_PASSWORD 环境变量
"""

import sys
import os
import json
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).parent

# ── Step 1: 拉取今日邮件 ──
print("=" * 60)
print("Step 1: 拉取今日 DM 雷达日报 HTML")
print("=" * 60)

result = subprocess.run(
    [sys.executable, str(BASE_DIR / "fetch_gmail_dm.py")],
    cwd=str(BASE_DIR),
    capture_output=True, text=True, timeout=60
)
print(result.stdout)
if result.returncode != 0:
    print(f"[STDERR]\n{result.stderr}")
    print("\n拉取失败，尝试复用已有 HTML...")

# ── Step 2: 找到最新的 HTML 文件 ──
html_files = sorted(BASE_DIR.glob("raw_email_body*.html"),
                    key=lambda f: f.stat().st_mtime, reverse=True)
if not html_files:
    print("[FATAL] 没有可用的 HTML 文件")
    sys.exit(1)

latest_html = html_files[0]
print(f"\n使用: {latest_html.name} ({latest_html.stat().st_size / 1024:.0f} KB)")

# ── Step 3: 解析并诊断 ──
print(f"\n{'=' * 60}")
print("Step 2: 解析 HTML")
print("=" * 60)

from parse_email_html import parse_all, print_summary
data = parse_all(str(latest_html))
print_summary(data)

# ── Step 4: 深度完整性检查 ──
print(f"\n{'=' * 60}")
print("Step 3: 完整性诊断")
print("=" * 60)

issues = []

# 检查1: 舆情板块
yuqing = data.get('yuqing', [])
print(f"\n[1] 舆情板块: {len(yuqing)} 条")
if len(yuqing) == 0:
    issues.append("舆情板块为空 — 可能 fcb948 颜色码或表格结构变化")
elif len(yuqing) < 30:
    issues.append(f"舆情仅 {len(yuqing)} 条，异常偏少（正常 ~80 条）")

# 检查2: 公告板块
gonggao = data.get('gonggao', [])
print(f"[2] 公告板块: {len(gonggao)} 条")
if len(gonggao) == 0:
    issues.append("公告板块为空")
elif len(gonggao) < 50:
    issues.append(f"公告仅 {len(gonggao)} 条，异常偏少（正常 ~200 条）")

# 检查3: 成交异动
chengjiao = data.get('chengjiao', {})
pianli = chengjiao.get('sub_rows', {}).get('估值偏离', [])
zhangdie = chengjiao.get('sub_rows', {}).get('前收涨跌', [])
print(f"[3] 成交异动: 估值偏离 {len(pianli)} 条, 前收涨跌 {len(zhangdie)} 条")

# 检查4: 一级发行
faxing = data.get('faxing', {}).get('rows', [])
print(f"[4] 一级发行: {len(faxing)} 条")

# 检查5: 评级
pingji = data.get('pingji', {}).get('rows', [])
print(f"[5] 评级变动: {len(pingji)} 条")

# 检查6: 诉讼
susong = data.get('susong', [])
print(f"[6] 诉讼: {len(susong)} 条")

# 检查7: 条目内容完整性（检查舆情条目是否有空字段）
empty_content = sum(1 for item in yuqing if not item.get('content', '').strip())
empty_name = sum(1 for item in yuqing if not item.get('name', '').strip())
print(f"[7] 数据质量: 空content={empty_content}, 空name={empty_name}")

# ── Step 5: 筛选测试 ──
print(f"\n{'=' * 60}")
print("Step 4: 筛选诊断")
print("=" * 60)

from generate_qige_report import should_exclude_yuqing, classify_yuqing, merge_same_company

kept = [item for item in yuqing if not should_exclude_yuqing(item)]
excluded = len(yuqing) - len(kept)
high = merge_same_company([item for item in kept if classify_yuqing(item) == 'high'])
medium = merge_same_company([item for item in kept if classify_yuqing(item) == 'medium'])

print(f"  筛选结果: {len(yuqing)} 原始 → 排除 {excluded} → 保留 {len(kept)}")
print(f"    [HIGH] {len(high)} 条")
print(f"    [MED]  {len(medium)} 条")

# ── 结论 ──
print(f"\n{'=' * 60}")
print("诊断结论")
print("=" * 60)

if not issues:
    print("\n[OK] 解析器和筛选逻辑工作正常")
    print("\n如果报告内容仍不全，问题可能在:")
    print("  1. daily_runner.py 的日期参数传递")
    print("  2. GitHub Actions 环境变量")
    print("  3. send_report_email.py 的 MD→HTML 转换")
else:
    print("\n[ISSUES FOUND]")
    for i, issue in enumerate(issues, 1):
        print(f"  {i}. {issue}")
