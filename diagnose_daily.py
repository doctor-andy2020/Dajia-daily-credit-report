#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日报诊断脚本：检查本地文件状态、解析器健康度、数据完整性
用法：python diagnose_daily.py
"""

import json
import sys
import os
from pathlib import Path
from datetime import datetime, timedelta, timezone

# Windows 控制台编码修复
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR = Path(__file__).parent
BEIJING_TZ = timezone(timedelta(hours=8))
TODAY = datetime.now(BEIJING_TZ).date()

print("=" * 70)
print("  大家资产持仓信用主体舆情日报 — 诊断报告")
print(f"  运行时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print(f"  北京日期：{TODAY}")
print("=" * 70)

# ── 1. 文件状态检查 ──
print("\n" + "─" * 50)
print("【1】本地文件状态")

# 原始HTML文件
html_files = sorted(BASE_DIR.glob("raw_email_body*.html"))
print(f"\n  原始HTML文件 ({len(html_files)} 个):")
for f in html_files:
    size_kb = f.stat().st_size / 1024
    flag = ""
    if not any(c.isdigit() for c in f.stem.replace('raw_email_body_', '')):
        flag = " ⚠️ 无日期后缀（fallback文件）"
    print(f"    {'✅' if size_kb > 100 else '⚠️ '} {f.name}  ({size_kb:.0f} KB){flag}")

latest_dated = [f for f in html_files if any(c.isdigit() for c in f.stem.replace('raw_email_body_', ''))]
if latest_dated:
    latest = sorted(latest_dated)[-1]
    date_in_name = latest.stem.replace('raw_email_body_', '')
    print(f"\n  📅 最新缓存日期：{date_in_name}")
else:
    print(f"\n  ❌ 没有找到带日期的HTML缓存文件")

# parsed JSON
json_file = BASE_DIR / "raw_email_body_parsed.json"
if json_file.exists():
    with open(json_file, 'r', encoding='utf-8') as f:
        cached_data = json.load(f)
    print(f"\n  📦 raw_email_body_parsed.json:")
    print(f"     报告名称：{cached_data.get('report_name', 'N/A')}")
    print(f"     数据日期：{cached_data.get('report_date', 'N/A')}")
    print(f"     舆情条目：{len(cached_data.get('yuqing', []))} 条")
    print(f"     公告条目：{len(cached_data.get('gonggao', []))} 条")
    days_old = (TODAY - datetime.strptime(cached_data.get('report_date', '2000-01-01'), '%Y-%m-%d').date()).days
    if days_old > 1:
        print(f"     ⚠️ 数据已过期 {days_old} 天！")
else:
    print(f"\n  ❌ raw_email_body_parsed.json 不存在")
    cached_data = None

# Sentinels
sentinels = sorted(BASE_DIR.glob(".daily_sentinel_*"))
daily_sentinels = sorted(BASE_DIR.glob(".dm_sentinel_*"))
print(f"\n  舆情日报 sentinel: {len(sentinels)} 个")
for s in sentinels:
    print(f"    📌 {s.name}")
print(f"  DM早报 sentinel: {len(daily_sentinels)} 个")
for s in daily_sentinels:
    print(f"    📌 {s.name}")

# Reports
md_reports = sorted(BASE_DIR.glob("【大家资产持仓信用主体舆情日报】_*.md"))
docx_reports = sorted(BASE_DIR.glob("【大家资产持仓信用主体舆情日报】_*.docx"))
print(f"\n  已有报告：{len(md_reports)} 个 MD，{len(docx_reports)} 个 DOCX")
if md_reports:
    latest_md = md_reports[-1]
    print(f"    最新 MD：{latest_md.name}")
if docx_reports:
    latest_docx = docx_reports[-1]
    print(f"    最新 DOCX：{latest_docx.name}")

# ── 2. 解析器健康度测试 ──
print("\n" + "─" * 50)
print("【2】HTML解析器健康度测试")

# 用最新的HTML文件测试解析器
test_file = None
if latest_dated:
    test_file = latest
elif html_files:
    test_file = html_files[-1]

if test_file and test_file.exists():
    print(f"\n  测试文件：{test_file.name} ({test_file.stat().st_size / 1024:.0f} KB)")

    # 直接文本检查（不依赖解析器）
    with open(test_file, 'r', encoding='utf-8', errors='replace') as f:
        raw_html = f.read()

    print(f"\n  📝 原始HTML结构检查：")
    sections = ['存续提醒', '舆情', '债圈热议', '成交异动', '一级发行', '评级', '公告', '诉讼']
    for s in sections:
        count = raw_html.count(s)
        icon = "✅" if count >= 1 else "❌"
        print(f"    {icon} 「{s}」：出现 {count} 次")

    import re
    fcb_count = len(re.findall(r'fcb948', raw_html, re.IGNORECASE))
    print(f"\n  🎨 板块标题颜色码 #fcb948：{fcb_count} 次（期望 8 次）")
    if fcb_count < 8:
        print(f"    ⚠️ 可能缺少 {8 - fcb_count} 个板块标题！")

    # 子板块检查
    for sub in ['估值偏离', '前收涨跌']:
        count = raw_html.count(sub)
        icon = "✅" if count >= 1 else "⚠️ "
        print(f"    {icon} 「{sub}」子标题：{count} 次")

    # 实际解析
    print(f"\n  🔧 实际解析测试：")
    try:
        from parse_email_html import parse_all, print_summary
        parsed = parse_all(str(test_file))
        print(f"    舆情：{len(parsed['yuqing'])} 条")
        print(f"    公告：{len(parsed['gonggao'])} 条")
        print(f"    成交异动：估值偏离 {len(parsed['chengjiao']['sub_rows'].get('估值偏离', []))} 条, "
              f"前收涨跌 {len(parsed['chengjiao']['sub_rows'].get('前收涨跌', []))} 条")
        print(f"    一级发行：{len(parsed['faxing']['rows'])} 条")
        print(f"    评级变动：{len(parsed['pingji']['rows'])} 条")
        print(f"    诉讼：{len(parsed['susong'])} 条")

        # 检查是否有空板块
        empty_sections = []
        if len(parsed['yuqing']) == 0: empty_sections.append('舆情')
        if len(parsed['gonggao']) == 0: empty_sections.append('公告')
        if len(parsed['chengjiao']['sub_rows'].get('估值偏离', [])) == 0: empty_sections.append('估值偏离')
        if len(parsed['faxing']['rows']) == 0: empty_sections.append('一级发行')

        if empty_sections:
            print(f"\n    ⚠️ 以下板块为空：{', '.join(empty_sections)}")
        else:
            print(f"\n    ✅ 所有板块解析正常")

    except Exception as e:
        print(f"    ❌ 解析失败：{e}")
        import traceback
        traceback.print_exc()

# ── 3. 报告生成器测试 ──
print("\n" + "─" * 50)
print("【3】筛选规则测试（基于 cached JSON）")

if cached_data and len(cached_data.get('yuqing', [])) > 0:
    from generate_qige_report import should_exclude_yuqing, classify_yuqing

    yuqing = cached_data['yuqing']
    excluded = [item for item in yuqing if should_exclude_yuqing(item)]
    kept = [item for item in yuqing if not should_exclude_yuqing(item)]
    high = [item for item in kept if classify_yuqing(item) == 'high']
    medium = [item for item in kept if classify_yuqing(item) == 'medium']
    low = [item for item in kept if classify_yuqing(item) == 'low']

    print(f"\n  舆情条目分析：")
    print(f"    原始：{len(yuqing)} 条")
    print(f"    排除：{len(excluded)} 条")
    if excluded:
        print(f"      排除原因分布：")
        from collections import Counter
        tag_counts = Counter(item.get('tags', '未知') for item in excluded)
        for tag, count in tag_counts.most_common(5):
            print(f"        - {tag}: {count} 条")
    print(f"    保留：{len(kept)} 条")
    print(f"      🔴 高风险：{len(high)} 条")
    print(f"      🟡 重要关注：{len(medium)} 条")
    print(f"      ⚪ 一般信息：{len(low)} 条（不纳入报告）")

    if len(high) == 0 and len(medium) == 0:
        print(f"\n    ⚠️ 筛选后无风险/关注事项 → 报告「重点舆情」部分会为空")

    # 检查标签分布
    all_tags = []
    for item in yuqing:
        tags = item.get('tags', '')
        all_tags.extend([t.strip() for t in tags.split(',') if t.strip()])
    from collections import Counter
    tag_dist = Counter(all_tags)
    print(f"\n  标签分布 Top 10：")
    for tag, count in tag_dist.most_common(10):
        print(f"    {tag}: {count} 条")

# ── 4. GitHub Actions 状态检查 ──
print("\n" + "─" * 50)
print("【4】GitHub Actions 状态")

import subprocess
# 尝试运行 gh CLI
try:
    result = subprocess.run(
        ['gh', 'run', 'list', '--workflow=daily-report.yml', '--limit=5'],
        capture_output=True, text=True, timeout=15, cwd=str(BASE_DIR)
    )
    if result.returncode == 0 and result.stdout.strip():
        print(f"\n  最近 5 次 daily-report.yml 运行：")
        print(f"  {result.stdout}")
    else:
        print(f"\n  ⚠️ 无法获取运行记录（可能是未登录 gh CLI 或不在 GitHub Actions 环境中）")
        print(f"     提示：运行 gh auth login 先登录")
except FileNotFoundError:
    print(f"\n  ⚠️ gh CLI 未安装")
except Exception as e:
    print(f"\n  ⚠️ 查询失败：{e}")

# ── 5. 结论与建议 ──
print("\n" + "=" * 70)
print("【诊断总结】")
print("=" * 70)

issues = []
suggestions = []

if not latest_dated:
    issues.append("本地无任何带日期的HTML缓存文件")
    suggestions.append("邮件拉取环节失败 → 检查 Gmail IMAP 凭据和网络连接")

if cached_data:
    days_old = (TODAY - datetime.strptime(cached_data.get('report_date', '2000-01-01'), '%Y-%m-%d').date()).days
    if days_old > 1:
        issues.append(f"cached JSON 数据过期 {days_old} 天（日期：{cached_data.get('report_date')}）")
        suggestions.append("GitHub Actions 上的 parsed JSON 可能也是旧的 → 检查 Action 日志中 Step 2 的输出")

if not any(s.name.endswith(f"_{TODAY.strftime('%Y%m%d')}") for s in sentinels):
    issues.append(f"今日（{TODAY}）本地无 sentinel 标记")

today_md = list(BASE_DIR.glob(f"【大家资产持仓信用主体舆情日报】_{TODAY}.md"))
today_docx = list(BASE_DIR.glob(f"【大家资产持仓信用主体舆情日报】_{TODAY}.docx"))
if not today_md:
    issues.append("今日本地无 MD 报告文件")
if not today_docx:
    issues.append("今日本地无 DOCX 报告文件")

if not issues:
    print("\n✅ 未发现明显问题")
else:
    print("\n⚠️ 发现以下问题：")
    for i, issue in enumerate(issues, 1):
        print(f"  {i}. {issue}")

if suggestions:
    print("\n💡 建议操作：")
    for i, sug in enumerate(suggestions, 1):
        print(f"  {i}. {sug}")

# 额外建议
print(f"\n🔍 要进一步定位「内容不全」根因，最有效的方法是：")
print(f"  1. 登录 GitHub → Actions → daily-report.yml → 今天最新的 run")
print(f"  2. 查看 Step 2 (HTML解析) 的日志输出 → 对比历史正常 run 的数据量")
print(f"  3. 下载 Artifact (daily-report) → 对比 MD/DOCX 与收到的邮件")
print(f"  4. 如果 Step 2 输出数据量正常但报告内容少 → 问题在筛选规则")
print(f"  5. 如果 Step 2 输出数据量就少 → 问题在邮件获取或DM邮件模板变化")

print("\n" + "=" * 70)
