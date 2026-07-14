#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速测试：拖入 DM 邮件 HTML 文件，立即诊断解析结果
用法：
  方法1（推荐）: 在 Gmail 中打开 DM 雷达日报邮件 → 右键 → "另存为" → 保存为 .html
              然后: python quick_test.py 保存的文件.html

  方法2: 如果有环境变量，直接拉取:
               python quick_test.py --fetch
"""

import sys
import os
from pathlib import Path

def test_file(html_path):
    html_path = Path(html_path)
    if not html_path.exists():
        print(f"[ERROR] 文件不存在: {html_path}")
        sys.exit(1)

    print(f"测试文件: {html_path.name} ({html_path.stat().st_size / 1024:.0f} KB)")
    print("=" * 60)

    # 原始HTML结构检查
    with open(html_path, 'r', encoding='utf-8', errors='replace') as f:
        raw = f.read()

    import re
    sections = ['存续提醒', '舆情', '债圈热议', '成交异动', '一级发行', '评级', '公告', '诉讼']
    print("\n[结构检查] 板块标题出现次数:")
    for s in sections:
        count = raw.count(s)
        print(f"  {'OK' if count >= 1 else 'MISS'} 「{s}」: {count} 次")

    fcb = len(re.findall(r'fcb948', raw, re.IGNORECASE))
    print(f"\n  fcb948 颜色码: {fcb} 次 (期望 8 次)")

    # 解析
    from parse_email_html import parse_all
    data = parse_all(str(html_path))

    print(f"\n[解析结果]")
    print(f"  报告名称: {data['report_name']}")
    print(f"  邮件日期: {data['report_date']}")
    print(f"  舆情: {len(data['yuqing'])} 条")
    print(f"  公告: {len(data['gonggao'])} 条")
    print(f"  成交异动-估值偏离: {len(data['chengjiao']['sub_rows'].get('估值偏离', []))} 条")
    print(f"  成交异动-前收涨跌: {len(data['chengjiao']['sub_rows'].get('前收涨跌', []))} 条")
    print(f"  一级发行: {len(data['faxing']['rows'])} 条")
    print(f"  评级变动: {len(data['pingji']['rows'])} 条")
    print(f"  诉讼: {len(data['susong'])} 条")

    # 送 generate_qige_report 的筛选逻辑
    from generate_qige_report import should_exclude_yuqing, classify_yuqing, merge_same_company

    yuqing = data['yuqing']
    kept = [item for item in yuqing if not should_exclude_yuqing(item)]
    excluded = len(yuqing) - len(kept)
    high = [item for item in kept if classify_yuqing(item) == 'high']
    medium = [item for item in kept if classify_yuqing(item) == 'medium']

    gonggao = data['gonggao']
    from generate_qige_report import should_exclude_gonggao
    gonggao_kept = [item for item in gonggao if not should_exclude_gonggao(item)]

    print(f"\n[筛选后]")
    print(f"  舆情: {len(yuqing)} 原始 → 排除{excluded} → 保留{len(kept)}")
    print(f"    高风险: {len(high)} 条")
    print(f"    重要关注: {len(medium)} 条")
    print(f"  公告: {len(gonggao)} 原始 → 保留{len(gonggao_kept)} 条")

    # 判断
    print(f"\n{'=' * 60}")
    print("诊断结论:")
    print("=" * 60)

    if len(yuqing) <= 10:
        print(f"  [FAIL] 舆情仅 {len(yuqing)} 条 → 解析环节丢失了大部分数据!")
        print(f"  → 检查 fcb948 颜色码是否有 8 个")
        print(f"  → 检查 HTML 中 table 结构是否变化")
    elif len(yuqing) >= 50:
        print(f"  [OK] 舆情 {len(yuqing)} 条 → 解析正常")
        print(f"  → 如果报告仍不全，问题在 generate_qige_report.py 筛选逻辑")

    if data['report_date'] != '2026-06-30':
        print(f"  [WARN] 邮件日期是 {data['report_date']}，不是 2026-06-30!")
        print(f"  → 可能拉取了错误的邮件")

    # 打印前10条舆情用于比较
    print(f"\n[舆情前10条]")
    for i, item in enumerate(yuqing[:10]):
        print(f"  {i+1}. [{item.get('tags','')}] {item.get('name','')}: {item.get('content','')[:100]}")
        print(f"     时间: {item.get('time','')}")


if __name__ == '__main__':
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg == '--fetch':
            import subprocess
            print("正在从 Gmail 拉取今日邮件...")
            result = subprocess.run(
                [sys.executable, 'fetch_gmail_dm.py'],
                capture_output=True, text=True, timeout=60
            )
            print(result.stdout)
            if result.returncode != 0:
                print(result.stderr)
                sys.exit(1)
            # Find the HTML file just created
            html_files = sorted(Path('.').glob('raw_email_body*.html'),
                              key=lambda f: f.stat().st_mtime, reverse=True)
            if html_files:
                test_file(html_files[0])
            else:
                print("[ERROR] 未找到拉取的HTML文件")
        else:
            test_file(arg)
    else:
        print("用法:")
        print("  1. python quick_test.py <DM邮件.html>")
        print("     (在Gmail中打开DM雷达日报 → 右键另存为 → 拖入终端)")
        print("  2. python quick_test.py --fetch")
        print("     (需要 EMAIL_ACCOUNT 和 EMAIL_PASSWORD 环境变量)")
