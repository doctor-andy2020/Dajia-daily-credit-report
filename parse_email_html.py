#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DM雷达日报 HTML 解析器 v2
从 raw_email_body.html 中提取结构化数据，供报告生成脚本使用

采用 sourceline 定位策略：先找到所有板块标题的 sourceline，
再找到标题之间的所有数据 table，按板块分别解析。
"""

import re
import json
import sys
from pathlib import Path
from bs4 import BeautifulSoup


def clean_text(text):
    if not text:
        return ""
    return re.sub(r'\s+', ' ', text).strip()


# ============================================================
# 板块标题定位
# ============================================================

SECTION_NAMES = ['存续提醒', '舆情', '债圈热议', '成交异动', '一级发行', '评级', '公告', '诉讼']


def find_section_boundaries(soup):
    """找到所有板块标题元素的 sourceline，返回 [(name, line, element), ...]"""
    boundaries = []
    for div in soup.find_all('div'):
        text = clean_text(div.get_text())
        if text in SECTION_NAMES:
            style = div.get('style', '')
            if 'fcb948' in style:
                boundaries.append((text, div.sourceline, div))
    boundaries.sort(key=lambda x: x[1])
    return boundaries


# ============================================================
# 表格分类
# ============================================================

def is_yuqing_style_table(table):
    """判断是否为舆情/公告风格的表格（左侧公司信息 + 右侧内容）"""
    rows = table.find_all('tr', recursive=False)
    if len(rows) < 2:
        return False
    # 第一个 tr 中找 rowspan
    for td in rows[0].find_all('td', recursive=False):
        rowspan = td.get('rowspan', None)
        if rowspan and int(rowspan) >= 2:
            return True
    return False


def is_data_table(table):
    """判断是否为标准数据表格（有 thead/tbody）"""
    return table.find('thead') is not None


# ============================================================
# 单元格解析
# ============================================================

def parse_company_info(td):
    """从左侧单元格提取公司名、评级信息"""
    divs = td.find_all('div', recursive=False)
    result = {'name': '', 'rating': '', 'rating_agency': '', 'rating_level': '', 'rating_date': ''}

    if not divs:
        result['name'] = clean_text(td.get_text())
        return result

    result['name'] = clean_text(divs[0].get_text())
    # 解析后续div中的评级信息
    for div in divs[1:]:
        text = clean_text(div.get_text())
        if not text:
            continue
        # 提取评级机构:等级
        m = re.match(r'(\S+)\s*[：:]\s*(\S+)', text)
        if m:
            result['rating_agency'] = m.group(1)
            result['rating_level'] = m.group(2)
        # 提取日期
        dm = re.search(r'\((\d{4}-\d{2}-\d{2})\)', text)
        if dm:
            result['rating_date'] = dm.group(1)
        # 构建完整的评级字符串
        if result['rating']:
            result['rating'] += ' | ' + text
        else:
            result['rating'] = text

    return result


def parse_content_cell(td, default_label='标签'):
    """从右侧内容单元格提取时间、标签/分类、内容、链接"""
    result = {'time': '', 'tags': '', 'content': '', 'link': '', 'label_type': default_label}

    # 提取纯文本中标记为"标签："或"分类："的部分
    full_text = td.get_text()

    # 查找时间
    time_m = re.search(r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', full_text)
    if time_m:
        result['time'] = time_m.group(1)

    # 查找标签/分类
    label_m = re.search(r'(标签|分类)[：:]\s*', full_text)
    if label_m:
        result['label_type'] = '标签' if label_m.group(1) == '标签' else '分类'
        # 标签值在"标签："之后
        after_label = full_text[label_m.end():]
        # 提取到换行或下一段为止
        label_end = re.search(r'[\n\r]|$', after_label)
        if label_end:
            result['tags'] = clean_text(after_label[:label_end.start()])

    # 查找链接
    link = td.find('a')
    if link:
        result['content'] = clean_text(link.get_text())
        result['link'] = link.get('href', '')

    if not result['content']:
        # 没有链接的话，取所有非标签/时间的文本
        parts = []
        for span in td.find_all('span'):
            t = clean_text(span.get_text())
            if t and '标签：' not in t and '分类：' not in t and not re.match(r'\d{4}-\d{2}-\d{2}', t):
                parts.append(t)
        if parts:
            result['content'] = ' '.join(parts)

    return result


# ============================================================
# 板块解析
# ============================================================

def parse_yuqing_style_tables(tables):
    """解析舆情/公告风格的一组表格"""
    items = []
    for table in tables:
        rows = table.find_all('tr', recursive=False)
        if len(rows) < 2:
            continue

        # 左侧：第一个tr中rowspan的td
        left_td = None
        for td in rows[0].find_all('td', recursive=False):
            if td.get('rowspan'):
                left_td = td
                break
        if not left_td:
            continue

        company_info = parse_company_info(left_td)

        # 右侧：后续tr中的td
        for tr in rows[1:]:
            tds = tr.find_all('td', recursive=False)
            if not tds:
                continue
            content_info = parse_content_cell(tds[0])
            if content_info['content']:
                items.append({**company_info, **content_info})

    return items


def parse_standard_table(table):
    """解析标准表格（有thead/tbody）"""
    headers = []
    rows = []

    thead = table.find('thead')
    if thead:
        for th in thead.find_all('th'):
            headers.append(clean_text(th.get_text()))

    tbody = table.find('tbody')
    if not tbody:
        tbody = table

    for tr in tbody.find_all('tr', recursive=False):
        # 跳过空白占位行 (height:0)
        tr_style = tr.get('style', '')
        if 'height: 0' in tr_style or 'height:0' in tr_style:
            continue
        row = []
        for td in tr.find_all('td', recursive=False):
            row.append(clean_text(td.get_text()))
        # 跳过全空行
        if row and any(cell for cell in row):
            rows.append(row)

    return {'headers': headers, 'rows': rows}


def parse_section_tables(soup, start_line, end_line):
    """解析两个sourceline之间的所有数据表"""
    all_tables = soup.find_all('table')

    yuqing_style = []
    standard_tables = []

    for t in all_tables:
        line = t.sourceline
        if not line:
            continue
        if line < start_line:
            continue
        if end_line and line >= end_line:
            continue

        style = t.get('style', '')

        # 跳过外层布局/装饰table
        if 'width: 870px' in style or 'width:870px' in style:
            continue
        if 'width: 100%' in style or 'width:100%' in style:
            # 只有内部嵌套的数据table才有width:100%（如存续提醒的数据区）
            # 但如果它没有被is_data_table或is_yuqing_style识别，仍跳过
            pass

        # 优先判断类型
        is_yq = is_yuqing_style_table(t)
        is_dt = is_data_table(t)

        if not is_yq and not is_dt:
            continue

        if is_yq:
            yuqing_style.append(t)
        elif is_dt:
            standard_tables.append(t)

    return yuqing_style, standard_tables


# ============================================================
# 提取子板块（估值偏离/前收涨跌）
# ============================================================

def find_sub_sections(soup, start_line, end_line):
    """找到估值偏离、前收涨跌的位置"""
    sub_sections = {}
    for div in soup.find_all('div'):
        text = clean_text(div.get_text())
        if text in ('估值偏离', '前收涨跌'):
            line = div.sourceline
            if line and line > start_line and (not end_line or line < end_line):
                sub_sections[line] = text
    return sub_sections


# ============================================================
# 主解析函数
# ============================================================

def parse_all(html_file):
    with open(html_file, 'r', encoding='utf-8') as f:
        html = f.read()

    soup = BeautifulSoup(html, 'html.parser')

    # 提取报告标题
    title_text = soup.get_text()
    report_name = ''
    report_date = ''
    m1 = re.search(r'DM雷达日报\s*(\S+)\s*重要提醒', title_text)
    m2 = re.search(r'(\d{4}-\d{2}-\d{2})', title_text)
    if m1:
        report_name = m1.group(1)
    if m2:
        report_date = m2.group(1)

    # 找到所有板块边界
    boundaries = find_section_boundaries(soup)
    # 构建板块 → (start_line, end_line) 映射
    section_ranges = {}
    for i, (name, line, _) in enumerate(boundaries):
        end_line = boundaries[i + 1][1] if i + 1 < len(boundaries) else None
        section_ranges[name] = (line, end_line)

    result = {
        'report_name': report_name,
        'report_date': report_date,
        'yuqing': [],
        'zhaiquan_resuo': [],
        'chengjiao': {'headers': [], 'rows': [], 'sub_rows': {}},
        'faxing': {'headers': [], 'rows': []},
        'pingji': {'headers': [], 'rows': []},
        'gonggao': [],
        'susong': [],
    }

    for section_name, (start, end) in section_ranges.items():
        yuqing_style, standard = parse_section_tables(soup, start, end or 999999)

        if section_name in ('舆情',):
            result['yuqing'] = parse_yuqing_style_tables(yuqing_style)

        elif section_name in ('公告',):
            result['gonggao'] = parse_yuqing_style_tables(yuqing_style)
            # 公告板块的label_type应为"分类"
            for item in result['gonggao']:
                item['label_type'] = '分类'

        elif section_name == '成交异动':
            sub_secs = find_sub_sections(soup, start, end or 999999)

            for st in standard:
                st_start = st.sourceline
                # 找到这个table之前最近的子板块标题
                sub_name = None
                for sl in sorted(sub_secs.keys(), reverse=True):
                    if sl < st_start:
                        sub_name = sub_secs[sl]
                        break
                parsed = parse_standard_table(st)
                if sub_name:
                    result['chengjiao']['sub_rows'][sub_name] = parsed['rows']
                else:
                    result['chengjiao']['headers'] = parsed['headers']
                    result['chengjiao']['rows'] = parsed['rows']

        elif section_name == '一级发行':
            for st in standard:
                parsed = parse_standard_table(st)
                if parsed['headers']:
                    result['faxing']['headers'] = parsed['headers']
                result['faxing']['rows'].extend(parsed['rows'])

        elif section_name == '评级':
            for st in standard:
                parsed = parse_standard_table(st)
                if parsed['headers']:
                    result['pingji']['headers'] = parsed['headers']
                result['pingji']['rows'].extend(parsed['rows'])

    return result


def print_summary(data):
    print(f"报告名称: {data['report_name']}")
    print(f"报告日期: {data['report_date']}")
    print(f"舆情条目: {len(data['yuqing'])}")
    print(f"公告条目: {len(data['gonggao'])}")
    print(f"成交异动: {len(data['chengjiao']['rows'])} 行")
    if data['chengjiao']['sub_rows']:
        for k, v in data['chengjiao']['sub_rows'].items():
            print(f"  - {k}: {len(v)} 行")
    print(f"一级发行: {len(data['faxing']['rows'])} 行")
    print(f"评级变动: {len(data['pingji']['rows'])} 行")

    if data['yuqing']:
        print("\n--- 舆情前5条示例 ---")
        for item in data['yuqing'][:5]:
            print(f"  [{item.get('tags', '')}] {item.get('name', '')}: {item.get('content', '')[:80]}")

    if data['gonggao']:
        print("\n--- 公告前5条示例 ---")
        for item in data['gonggao'][:5]:
            print(f"  [{item.get('tags', '')}] {item.get('name', '')}: {item.get('content', '')[:80]}")


def merge_data(data_list):
    """合并多份解析结果，去重（按 主体+内容 去重）"""
    if len(data_list) == 1:
        return data_list[0]

    merged = data_list[0].copy()
    seen = set()

    # 记录已有条目
    for item in merged.get('yuqing', []):
        key = (item.get('name', ''), item.get('content', '')[:100])
        seen.add(('yuqing', key))
    for item in merged.get('gonggao', []):
        key = (item.get('name', ''), item.get('content', '')[:100])
        seen.add(('gonggao', key))

    # 合并后续数据
    for data in data_list[1:]:
        for item in data.get('yuqing', []):
            key = (item.get('name', ''), item.get('content', '')[:100])
            if ('yuqing', key) not in seen:
                merged['yuqing'].append(item)
                seen.add(('yuqing', key))

        for item in data.get('gonggao', []):
            key = (item.get('name', ''), item.get('content', '')[:100])
            if ('gonggao', key) not in seen:
                merged['gonggao'].append(item)
                seen.add(('gonggao', key))

        # 成交异动：合并各子板块
        for sub_name, rows in data.get('chengjiao', {}).get('sub_rows', {}).items():
            if sub_name not in merged['chengjiao']['sub_rows']:
                merged['chengjiao']['sub_rows'][sub_name] = []
            existing_rows = set(
                (r[0], r[1]) for r in merged['chengjiao']['sub_rows'][sub_name] if len(r) >= 2
            )
            for row in rows:
                row_key = (row[0], row[1]) if len(row) >= 2 else tuple(row)
                if row_key not in existing_rows:
                    merged['chengjiao']['sub_rows'][sub_name].append(row)
                    existing_rows.add(row_key)

        # 一级发行
        existing_faxing = set(tuple(r) for r in merged['faxing']['rows'])
        for row in data.get('faxing', {}).get('rows', []):
            if tuple(row) not in existing_faxing:
                merged['faxing']['rows'].append(row)
                existing_faxing.add(tuple(row))

    # 更新报告日期为最新
    dates = [d.get('report_date', '') for d in data_list]
    merged['report_date'] = max(d for d in dates if d)

    return merged


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description='解析DM雷达日报HTML，支持多文件合并')
    ap.add_argument('html_files', nargs='+', help='HTML文件路径（可多个）')
    ap.add_argument('-o', '--output', default='raw_email_body_parsed.json', help='输出JSON路径')
    args = ap.parse_args()

    all_data = []
    for html_file in args.html_files:
        if not Path(html_file).exists():
            print(f"[跳过] 文件不存在: {html_file}")
            continue
        data = parse_all(html_file)
        print(f"\n--- {html_file} ---")
        print_summary(data)
        all_data.append(data)

    if not all_data:
        print("[错误] 没有可解析的HTML文件")
        sys.exit(1)

    merged = merge_data(all_data)
    if len(all_data) > 1:
        print(f"\n=== 合并后 ===")
        print_summary(merged)

    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(merged, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n结构化数据已保存至: {args.output}")
