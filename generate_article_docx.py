"""Generate a formatted DOCX from the extracted DM article text."""
import os, re
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

OUT = os.path.dirname(os.path.abspath(__file__))
INPUT = os.path.join(OUT, "dm_article_output.txt")

def read_and_clean(path):
    with open(path, encoding="utf-8") as f:
        lines = f.readlines()

    # Find article start (first line containing "债市要闻速览" or "DM信用早报")
    start = 0
    for i, line in enumerate(lines):
        if re.search(r"(债市要闻速览|DM信用早报|DM早报)\d{4}", line):
            start = i
            break

    # Find article end (line with "涉及区域" followed by numbers only)
    end = len(lines)
    for i in range(len(lines) - 1, start, -1):
        if "涉及区域" in lines[i] or "更多债市要闻详情" in lines[i]:
            end = i
            break

    content = lines[start:end]
    # Strip empty/lines that are just numbers (menu remnants)
    cleaned = []
    for line in content:
        s = line.strip()
        if not s:
            cleaned.append("")
            continue
        # Skip isolated numbers (menu items from left nav)
        if s.isdigit() and len(s) <= 2:
            continue
        # Skip standalone menu items
        if s in ("菜单", "首页", "固收综合屏", "债券违约", "舆情", "城投地图",
                 "利率债市场观点", "债市复盘", "行业分类", "利差曲线", "多资产综合屏",
                 "评级调整", "商票逾期", "DM专栏", "金融板块", "自选雷达", "实时利差",
                 "AI研报", "标记颜色", "刷新", "固定标签页", "关闭其他标签页",
                 "关闭右边标签页", "收藏", "字号", "小", "中", "大", "区域名称",
                 "中华人民共和国"):
            continue
        # Skip user name pattern
        if re.match(r"^孟思锜\s*$", s):
            continue
        # Skip "涉及区域" standalone
        if s == "涉及区域":
            continue
        cleaned.append(line.rstrip())

    # Merge empty line runs
    merged = []
    prev_empty = False
    for line in cleaned:
        empty = not line.strip()
        if empty and prev_empty:
            continue
        merged.append(line)
        prev_empty = empty

    return merged


def parse_sections(lines):
    """Parse article into structured sections."""
    article = {
        "title": "",
        "date": "",
        "tags": "",
        "sections": []  # list of {heading, items: [{subheading, paragraphs}]}
    }

    i = 0
    # Title
    while i < len(lines):
        s = lines[i].strip()
        if s and ("要闻速览" in s or "信用早报" in s or "早报" in s):
            article["title"] = s
            i += 1
            break
        i += 1

    # Date
    if i < len(lines) and re.match(r"\d{4}-\d{2}-\d{2}", lines[i].strip()):
        article["date"] = lines[i].strip()
        i += 1

    # Tags (skip 收藏, 字号 etc.)
    while i < len(lines):
        s = lines[i].strip()
        if s.startswith("标签") or s.startswith("标签"):
            article["tags"] = s
            i += 1
            break
        i += 1

    # Skip to content (skip Excel link, empty lines)
    while i < len(lines):
        s = lines[i].strip()
        if s.startswith("【"):
            break
        i += 1

    # Parse sections
    current_section = None
    current_sub = None
    current_paras = []

    def flush():
        nonlocal current_section, current_sub, current_paras
        if current_paras and current_section is not None:
            text = "\n".join(current_paras).strip()
            if text:
                if current_sub:
                    current_section["items"].append({"subheading": current_sub, "text": text})
                else:
                    current_section["items"].append({"subheading": None, "text": text})
        current_paras = []

    while i < len(lines):
        s = lines[i].strip()

        # Section heading: 【...】
        if s.startswith("【") and s.endswith("】"):
            flush()
            current_section = {"heading": s, "items": []}
            article["sections"].append(current_section)
            current_sub = None
            i += 1
            continue

        # Sub-heading: —xxx— or just 地产公司/金融机构 etc (short standalone lines)
        if s.startswith("—") and s.endswith("—"):
            flush()
            current_sub = s
            i += 1
            continue

        # Content-start markers: these lines begin a new content paragraph
        CONTENT_STARTERS = ["据", "根据", "新华社", "央视新闻", "中新社", "中新网", "人民日报",
                           "当地时间", "此前一天", "同日", "另据"]

        def is_content_start(text, prev_title=""):
            # Pattern 1: starts with content marker
            for m in CONTENT_STARTERS:
                if text.startswith(m):
                    return True
            # Pattern 2: same entity prefix as previous title (continuation)
            if prev_title and len(prev_title) >= 4:
                # Try separator-based prefix first
                sep = next((c for c in ["：", ":", "，", "、", "（"] if c in prev_title[:15]), None)
                if sep:
                    prefix = prev_title.split(sep)[0]
                    if len(prefix) >= 2 and text.startswith(prefix):
                        return True
                else:
                    # No separator: use first ~8 chars as prefix (e.g. 中国央行今日开展, 韩国央行行长称)
                    prefix = prev_title[:8]
                    if len(prefix) >= 4 and text.startswith(prefix):
                        return True
            return False

        # Paragraph content
        if s:
            # Detect title→content boundary where no empty line separates them
            if current_paras and is_content_start(s, current_paras[0]):
                # Previous paragraph(s) = title, this line starts content → split
                flush()
            current_paras.append(s)
        else:
            # Empty line = paragraph boundary
            if current_paras:
                flush()

        i += 1

    flush()
    return article


def build_docx(article):
    doc = Document()

    # --- Styles ---
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(10.5)
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.line_spacing = 1.35
    # Set East Asian font
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.makeelement(qn("w:rFonts"), {})
    rFonts.set(qn("w:eastAsia"), "微软雅黑")
    rPr.insert(0, rFonts)

    # --- Title ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(20)
    title.paragraph_format.space_after = Pt(4)
    run = title.add_run(article["title"])
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.name = "微软雅黑"
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.makeelement(qn("w:rFonts"), {})
    rFonts.set(qn("w:eastAsia"), "微软雅黑")
    rPr.insert(0, rFonts)

    # --- Date & Tags ---
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.paragraph_format.space_after = Pt(2)
    if article["date"]:
        run = meta.add_run(f"发布日期：{article['date']}")
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(128, 128, 128)
        run.font.name = "微软雅黑"
    if article["tags"]:
        meta2 = doc.add_paragraph()
        meta2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        meta2.paragraph_format.space_after = Pt(6)
        run = meta2.add_run(article["tags"])
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(100, 149, 237)
        run.font.name = "微软雅黑"

    # Divider
    div = doc.add_paragraph()
    div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    div.paragraph_format.space_after = Pt(12)
    run = div.add_run("━" * 40)
    run.font.size = Pt(6)
    run.font.color.rgb = RGBColor(200, 200, 200)

    # --- Sections ---
    for sec in article["sections"]:
        # Section heading
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(6)
        run = h.add_run(sec["heading"])
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 60, 120)
        run.font.name = "微软雅黑"
        # Add bottom border to heading
        pPr = h._element.get_or_add_pPr()
        pBdr = pPr.makeelement(qn("w:pBdr"), {})
        bottom = pBdr.makeelement(qn("w:bottom"), {
            qn("w:val"): "single",
            qn("w:sz"): "4",
            qn("w:space"): "1",
            qn("w:color"): "1E3C78",
        })
        pBdr.append(bottom)
        pPr.append(pBdr)

        # --- Special handling for 【宏观要闻】: pair title+content, numbered bold titles ---
        if "宏观要闻" in sec["heading"]:
            items = sec["items"]
            idx = 0
            news_num = 1

            # Content-indicator patterns: next paragraph is content (not a new title) if it starts with:
            CONTENT_MARKERS = [
                "据",      # 据中国人民银行..., 据东方网..., 据彭博社...
                "根据",
                "数据显",  # 数据显示
                "新华社",  # 新华社德黑兰6月12日电...
                "央视新闻", # 央视新闻报道...
                "中新社",
                "中新网",
                "人民日报",
                "当地时间", # 当地时间周四...
                "此前",
                "此前一天",
                "同日",
                "另据",
                "公告显",
                "募集说",
                "上述",
                "该",
                "此次",
                "本次",
                "目前",
                "截至",
                "其中",
                "具体",
                "近日",
            ]

            def is_content(text, title_text=""):
                """Check if text looks like a content/follow-up paragraph (not a new title)."""
                text = text.strip()
                # Pattern 1: starts with source-attribution marker
                for marker in CONTENT_MARKERS:
                    if text.startswith(marker):
                        return True
                # Pattern 2: shares the same entity/subject prefix as title (continuation)
                # e.g. title="中国央行：..." content="中国央行数据显示..."
                if title_text and len(title_text) >= 4:
                    sep = next((c for c in ["：", ":", "，", "、", "（"] if c in title_text[:15]), None)
                    if sep:
                        prefix = title_text.split(sep)[0]
                        if len(prefix) >= 2 and text.startswith(prefix):
                            return True
                    else:
                        # No separator: use first ~8 chars as prefix
                        prefix = title_text[:8]
                        if len(prefix) >= 4 and text.startswith(prefix):
                            return True
                return False

            while idx < len(items):
                title_item = items[idx]
                title_text = title_item["text"].strip()

                # Check if next item is content (continuation) or a new standalone title
                content_item = None
                if idx + 1 < len(items):
                    next_text = items[idx + 1]["text"].strip()
                    if is_content(next_text, title_text):
                        content_item = items[idx + 1]
                        idx += 2
                    else:
                        idx += 1
                else:
                    idx += 1

                # --- Handle merged title+content in standalone items ---
                # If this is a standalone item with an internal newline, it means
                # title and content were merged during parsing (no empty line between them).
                # Split at first newline: bold first line as title, rest as content.
                merged_content = ""
                if not content_item and "\n" in title_text:
                    parts = title_text.split("\n", 1)
                    title_text = parts[0].strip()
                    merged_content = parts[1].strip()

                # --- Numbered bold title ---
                tp = doc.add_paragraph()
                tp.paragraph_format.space_before = Pt(10)
                tp.paragraph_format.space_after = Pt(2)

                # Number prefix
                num_run = tp.add_run(f"{news_num}. ")
                num_run.font.size = Pt(10.5)
                num_run.font.bold = True
                num_run.font.name = "微软雅黑"
                num_run.font.color.rgb = RGBColor(30, 60, 120)

                # Bold title
                title_run = tp.add_run(title_text)
                title_run.font.size = Pt(10.5)
                title_run.font.bold = True
                title_run.font.name = "微软雅黑"

                # --- Content paragraph (from pair or from merged split) ---
                content_text = ""
                if content_item:
                    content_text = content_item["text"].strip()
                elif merged_content:
                    content_text = merged_content

                if content_text:
                    cp = doc.add_paragraph()
                    cp.paragraph_format.space_after = Pt(6)
                    cp.paragraph_format.first_line_indent = Cm(0.6)
                    cr = cp.add_run(content_text)
                    cr.font.size = Pt(10)
                    cr.font.name = "微软雅黑"

                news_num += 1

        else:
            # --- Default rendering for other sections ---
            for item in sec["items"]:
                # Sub-heading
                if item["subheading"]:
                    sh = doc.add_paragraph()
                    sh.paragraph_format.space_before = Pt(8)
                    sh.paragraph_format.space_after = Pt(3)
                    run = sh.add_run(item["subheading"])
                    run.font.size = Pt(11)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(60, 90, 150)
                    run.font.name = "微软雅黑"

                # Split item text into paragraphs by double-newline
                paras = item["text"].split("\n\n")
                for p_text in paras:
                    p_text = p_text.strip()
                    if not p_text:
                        continue
                    p = doc.add_paragraph()
                    p.paragraph_format.space_after = Pt(4)
                    p.paragraph_format.first_line_indent = Cm(0.6)

                    run = p.add_run(p_text)
                    run.font.size = Pt(10)
                    run.font.name = "微软雅黑"

    # --- Footer ---
    doc.add_paragraph()  # spacer
    div2 = doc.add_paragraph()
    div2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = div2.add_run("━" * 40)
    run.font.size = Pt(6)
    run.font.color.rgb = RGBColor(200, 200, 200)

    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(6)
    run = footer.add_run("更多债市要闻详情，请查看DM终端舆情板块")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(128, 128, 128)
    run.font.name = "微软雅黑"
    run.font.italic = True

    # --- Page setup ---
    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)

    return doc


def generate_docx(input_file, output_file=None):
    """Generate formatted DOCX from article text file.

    Args:
        input_file: path to the article text file
        output_file: path for output .docx (auto-generated if None)

    Returns:
        path to generated .docx file, or None on failure
    """
    lines = read_and_clean(input_file)
    article = parse_sections(lines)

    if not article["title"]:
        print("[ERROR] Could not extract article title")
        return None

    print(f"Title: {article['title']}")
    print(f"Date: {article['date']}")
    print(f"Sections: {len(article['sections'])}")
    for sec in article["sections"]:
        print(f"  {sec['heading']}: {len(sec['items'])} items")

    doc = build_docx(article)

    if output_file is None:
        output_file = os.path.join(OUT, "债市要闻速览.docx")

    doc.save(output_file)
    print(f"[OK] Saved: {output_file}")
    return output_file


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Generate formatted DOCX from DM article")
    ap.add_argument("input", nargs="?", help="Input article text file")
    ap.add_argument("--output", "-o", help="Output DOCX file path")
    args = ap.parse_args()

    infile = args.input or INPUT
    generate_docx(infile, args.output)
