#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gmail 邮箱自动拉取脚本 v3
  - 从 DM 雷达日报邮件中提取 Excel 下载链接并尝试下载
  - 同时保存邮件 HTML 正文备用
  - 下载成功 → 保存 .xlsx 文件 →供报告生成脚本直接使用
  - 下载失败 → 保存 raw_email_body.html → 供 HTML 解析后备方案
"""

import imaplib
import email
import datetime
import re
import sys
import os
import urllib.request
from email.header import decode_header


# ============================================================
# 用户配置区
# ============================================================
EMAIL_ACCOUNT = os.environ.get("EMAIL_ACCOUNT", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")  # Gmail 应用专用密码
KEYWORD = os.environ.get("DM_KEYWORD", "DM雷达 2026持仓1")

# ============================================================
# Gmail IMAP 配置
# ============================================================
IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993


def decode_mime_header(header_value):
    """解码 MIME 编码的邮件头"""
    if header_value is None:
        return ""
    decoded_parts = decode_header(header_value)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            try:
                result.append(part.decode(charset if charset else "utf-8", errors="replace"))
            except (LookupError, UnicodeDecodeError):
                result.append(part.decode("utf-8", errors="replace"))
        else:
            result.append(str(part))
    return "".join(result)


def extract_text_body(msg):
    """从邮件中提取纯文本正文"""
    body_parts = []
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in content_disposition:
                continue
            charset = part.get_content_charset() or "utf-8"
            if content_type in ("text/plain", "text/html"):
                try:
                    payload = part.get_payload(decode=True)
                    text = payload.decode(charset, errors="replace")
                    body_parts.append((content_type, text))
                except Exception:
                    pass
    else:
        charset = msg.get_content_charset() or "utf-8"
        try:
            payload = msg.get_payload(decode=True)
            text = payload.decode(charset, errors="replace")
            body_parts.append((msg.get_content_type(), text))
        except Exception:
            pass

    # 优先纯文本
    for ct, content in body_parts:
        if ct == "text/plain":
            return content.strip()
    for ct, content in body_parts:
        if ct == "text/html":
            return strip_html(content).strip()
    return ""


def strip_html(html_text):
    """简易 HTML 标签剥离"""
    html_text = re.sub(r"<(style|script)[^>]*>.*?</\1>", "", html_text, flags=re.DOTALL | re.IGNORECASE)
    html_text = re.sub(r"<br\s*/?>", "\n", html_text, flags=re.IGNORECASE)
    html_text = re.sub(r"</?p[^>]*>", "\n", html_text, flags=re.IGNORECASE)
    html_text = re.sub(r"</?div[^>]*>", "\n", html_text, flags=re.IGNORECASE)
    html_text = re.sub(r"</?(tr|td|th)[^>]*>", " ", html_text, flags=re.IGNORECASE)
    html_text = re.sub(r"<li[^>]*>", "\n- ", html_text, flags=re.IGNORECASE)
    html_text = re.sub(r"<[^>]+>", "", html_text)
    html_text = html_text.replace("&nbsp;", " ").replace("&lt;", "<").replace("&gt;", ">")
    html_text = html_text.replace("&amp;", "&").replace("&quot;", '"')
    html_text = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), html_text)
    html_text = re.sub(r"\n\s*\n", "\n\n", html_text)
    return html_text


def fetch_emails_for_date(target_date):
    """拉取指定日期的 DM 雷达日报邮件，返回 (成功标志, 日期字符串)"""
    date_str = target_date.strftime("%d-%b-%Y")
    date_iso = target_date.strftime("%Y-%m-%d")

    if not EMAIL_ACCOUNT or not EMAIL_PASSWORD:
        print("[错误] 请通过环境变量 EMAIL_ACCOUNT 和 EMAIL_PASSWORD 设置 Gmail 账号和应用专用密码。")
        sys.exit(1)

    # --- 连接 ---
    print(f"[连接] {IMAP_SERVER}:{IMAP_PORT} ...")
    try:
        conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    except Exception as e:
        print(f"[错误] 连接失败: {e}")
        sys.exit(1)

    # --- 登录 ---
    print(f"[登录] {EMAIL_ACCOUNT} ...")
    try:
        conn.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    except imaplib.IMAP4.error as e:
        print(f"[错误] 登录失败: {e}")
        print("  提示：Gmail 需使用「应用专用密码」而非登录密码。")
        print("  开启步骤：Gmail 设置 → 安全性 → 两步验证 → 应用专用密码")
        conn.logout()
        sys.exit(1)

    # --- 选择收件箱 ---
    status, _ = conn.select("INBOX")
    if status != "OK":
        print("[错误] 无法打开收件箱。")
        conn.logout()
        sys.exit(1)

    # --- 搜索邮件（日期范围，容错时区差异）---
    # Gmail IMAP 使用服务器内部日期（受发件/收件时区影响），
    # 用 SINCE...BEFORE 范围搜索比 ON 精确日期更可靠
    yesterday = target_date - datetime.timedelta(days=1)
    tomorrow = target_date + datetime.timedelta(days=1)
    since_str = yesterday.strftime("%d-%b-%Y")
    before_str = tomorrow.strftime("%d-%b-%Y")
    print(f"[搜索] 日期范围 {since_str} → {before_str} ...")
    status, message_ids = conn.search(None, f'(SINCE "{since_str}" BEFORE "{before_str}")')
    if status != "OK":
        print("[错误] 搜索失败。")
        conn.logout()
        sys.exit(1)

    all_ids = message_ids[0].split()
    print(f"[结果] {date_str} 共 {len(all_ids)} 封邮件，正在逐封检查标题 ...")

    # --- 按标题关键词过滤 ---
    matched = []  # (mail_id, subject_str)
    for mail_id in all_ids:
        status, msg_data = conn.fetch(mail_id, "(BODY.PEEK[HEADER.FIELDS (SUBJECT)])")
        if status == "OK":
            raw = b''
            if isinstance(msg_data[0], tuple) and len(msg_data[0]) > 1:
                raw = msg_data[0][1]
            elif isinstance(msg_data[0], bytes):
                raw = msg_data[0]
            fake_msg = email.message_from_bytes(raw)
            subject_decoded = decode_mime_header(fake_msg.get('Subject', ''))
            if KEYWORD in subject_decoded:
                matched.append((mail_id, subject_decoded))

    print(f"[结果] 找到 {len(matched)} 封匹配邮件。")

    if len(matched) == 0:
        print("[写入] 无匹配邮件。")
        conn.logout()
        return

    # --- 处理第一封匹配邮件 ---
    mail_id, subject = matched[0]
    print(f"[处理] {subject}")

    status, msg_data = conn.fetch(mail_id, "(RFC822)")
    if status != "OK":
        print("[错误] 获取邮件失败。")
        conn.logout()
        sys.exit(1)

    msg = email.message_from_bytes(msg_data[0][1])
    sender = decode_mime_header(msg.get("From", ""))
    date_str = msg.get("Date", "")

    # --- 提取 HTML 正文 ---
    html_body = ""
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            charset = part.get_content_charset() or "utf-8"
            try:
                html_body = part.get_payload(decode=True).decode(charset, errors="replace")
            except Exception:
                pass
            break

    # --- 尝试从 HTML 中提取 Excel 下载链接 ---
    xlsx_url = None
    if html_body:
        # 查找 OSS 下载链接
        links = re.findall(r'href="([^"]*oss-cn-shanghai[^"]*)"', html_body)
        if links:
            xlsx_url = links[0]

    # --- 尝试下载 Excel ---
    xlsx_downloaded = False
    if xlsx_url:
        # 清理 URL（去除可能的截断问题）
        xlsx_url = xlsx_url.rstrip('.')
        if not xlsx_url.endswith('.xlsx'):
            xlsx_url += '.xlsx'

        print(f"[下载] 尝试下载 Excel: {xlsx_url}")
        try:
            req = urllib.request.Request(xlsx_url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Referer': 'https://mail.google.com/',
            })
            resp = urllib.request.urlopen(req, timeout=30)
            content = resp.read()
            content_type = resp.headers.get('Content-Type', '')
            if len(content) > 5000 and ('spreadsheet' in content_type.lower() or 'excel' in content_type.lower() or 'xlsx' in content_type.lower() or content[:2] == b'PK'):
                # 用原始 URL 中的文件名
                url_filename = xlsx_url.split('/')[-1]
                if not url_filename.endswith('.xlsx'):
                    url_filename += '.xlsx'
                with open(url_filename, 'wb') as f:
                    f.write(content)
                print(f"[下载] 成功: {url_filename} ({len(content)} bytes)")
                xlsx_downloaded = True
            else:
                print(f"[下载] 文件类型异常: {content_type}, 大小: {len(content)}")
        except Exception as e:
            print(f"[下载] 失败: {e}")

    # --- 保存 HTML 正文（带日期后缀）---
    html_file = f'raw_email_body_{date_iso}.html'
    if html_body:
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_body)
        print(f"[保存] HTML 正文 → {html_file} ({len(html_body)} bytes)")

    # --- 保存纯文本摘要 ---
    text_body = extract_text_body(msg)
    txt_file = f'raw_email_{date_iso}.txt'
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write(f"邮件拉取结果\n")
        f.write(f"拉取时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"发件人：{sender}\n")
        f.write(f"主题：{subject}\n")
        f.write(f"日期：{date_str}\n")
        if xlsx_downloaded:
            f.write(f"Excel 已下载：是\n")
        f.write(f"HTML 正文：{html_file if html_body else '无'}\n")
        f.write("=" * 70 + "\n\n")
        f.write(text_body[:50000])

    conn.logout()

    print()
    if xlsx_downloaded:
        print("[完成] Excel 已下载，可直接用于报告生成。")
    else:
        print("[完成] Excel 未下载成功。请检查链接有效性，或手动下载后放入当前目录。")
    print(f"[完成] HTML 正文已保存至 {html_file}，可作为备选数据源。")

    return html_body != ""


def fetch_today_emails():
    """兼容旧接口：拉取当天邮件"""
    fetch_emails_for_date(datetime.date.today())


if __name__ == "__main__":
    target_date = datetime.date.today()
    if len(sys.argv) > 1 and sys.argv[1] == '--date' and len(sys.argv) > 2:
        target_date = datetime.datetime.strptime(sys.argv[2], '%Y-%m-%d').date()
    fetch_emails_for_date(target_date)
