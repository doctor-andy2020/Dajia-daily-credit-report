#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发送舆情日报邮件（DOCX附件 + MD转HTML正文）
用法：python send_report_email.py 【大家资产持仓信用主体舆情日报】_2026-05-27.md
"""

import smtplib
import sys
import os
import datetime
import markdown
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from pathlib import Path

# ============================================================
# 配置
# ============================================================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER = os.environ.get("EMAIL_ACCOUNT", "")
PASSWORD = os.environ.get("EMAIL_PASSWORD", "")  # Gmail 应用专用密码
RECIPIENTS = os.environ.get("EMAIL_RECIPIENTS", "mengsiqi@djbx.com,lizhibo@djbx.com,gr_zc_xypg@djbx.com").split(",")


def md_to_html(md_path):
    """读取MD文件，转换为HTML"""
    with open(md_path, 'r', encoding='utf-8') as f:
        md_content = f.read()

    html_body = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])

    # 包裹在完整的HTML样式中，适配邮件客户端
    styled_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {{ font-family: 'Microsoft YaHei', '微软雅黑', Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.8; }}
  h1 {{ color: #1a1a1a; border-bottom: 2px solid #c00; padding-bottom: 8px; }}
  h2 {{ color: #1a1a1a; margin-top: 24px; }}
  h3, h4 {{ color: #333; }}
  table {{ border-collapse: collapse; width: 100%; margin: 10px 0; }}
  th, td {{ border: 1px solid #ddd; padding: 8px 10px; text-align: left; }}
  th {{ background-color: #f5f5f5; font-weight: bold; }}
  blockquote {{ border-left: 3px solid #c00; padding-left: 12px; color: #666; margin: 10px 0; }}
  hr {{ border: none; border-top: 1px solid #eee; margin: 20px 0; }}
  strong {{ color: #c00; }}
  em {{ font-style: normal; color: #888; }}
</style>
</head>
<body>
{html_body}
</body>
</html>"""

    return styled_html


def attach_file(msg, filepath, display_name=None):
    """Attach a file to the email message."""
    filepath = Path(filepath)
    if not filepath.exists():
        print(f"[警告] 附件不存在: {filepath}")
        return
    with open(filepath, 'rb') as f:
        attachment = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
        attachment.set_payload(f.read())
        encoders.encode_base64(attachment)
        filename = display_name or filepath.name
        attachment.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', filename))
        msg.attach(attachment)
    print(f"[附件] {filepath.name}")


def send_report(md_path, recipients=None, extra_attachments=None):
    """发送报告邮件

    Args:
        md_path: 报告MD文件路径
        recipients: 收件人列表
        extra_attachments: 额外附件文件路径列表 (如 DM 早报 DOCX)
    """
    if recipients is None:
        recipients = RECIPIENTS

    md_path = Path(md_path)
    if not md_path.exists():
        print(f"[错误] 找不到报告文件: {md_path}")
        sys.exit(1)

    # 推断DOCX文件路径
    docx_path = md_path.with_suffix('.docx')
    if not docx_path.exists():
        print(f"[警告] 未找到DOCX文件: {docx_path}，将不附加")

    # 提取报告日期
    report_date = md_path.stem.replace('【大家资产持仓信用主体舆情日报】_', '')

    # 转换为HTML
    html_body = md_to_html(str(md_path))

    # 构建邮件
    msg = MIMEMultipart('mixed')
    msg['From'] = f"大家资产持仓信用主体舆情日报 <{SENDER}>"
    subject = f"【大家资产持仓信用主体舆情日报】{report_date}"
    msg['Subject'] = Header(subject, 'utf-8')
    msg['To'] = ';'.join(recipients)

    # HTML正文
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))

    # DOCX附件（主报告）
    if docx_path.exists():
        attach_file(msg, docx_path, f'【大家资产持仓信用主体舆情日报】{report_date}.docx')

    # 额外附件（DM早报等）
    if extra_attachments:
        for att in extra_attachments:
            if isinstance(att, (str, Path)):
                attach_file(msg, Path(att))

    # 发送 (Gmail 用 STARTTLS)
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.starttls()
            server.login(SENDER, PASSWORD)
            server.sendmail(SENDER, recipients, msg.as_string())
        print(f"[发送成功] 报告已发送至: {', '.join(recipients)}")
    except smtplib.SMTPAuthenticationError:
        print("[发送失败] SMTP认证失败，请检查邮箱账号密码")
        sys.exit(1)
    except Exception as e:
        print(f"[发送失败] {e}")
        sys.exit(1)


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(description="发送舆情日报邮件")
    ap.add_argument("md_file", nargs="?", help="报告MD文件路径")
    ap.add_argument("--attach", action="append", default=[], help="额外附件文件路径（可多次指定）")
    args = ap.parse_args()

    md_file = args.md_file
    if not md_file:
        # 尝试找当天的报告
        today = datetime.date.today().strftime('%Y-%m-%d')
        md_file = f'【大家资产持仓信用主体舆情日报】_{today}.md'
        if not Path(md_file).exists():
            print(f"用法: python {sys.argv[0]} <报告MD文件路径> [--attach 附件]")
            print(f"示例: python {sys.argv[0]} 大家资产持仓信用主体舆情日报_2026-05-29.md --attach DM信用早报_20260529.docx")
            sys.exit(1)

    send_report(md_file, extra_attachments=args.attach)
