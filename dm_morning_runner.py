#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DM信用早报 — 独立工作流入口
串联：DM终端提取 → DOCX生成 → 邮件发送
与舆情日报完全分离，独立调度

用法：
    python dm_morning_runner.py              # 正常执行（含时间窗口检查）
    python dm_morning_runner.py --force       # 强制运行（绕过所有检查）
"""
import subprocess
import sys
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from pathlib import Path
import datetime

BASE_DIR = Path(__file__).parent
BEIJING_TZ = datetime.timezone(datetime.timedelta(hours=8))

# ── 邮件配置 ──
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER = os.environ.get("EMAIL_ACCOUNT", "")
PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
RECIPIENTS = os.environ.get("EMAIL_RECIPIENTS", "mengsiqi@djbx.com,lizhibo@djbx.com").split(",")


def beijing_now():
    return datetime.datetime.now(BEIJING_TZ)


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


def send_dm_email(docx_files, report_date, recipients=None):
    """发送DM早报邮件（HTML正文 + DOCX附件）"""
    if recipients is None:
        recipients = RECIPIENTS

    if not SENDER or not PASSWORD:
        print("[错误] 缺少 EMAIL_ACCOUNT 或 EMAIL_PASSWORD 环境变量")
        sys.exit(1)

    msg = MIMEMultipart('mixed')
    msg['From'] = f"DM信用早报 <{SENDER}>"
    weekday_cn = ['一','二','三','四','五','六','日'][datetime.date.today().weekday()]
    msg['Subject'] = Header(f"DM信用早报 {report_date}（周{weekday_cn}）", 'utf-8')
    msg['To'] = ';'.join(recipients)

    # HTML正文
    file_list = ''.join(f'<li>{f.name}</li>' for f in docx_files)
    html_body = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8">
<style>
  body {{ font-family: 'Microsoft YaHei', '微软雅黑', Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.8; }}
  h2 {{ color: #1a1a1a; }}
  .footer {{ color: #999; font-size: 12px; margin-top: 24px; }}
</style>
</head>
<body>
<h2>📊 DM信用早报 — {report_date}（周{weekday_cn}）</h2>
<p>请查收附件中的今日DM早报：</p>
<ul>{file_list}</ul>
<p>共 <strong>{len(docx_files)}</strong> 个文件。</p>
<p class="footer">此邮件由自动化系统发送，如有问题请联系管理员。</p>
</body>
</html>"""
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))

    # 附加DOCX文件
    for f in docx_files:
        if not f.exists():
            print(f"[警告] 附件不存在，跳过: {f}")
            continue
        with open(f, 'rb') as fh:
            attachment = MIMEBase(
                'application',
                'vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            attachment.set_payload(fh.read())
            encoders.encode_base64(attachment)
            attachment.add_header(
                'Content-Disposition', 'attachment',
                filename=('utf-8', '', f.name)
            )
            msg.attach(attachment)
        print(f"[附件] {f.name}")

    # 发送
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.starttls()
            server.login(SENDER, PASSWORD)
            server.sendmail(SENDER, recipients, msg.as_string())
        print(f"[发送成功] DM早报已发送至: {', '.join(recipients)}")
    except smtplib.SMTPAuthenticationError:
        print("[发送失败] SMTP认证失败，请检查邮箱账号密码")
        sys.exit(1)
    except Exception as e:
        print(f"[发送失败] {e}")
        sys.exit(1)


def main():
    import datetime as dt
    now = beijing_now()
    today = now.date()
    weekday = today.weekday()

    force_run = "--force" in sys.argv

    # ── 时间窗口限制：仅在北京时间 07:00-10:00 执行 ──
    hour_bj = now.hour
    if not force_run and not (7 <= hour_bj < 10):
        print(f"[跳过] 当前北京时间 {now.strftime('%H:%M')}，不在执行窗口(07:00-10:00)内。")
        print(f"       使用 --force 可强制运行。")
        sys.exit(0)

    # ── 周末跳过 ──
    if weekday >= 5 and not force_run:
        print(f"[跳过] {today} 是{'周六' if weekday==5 else '周日'}，不执行。使用 --force 强制运行。")
        sys.exit(0)

    # ── 去重检查：今天已成功发送过则跳过 ──
    sentinel = BASE_DIR / f".dm_sentinel_{today.strftime('%Y%m%d')}"
    if sentinel.exists() and not force_run:
        print(f"[跳过] 今日DM早报已发送（sentinel: {sentinel}），使用 --force 强制重跑。")
        sys.exit(0)

    print("=" * 60)
    print("  DM信用早报 — 独立工作流")
    print(f"  启动时间：{now.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  今日星期：{['一','二','三','四','五','六','日'][weekday]}")
    print("=" * 60)

    # Step 1: DM早报提取与DOCX生成
    dm_runner = BASE_DIR / 'dm_daily.py'
    if not dm_runner.exists():
        print("[错误] 找不到 dm_daily.py")
        sys.exit(1)

    dm_cmd = (
        f'{sys.executable} "{dm_runner}" --force'
        f' --date {today.strftime("%Y-%m-%d")}'
    )
    if not run_step("Step 1: DM 早报提取与 DOCX 生成", dm_cmd):
        # 检查是否有可用的缓存文件，有则继续尝试发送
        dm_files = sorted(BASE_DIR.glob('DM早报_*.docx')) + sorted(BASE_DIR.glob('DM要闻速览_*.docx'))
        today_dm = [f for f in dm_files if today.strftime('%Y%m%d') in f.name]
        if not today_dm:
            print("[错误] DM 早报生成失败，且无当日缓存文件。")
            sys.exit(1)
        print(f"[警告] 提取失败，使用已有缓存文件（{len(today_dm)} 个）。")
    else:
        # 查找生成的文件
        dm_files = sorted(BASE_DIR.glob('DM早报_*.docx')) + sorted(BASE_DIR.glob('DM要闻速览_*.docx'))
        today_dm = [f for f in dm_files if today.strftime('%Y%m%d') in f.name]

    if not today_dm:
        print("[错误] 未找到当日DM早报DOCX文件。")
        sys.exit(1)

    for f in today_dm:
        print(f"[信息] DM 早报 DOCX: {f}")

    # Step 2: 发送邮件
    report_date = today.strftime('%Y年%m月%d日')
    send_dm_email(today_dm, report_date)

    # ── 去重标记 ──
    sentinel = BASE_DIR / f".dm_sentinel_{today.strftime('%Y%m%d')}"
    sentinel.write_text(f"done at {now.isoformat()}\nfiles: {', '.join(f.name for f in today_dm)}\n")
    print(f"[去重] sentinel 已写入: {sentinel}")

    print()
    print("=" * 60)
    print("  完成！DM早报已生成并发送。")
    for f in today_dm:
        print(f"  DM 早报附件：{f.name}")
    print("=" * 60)


if __name__ == '__main__':
    main()
