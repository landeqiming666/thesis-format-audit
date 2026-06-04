from __future__ import annotations

import os
import smtplib
import ssl
import sys
from email.message import EmailMessage

from dotenv import load_dotenv


def main() -> int:
    load_dotenv()
    to_email = sys.argv[1].strip() if len(sys.argv) > 1 else os.environ.get("GMAIL_SMTP_USER", "").strip()
    smtp_user = os.environ.get("GMAIL_SMTP_USER", "").strip()
    smtp_password = os.environ.get("GMAIL_SMTP_APP_PASSWORD", "").strip()
    smtp_host = os.environ.get("GMAIL_SMTP_HOST", "smtp.gmail.com").strip()
    smtp_port = int(os.environ.get("GMAIL_SMTP_PORT", "465"))
    from_name = os.environ.get("EMAIL_FROM_NAME", "UPC论文格式检测工具").strip()

    if not smtp_user or not smtp_password:
        print("缺少 GMAIL_SMTP_USER 或 GMAIL_SMTP_APP_PASSWORD，请先填写 .env。")
        return 1
    if not to_email:
        print("缺少收件邮箱。用法：python scripts/maintenance/check_email_smtp.py your-email@example.com")
        return 1

    message = EmailMessage()
    message["Subject"] = "邮箱验证码服务测试"
    message["From"] = f"{from_name} <{smtp_user}>"
    message["To"] = to_email
    message.set_content(
        "这是一封 SMTP 测试邮件。如果你收到了它，说明 Gmail 注册验证码服务已经配置成功。"
    )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=20) as server:
        server.login(smtp_user, smtp_password)
        server.send_message(message)

    print(f"测试邮件已发送到：{to_email}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
