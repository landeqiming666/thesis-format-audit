from __future__ import annotations

import argparse
import json
import smtplib
import ssl
import time
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

import app


def current_remaining(user: dict) -> int:
    quota = int(user.get("submission_quota") or app.MAX_SUBMISSIONS)
    used = int(user.get("submissions_used") or 0)
    return max(0, quota - used)


def target_users(threshold: int) -> list[dict]:
    users = app.list_users()
    return [
        user
        for user in users
        if user.get("account_status", app.ACCOUNT_STATUS_ACTIVE) == app.ACCOUNT_STATUS_ACTIVE
        and current_remaining(user) < threshold
        and user.get("email")
    ]


def build_email(user: dict, amount: int, before_remaining: int, after_remaining: int, site_url: str) -> EmailMessage:
    message = EmailMessage()
    message["Subject"] = "你的论文格式检测次数已补充"
    message["From"] = f"{app.EMAIL_FROM_NAME} <{app.GMAIL_SMTP_USER}>"
    message["To"] = user["email"]
    message.set_content(
        "\n".join(
            [
                "你好，",
                "",
                f"我们已为你的账号额外补充 {amount} 次论文格式检测机会。",
                f"补充前剩余次数：{before_remaining} 次",
                f"补充后剩余次数：{after_remaining} 次",
                "",
                f"检测入口：{site_url}",
                "",
                "如果你在使用中遇到问题，可以加入官方 QQ 群 537124215 反馈。",
                "",
                "UPC本科论文格式检测工具",
            ]
        )
    )
    return message


def send_email(server: smtplib.SMTP_SSL, user: dict, amount: int, before_remaining: int, after_remaining: int, site_url: str) -> None:
    server.send_message(build_email(user, amount, before_remaining, after_remaining, site_url))


def run_grant(amount: int, threshold: int, execute: bool, backup_dir: Path, site_url: str, delay: float) -> dict:
    backup_dir.mkdir(parents=True, exist_ok=True)
    users = target_users(threshold)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = backup_dir / f"low-quota-grant-{timestamp}.json"

    plan = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "execute": execute,
        "threshold": threshold,
        "grant_amount": amount,
        "site_url": site_url,
        "selected_users": [
            {
                "id": user.get("id"),
                "email": user.get("email"),
                "submissions_used": int(user.get("submissions_used") or 0),
                "submission_quota": int(user.get("submission_quota") or app.MAX_SUBMISSIONS),
                "remaining_before": current_remaining(user),
                "remaining_after": current_remaining(user) + amount,
            }
            for user in users
        ],
    }
    backup_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    summary = {
        "execute": execute,
        "backup_path": str(backup_path),
        "threshold": threshold,
        "grant_amount": amount,
        "selected": len(users),
        "quota_updated": 0,
        "emails_sent": 0,
        "email_failed": [],
        "updated_users": [],
    }

    if not execute or not users:
        return summary

    if not app.GMAIL_SMTP_USER or not app.GMAIL_SMTP_APP_PASSWORD:
        raise RuntimeError("缺少 GMAIL_SMTP_USER 或 GMAIL_SMTP_APP_PASSWORD，无法发送邮件提醒。")

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(app.GMAIL_SMTP_HOST, app.GMAIL_SMTP_PORT, context=context, timeout=30) as server:
        server.login(app.GMAIL_SMTP_USER, app.GMAIL_SMTP_APP_PASSWORD)
        for user in users:
            before_remaining = current_remaining(user)
            before_quota = int(user.get("submission_quota") or app.MAX_SUBMISSIONS)
            after_quota = before_quota + amount
            after_remaining = before_remaining + amount

            (
                app.get_supabase()
                .table(app.SUPABASE_TABLE)
                .update({"submission_quota": after_quota})
                .eq("id", user["id"])
                .execute()
            )
            summary["quota_updated"] += 1
            summary["updated_users"].append(
                {
                    "id": user.get("id"),
                    "email": user.get("email"),
                    "quota_before": before_quota,
                    "quota_after": after_quota,
                    "remaining_before": before_remaining,
                    "remaining_after": after_remaining,
                }
            )

            try:
                send_email(server, user, amount, before_remaining, after_remaining, site_url)
                summary["emails_sent"] += 1
            except Exception as exc:
                summary["email_failed"].append({"email": user.get("email"), "error": str(exc)})

            if delay > 0:
                time.sleep(delay)

    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Grant extra quota to active users whose remaining submissions are below a threshold.")
    parser.add_argument("--execute", action="store_true", help="Actually update quotas and send notification emails.")
    parser.add_argument("--amount", type=int, default=5, help="Quota amount to add to each selected user.")
    parser.add_argument("--threshold", type=int, default=5, help="Select users with remaining submissions below this number.")
    parser.add_argument("--site-url", default="https://upc-thesis-audit.salmonbeach-95e48227.southeastasia.azurecontainerapps.io", help="Site URL included in emails.")
    parser.add_argument("--delay", type=float, default=0.4, help="Delay between emails to reduce SMTP throttling risk.")
    parser.add_argument("--backup-dir", default="maintenance_backups", help="Directory for JSON operation plans.")
    args = parser.parse_args()

    if args.amount <= 0:
        raise SystemExit("--amount must be greater than 0")
    if args.threshold <= 0:
        raise SystemExit("--threshold must be greater than 0")

    summary = run_grant(args.amount, args.threshold, args.execute, Path(args.backup_dir), args.site_url, args.delay)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
