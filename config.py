from __future__ import annotations

import os

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

if load_dotenv:
    load_dotenv()


MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "32"))
MAX_DOCX_ENTRIES = int(os.environ.get("MAX_DOCX_ENTRIES", "1500"))
MAX_DOCX_UNCOMPRESSED_MB = int(os.environ.get("MAX_DOCX_UNCOMPRESSED_MB", "180"))
AUDIT_TIMEOUT_SECONDS = int(os.environ.get("AUDIT_TIMEOUT_SECONDS", "105"))
DOC_CONVERT_TIMEOUT_SECONDS = int(os.environ.get("DOC_CONVERT_TIMEOUT_SECONDS", "60"))
AUTH_TOKEN_MAX_AGE = int(os.environ.get("AUTH_TOKEN_MAX_AGE", str(30 * 24 * 60 * 60)))
MAX_SUBMISSIONS = int(os.environ.get("MAX_SUBMISSIONS", "100"))
MAX_STORED_REPORTS_PER_USER = int(os.environ.get("MAX_STORED_REPORTS_PER_USER", "5"))

SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
SUPABASE_TABLE = "thesis_audit_users"
ADMIN_LOG_TABLE = "thesis_audit_admin_logs"
REPORTS_TABLE = "thesis_audit_reports"
REGISTRATION_CODES_TABLE = "thesis_audit_registration_codes"
EVENTS_TABLE = "thesis_audit_events"
REPORTS_BUCKET = os.environ.get("REPORTS_BUCKET", "thesis-audit-reports")

GITHUB_REPO_URL = os.environ.get("GITHUB_REPO_URL", "https://github.com/landeqiming666/thesis-format-audit").strip()

GMAIL_SMTP_HOST = os.environ.get("GMAIL_SMTP_HOST", "smtp.gmail.com")
GMAIL_SMTP_PORT = int(os.environ.get("GMAIL_SMTP_PORT", "465"))
GMAIL_SMTP_USER = os.environ.get("GMAIL_SMTP_USER", "").strip()
GMAIL_SMTP_APP_PASSWORD = os.environ.get("GMAIL_SMTP_APP_PASSWORD", "").strip()
EMAIL_FROM_NAME = os.environ.get("EMAIL_FROM_NAME", "UPC论文格式检测工具").strip()

GCS_BUCKET = os.environ.get("GCS_BUCKET", "")
GCS_PREFIX = os.environ.get("GCS_PREFIX", "thesis-audit").strip("/")
GCS_PROJECT = os.environ.get("GCS_PROJECT", "")
GCS_CREDENTIALS_JSON = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON", "")

GOOGLE_DRIVE_CREDENTIALS_JSON = os.environ.get("GOOGLE_DRIVE_CREDENTIALS_JSON", "").strip()
GOOGLE_DRIVE_FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "").strip()
GOOGLE_DRIVE_PREFIX = os.environ.get("GOOGLE_DRIVE_PREFIX", "thesis-audit").strip("/")

SUPER_ADMIN_EMAILS = {
    email.strip().lower()
    for email in os.environ.get("SUPER_ADMIN_EMAILS", "2818242447@qq.com").split(",")
    if email.strip()
}
LEGACY_ADMIN_EMAILS = {
    email.strip().lower()
    for email in os.environ.get("ADMIN_EMAILS", "").split(",")
    if email.strip()
}
