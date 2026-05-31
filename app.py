from __future__ import annotations

import io
import hashlib
import json
import logging
import multiprocessing
import os
import queue
import random
import re
import shutil
import smtplib
import ssl
import subprocess
import tempfile
import time
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urlencode, urlsplit, urlunsplit, parse_qsl
from uuid import uuid4

from flask import Flask, Response, jsonify, has_request_context, redirect, render_template_string, request, send_file, session, url_for
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from postgrest.exceptions import APIError
from supabase import Client, create_client
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
from zipfile import BadZipFile, ZipFile

from thesis_format_audit import open_docx_document, run_audit

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

try:
    from google.cloud import storage as gcs_storage
except ImportError:
    gcs_storage = None

try:
    from google.oauth2 import service_account
except ImportError:
    service_account = None

try:
    from googleapiclient.discovery import build as google_api_build
    from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
except ImportError:
    google_api_build = None
    MediaIoBaseDownload = None
    MediaIoBaseUpload = None

if load_dotenv:
    load_dotenv()


app = Flask(__name__)
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "32"))
MAX_DOCX_ENTRIES = int(os.environ.get("MAX_DOCX_ENTRIES", "1500"))
MAX_DOCX_UNCOMPRESSED_MB = int(os.environ.get("MAX_DOCX_UNCOMPRESSED_MB", "180"))
AUDIT_TIMEOUT_SECONDS = int(os.environ.get("AUDIT_TIMEOUT_SECONDS", "105"))
DOC_CONVERT_TIMEOUT_SECONDS = int(os.environ.get("DOC_CONVERT_TIMEOUT_SECONDS", "60"))

app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["SESSION_COOKIE_SAMESITE"] = "None"
app.config["SESSION_COOKIE_SECURE"] = True
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")
logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO"))

MAX_SUBMISSIONS = 5
AUTH_TOKEN_MAX_AGE = 7 * 24 * 60 * 60
MAX_TRACKED_USER_AGENT_LENGTH = 320
EMAIL_PATTERN = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$")
EMAIL_CODE_LENGTH = 6
EMAIL_CODE_MAX_AGE = 10 * 60
EMAIL_CODE_RESEND_SECONDS = 60
ACCOUNT_STATUS_ACTIVE = "active"
ACCOUNT_STATUS_FROZEN = "frozen"
ACCOUNT_STATUS_DISABLED = "disabled"
CHINA_TZ = timezone(timedelta(hours=8))
WORD_UPLOAD_EXTENSIONS = {".doc", ".docx"}
DOCX_MIMETYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DOC_MIMETYPE = "application/msword"
SOFFICE_BINARY = os.environ.get("SOFFICE_BINARY", "").strip() or shutil.which("soffice") or shutil.which("libreoffice")
UNKNOWN_COLLEGE = "未识别"
COLLEGE_ALIASES = {
    "信息科学与工程学院": ("信息科学与工程学院", "计算机科学与技术学院", "软件学院", "人工智能学院", "网信学院"),
    "石油工程学院": ("石油工程学院",),
    "化学化工学院": ("化学化工学院", "化工学院"),
    "机电工程学院": ("机电工程学院",),
    "储运与建筑工程学院": ("储运与建筑工程学院", "储运学院", "建筑工程学院"),
    "地球科学与技术学院": ("地球科学与技术学院", "地学院"),
    "地球物理学院": ("地球物理学院",),
    "新能源学院": ("新能源学院",),
    "材料科学与工程学院": ("材料科学与工程学院", "材料学院"),
    "海洋与空间信息学院": ("海洋与空间信息学院",),
    "控制科学与工程学院": ("控制科学与工程学院", "控制学院"),
    "经济管理学院": ("经济管理学院", "经管学院"),
    "理学院": ("理学院",),
    "外国语学院": ("外国语学院",),
    "文法学院": ("文法学院",),
    "马克思主义学院": ("马克思主义学院",),
    "体育教学部": ("体育教学部", "体育学院"),
}
COLLEGE_LABEL_PATTERN = re.compile(r"(?:所在)?(?:学院|院系|院（系）|院\(系\)|培养单位|教学单位|系别)\s*[:：]?\s*(.{0,40})")
COLLEGE_NAME_PATTERN = re.compile(r"([\u4e00-\u9fa5A-Za-z0-9（）()·]{2,30}(?:学院|教学部))")
RATE_LIMITS = {
    "login": (10, 5 * 60),
    "register": (5, 60 * 60),
    "email_code": (8, 60 * 60),
    "audit": (8, 60 * 60),
    "admin": (30, 5 * 60),
}
RATE_BUCKETS: defaultdict[tuple[str, str], list[float]] = defaultdict(list)
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
SUPABASE_TABLE = "thesis_audit_users"
ADMIN_LOG_TABLE = "thesis_audit_admin_logs"
REPORTS_TABLE = "thesis_audit_reports"
REGISTRATION_CODES_TABLE = "thesis_audit_registration_codes"
EVENTS_TABLE = "thesis_audit_events"
REPORTS_BUCKET = os.environ.get("REPORTS_BUCKET", "thesis-audit-reports")
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
USER_COLUMNS = (
    "id,email,password_hash,submissions_used,submission_quota,account_status,is_admin,"
    "invite_code,invited_by,register_ip,register_user_agent,last_login_at,last_login_ip,"
    "last_login_user_agent,last_audit_at,last_audit_ip,last_audit_user_agent,created_at"
)
ADMIN_USER_COLUMNS = (
    "id,email,submissions_used,submission_quota,account_status,is_admin,invite_code,invited_by,"
    "register_ip,register_user_agent,last_login_at,last_login_ip,last_login_user_agent,"
    "last_audit_at,last_audit_ip,last_audit_user_agent,created_at"
)
REPORT_COLUMNS = (
    "id,user_id,user_email,original_filename,report_filename,report_storage_path,status,"
    "error_message,college_name,college_source,college_raw_text,client_ip,user_agent,original_storage_backend,original_storage_path,"
    "original_gcs_path,original_drive_file_id,original_drive_path,original_size_bytes,original_sha256,report_storage_backend,report_gcs_path,"
    "report_size_bytes,report_sha256,created_at"
)
REGISTRATION_CODE_COLUMNS = "id,code,note,max_uses,used_count,is_active,created_by,created_at"
EVENT_COLUMNS = "id,event_type,user_id,user_email,path,client_ip,user_agent,metadata,created_at"
_GCS_CLIENT = None
_DRIVE_SERVICE = None
ADMIN_SORT_OPTIONS = {
    "created_desc",
    "created_asc",
    "remaining_desc",
    "remaining_asc",
    "quota_desc",
    "quota_asc",
    "used_desc",
    "used_asc",
    "email_asc",
    "email_desc",
}
ADMIN_PER_PAGE_OPTIONS = {10, 20, 50, 100}
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


def uploaded_word_names(filename: str) -> tuple[str, str]:
    raw_name = (filename or "").strip()
    safe_name = secure_filename(raw_name)
    if not safe_name:
        safe_name = "thesis.docx"
    elif Path(safe_name).suffix.lower() not in WORD_UPLOAD_EXTENSIONS:
        safe_name = f"{safe_name}.docx"
    display_name = raw_name.replace("\\", "/").rsplit("/", 1)[-1]
    display_suffix = Path(display_name).suffix.lower()
    if display_suffix in WORD_UPLOAD_EXTENSIONS:
        display_stem = display_name[: -len(display_suffix)].strip()
    else:
        display_stem = Path(display_name).stem.strip() if display_name else ""
    display_stem = display_stem or Path(safe_name).stem or "thesis"
    return safe_name, f"{display_stem}_format_audit_report.html"


def wants_fetch_response() -> bool:
    return request.headers.get("X-Requested-With") == "fetch" or "application/json" in request.headers.get("Accept", "")


def audit_reject(message: str, status: int):
    if wants_fetch_response():
        return Response(message, status=status, mimetype="text/plain; charset=utf-8")
    return render_home(error=message), status


def validate_docx_package(path: Path) -> None:
    if path.stat().st_size <= 0:
        raise ValueError("上传的文件是空文件，请重新选择 .docx。")

    try:
        with ZipFile(path) as package:
            infos = package.infolist()
            names = {info.filename for info in infos}
            if "[Content_Types].xml" not in names or "word/document.xml" not in names:
                raise ValueError("这个 .docx 缺少 Word 正文结构，请在 Word/WPS 中另存为 .docx 后再上传。")
            if len(infos) > MAX_DOCX_ENTRIES:
                raise ValueError("这个 .docx 内部文件过多，暂时无法安全检测。请压缩图片或另存为 .docx 后再试。")
            uncompressed = sum(info.file_size for info in infos)
            if uncompressed > MAX_DOCX_UNCOMPRESSED_MB * 1024 * 1024:
                raise ValueError(f"这个 .docx 解压后超过 {MAX_DOCX_UNCOMPRESSED_MB}MB，建议先压缩图片后再上传。")
    except BadZipFile as exc:
        raise ValueError("这个文件后缀是 .docx，但内部不是有效的 Word 文档包，可能是损坏文件、网页改后缀或旧版 .doc 直接改名。请用 Word/WPS 打开原文，选择“另存为”真正的 .docx 后再上传。") from exc


def sniff_word_upload_kind(path: Path, filename: str = "") -> str:
    suffix = Path(filename or "").suffix.lower()
    header = path.read_bytes()[:8]
    if header.startswith(b"PK"):
        return "docx"
    if header.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "doc"
    if suffix == ".docx":
        return "docx"
    if suffix == ".doc":
        return "doc"
    return ""


def convert_doc_to_docx(source_path: Path, work_dir: Path) -> Path:
    if not SOFFICE_BINARY:
        raise ValueError("服务器暂时没有安装 Word 转换组件，无法检测 .doc 文件。请先另存为 .docx 后上传。")
    converted_dir = work_dir / "converted"
    converted_dir.mkdir(parents=True, exist_ok=True)
    conversion_source = source_path
    if conversion_source.suffix.lower() != ".doc":
        conversion_source = work_dir / f"{source_path.stem or uuid4().hex}.doc"
        shutil.copyfile(source_path, conversion_source)
    command = [
        SOFFICE_BINARY,
        "--headless",
        "--convert-to",
        "docx",
        "--outdir",
        str(converted_dir),
        str(conversion_source),
    ]
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=DOC_CONVERT_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired as exc:
        raise ValueError("这个 .doc 文件转换超时，请在 Word/WPS 中另存为 .docx 后再上传。") from exc
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        app.logger.warning("DOC conversion failed: %s", detail)
        raise ValueError("这个 .doc 文件无法自动转换，请在 Word/WPS 中另存为 .docx 后再上传。")
    candidates = sorted(converted_dir.glob("*.docx"), key=lambda item: item.stat().st_mtime, reverse=True)
    if not candidates:
        raise ValueError("这个 .doc 文件转换后没有生成有效 .docx，请在 Word/WPS 中另存为 .docx 后再上传。")
    validate_docx_package(candidates[0])
    return candidates[0]


def prepare_docx_for_audit(upload_path: Path, original_filename: str, work_dir: Path) -> Path:
    kind = sniff_word_upload_kind(upload_path, original_filename)
    if kind == "docx":
        validate_docx_package(upload_path)
        return upload_path
    if kind == "doc":
        return convert_doc_to_docx(upload_path, work_dir)
    raise ValueError("当前只支持 .doc 或 .docx 文件。请上传 Word 文档，或在 Word/WPS 中另存为 .docx 后再试。")


def compact_docx_text(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def readable_docx_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def normalize_college_candidate(value: str) -> str:
    text = readable_docx_text(value)
    text = re.sub(r"^(?:所在)?(?:学院|院系|院（系）|院\(系\)|培养单位|教学单位|系别)\s*[:：]?\s*", "", text)
    text = text.strip("：:;；,，。[]【】()（） \t\r\n")
    for delimiter in ("专业", "班级", "学生", "姓名", "学号", "题目", "论文", "指导", "日期", "年级", "届"):
        if delimiter in text:
            text = text.split(delimiter, 1)[0]
    return compact_docx_text(text).strip("：:;；,，。[]【】()（）")


def identify_college_from_text(value: str) -> str:
    text = compact_docx_text(value)
    if not text:
        return ""
    for canonical, aliases in COLLEGE_ALIASES.items():
        if any(compact_docx_text(alias) in text for alias in aliases):
            return canonical
    label_match = COLLEGE_LABEL_PATTERN.search(value)
    if label_match:
        normalized = normalize_college_candidate(label_match.group(1))
        if normalized:
            alias_match = identify_college_from_text(normalized)
            if alias_match:
                return alias_match
            name_match = COLLEGE_NAME_PATTERN.search(normalized)
            if name_match:
                return normalize_college_candidate(name_match.group(1))
    for name in COLLEGE_NAME_PATTERN.findall(value):
        normalized = normalize_college_candidate(name)
        if normalized and "大学" not in normalized and len(normalized) <= 24:
            return normalized
    return ""


def first_docx_cover_texts(path: Path, paragraph_limit: int = 80, table_limit: int = 8) -> list[dict]:
    texts: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="docx-college-") as tmp:
        doc, _compatible_path = open_docx_document(path, Path(tmp))
        for paragraph in doc.paragraphs[:paragraph_limit]:
            text = readable_docx_text(paragraph.text)
            if text:
                texts.append({"text": text, "source": "封面段落"})
        for table_index, table in enumerate(doc.tables[:table_limit], start=1):
            for row in table.rows[:24]:
                cells = [readable_docx_text(cell.text) for cell in row.cells]
                cells = [cell for cell in cells if cell]
                if not cells:
                    continue
                for cell_index, cell in enumerate(cells):
                    texts.append({"text": cell, "source": f"封面表格{table_index}"})
                    if re.fullmatch(r"(?:所在)?(?:学院|院系|院（系）|院\(系\)|培养单位|教学单位|系别)[:：]?", compact_docx_text(cell)):
                        for next_cell in cells[cell_index + 1:]:
                            if next_cell:
                                texts.append({"text": f"学院：{next_cell}", "source": f"封面表格{table_index}"})
                                break
                if len(cells) > 1:
                    texts.append({"text": " ".join(cells), "source": f"封面表格{table_index}"})
    return texts


def extract_college_from_docx(path: Path) -> dict:
    for item in first_docx_cover_texts(path):
        college = identify_college_from_text(item["text"])
        if college:
            return {
                "college_name": college,
                "college_source": item["source"],
                "college_raw_text": item["text"][:240],
            }
    return {"college_name": UNKNOWN_COLLEGE, "college_source": "", "college_raw_text": ""}


def safe_extract_college_from_docx(path: Path) -> dict:
    try:
        return extract_college_from_docx(path)
    except Exception:
        app.logger.warning("Failed to extract college from uploaded docx", exc_info=True)
        return {"college_name": UNKNOWN_COLLEGE, "college_source": "", "college_raw_text": ""}


def original_content_type(filename: str) -> str:
    return DOC_MIMETYPE if Path(filename or "").suffix.lower() == ".doc" else DOCX_MIMETYPE


def archive_original_upload(user: dict, upload_path: Path, original_filename: str) -> dict:
    original_bytes = upload_path.read_bytes()
    storage_path = original_storage_path_for(user, original_filename)
    archive_path = storage_path
    storage_backend = ""
    gcs_path = ""
    drive_file_id = ""
    drive_path = ""
    try:
        storage_backend = upload_original_to_storage(storage_path, original_bytes, original_content_type(original_filename))
    except Exception:
        storage_path = ""
        app.logger.warning("Failed to upload original Word file to Supabase for user %s", user["id"], exc_info=True)
    if gcs_is_configured():
        try:
            gcs_path = upload_original_to_gcs(user, archive_path, original_bytes, original_content_type(original_filename))
        except Exception:
            app.logger.warning("Failed to upload original Word file to GCS for user %s", user["id"], exc_info=True)
    if drive_is_configured():
        try:
            drive_upload = upload_original_to_drive(user, archive_path, original_bytes, original_content_type(original_filename))
            drive_file_id = drive_upload["file_id"]
            drive_path = drive_upload["path"]
        except Exception:
            app.logger.warning("Failed to upload original Word file to Google Drive for user %s", user["id"], exc_info=True)
    return {
        "original_storage_backend": storage_backend,
        "original_storage_path": storage_path,
        "original_gcs_path": gcs_path,
        "original_drive_file_id": drive_file_id,
        "original_drive_path": drive_path,
        "original_size_bytes": len(original_bytes),
        "original_sha256": sha256_hex(original_bytes),
    }


def audit_worker(docx_path: str, report_path: str, result_queue) -> None:
    try:
        run_audit(Path(docx_path), Path(report_path))
        result_queue.put(("ok", ""))
    except Exception as exc:
        result_queue.put((exc.__class__.__name__, str(exc)))


def run_audit_with_timeout(docx_path: Path, report_path: Path) -> None:
    start_method = "fork" if "fork" in multiprocessing.get_all_start_methods() else "spawn"
    ctx = multiprocessing.get_context(start_method)
    result_queue = ctx.Queue(maxsize=1)
    process = ctx.Process(target=audit_worker, args=(str(docx_path), str(report_path), result_queue))
    process.start()
    process.join(AUDIT_TIMEOUT_SECONDS)

    if process.is_alive():
        process.terminate()
        process.join(5)
        if process.is_alive():
            process.kill()
            process.join(5)
        raise TimeoutError("检测时间过长，已自动停止。请先压缩图片、删除无关附件，或把论文拆小后再试。")

    try:
        status, message = result_queue.get_nowait()
    except queue.Empty as exc:
        raise RuntimeError(f"检测进程异常退出，退出码 {process.exitcode}。") from exc

    if status == "ok":
        return
    if status == "ValueError":
        raise ValueError(message)
    raise RuntimeError(message or "未知检测错误")


def audit_error_response(exc: Exception) -> Response:
    if isinstance(exc, (ValueError, TimeoutError)):
        message = f"检测失败：{exc}"
    else:
        error_id = uuid4().hex[:8]
        app.logger.exception("Audit failed [%s]", error_id)
        message = f"检测失败：服务器处理这个文件时遇到异常（错误编号 {error_id}）。请稍后重试，或在 Word/WPS 中另存为 .docx 后再上传。"
    return Response(message, status=500, mimetype="text/plain; charset=utf-8")


def auth_serializer() -> URLSafeTimedSerializer:
    return URLSafeTimedSerializer(app.secret_key, salt="thesis-audit-auth")


def generate_auth_token(user_id: str) -> str:
    return auth_serializer().dumps({"user_id": user_id})


def user_id_from_token(token: str) -> str | None:
    if not token:
        return None
    try:
        data = auth_serializer().loads(token, max_age=AUTH_TOKEN_MAX_AGE)
    except (BadSignature, SignatureExpired):
        return None
    user_id = data.get("user_id")
    return str(user_id) if user_id else None


def client_ip() -> str:
    forwarded_for = request.headers.get("X-Forwarded-For", "")
    if forwarded_for:
        return forwarded_for.split(",", 1)[0].strip()
    return request.remote_addr or "unknown"


def email_verification_enabled() -> bool:
    return bool(GMAIL_SMTP_USER and GMAIL_SMTP_APP_PASSWORD)


def generate_email_code() -> str:
    return "".join(str(random.randint(0, 9)) for _ in range(EMAIL_CODE_LENGTH))


def store_email_code(email: str, code: str) -> None:
    session["email_code_target"] = email.lower().strip()
    session["email_code_value"] = code
    session["email_code_sent_at"] = int(time.time())


def email_code_remaining_seconds() -> int:
    sent_at = int(session.get("email_code_sent_at", 0) or 0)
    if sent_at <= 0:
        return 0
    elapsed = int(time.time()) - sent_at
    return max(0, EMAIL_CODE_RESEND_SECONDS - elapsed)


def clear_email_code() -> None:
    session.pop("email_code_target", None)
    session.pop("email_code_value", None)
    session.pop("email_code_sent_at", None)


def is_valid_email_code(email: str, code: str) -> bool:
    target = session.get("email_code_target", "")
    value = session.get("email_code_value", "")
    sent_at = int(session.get("email_code_sent_at", 0) or 0)
    if not target or not value or not sent_at:
        return False
    if int(time.time()) - sent_at > EMAIL_CODE_MAX_AGE:
        return False
    return target == email.lower().strip() and value == (code or "").strip()


def send_registration_email_code(email: str, code: str) -> None:
    if not email_verification_enabled():
        raise RuntimeError("邮箱验证码服务尚未配置。")
    message = EmailMessage()
    message["Subject"] = "你的注册验证码"
    message["From"] = f"{EMAIL_FROM_NAME} <{GMAIL_SMTP_USER}>"
    message["To"] = email
    message.set_content(
        "\n".join(
            [
                f"你好，",
                "",
                f"你的注册验证码是：{code}",
                f"验证码 {EMAIL_CODE_MAX_AGE // 60} 分钟内有效，请勿泄露给他人。",
                "",
                "如果这不是你的操作，请忽略这封邮件。",
            ]
        )
    )
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(GMAIL_SMTP_HOST, GMAIL_SMTP_PORT, context=context, timeout=20) as server:
        server.login(GMAIL_SMTP_USER, GMAIL_SMTP_APP_PASSWORD)
        server.send_message(message)


def request_user_agent() -> str:
    if not has_request_context():
        return ""
    return (request.headers.get("User-Agent", "") or "")[:MAX_TRACKED_USER_AGENT_LENGTH]


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def is_valid_email(email: str) -> bool:
    return bool(EMAIL_PATTERN.fullmatch((email or "").strip().lower()))


def is_valid_registration_email(email: str) -> bool:
    normalized = (email or "").strip().lower()
    return is_valid_email(normalized) and normalized.endswith("@qq.com")


def request_trace_payload(prefix: str = "") -> dict:
    return {
        f"{prefix}ip": client_ip(),
        f"{prefix}user_agent": request_user_agent(),
    }


def rate_limit(scope: str) -> Response | None:
    max_requests, window_seconds = RATE_LIMITS[scope]
    now = time.time()
    key = (scope, client_ip())
    recent = [timestamp for timestamp in RATE_BUCKETS[key] if now - timestamp < window_seconds]
    RATE_BUCKETS[key] = recent
    if len(recent) >= max_requests:
        return Response("请求太频繁，请稍后再试。", status=429, mimetype="text/plain; charset=utf-8")
    recent.append(now)
    return None


@app.after_request
def add_security_headers(response: Response) -> Response:
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    return response


@app.errorhandler(RequestEntityTooLarge)
def handle_large_upload(_exc):
    return Response(f"文件太大了，当前最多支持 {MAX_UPLOAD_MB}MB。请先压缩图片或另存为较小的 .docx。", status=413, mimetype="text/plain; charset=utf-8")


def get_supabase() -> Client:
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        raise RuntimeError("Supabase is not configured. Set SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY.")
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)


def gcs_is_configured() -> bool:
    return bool(GCS_BUCKET and gcs_storage is not None)


def drive_is_configured() -> bool:
    return bool(
        GOOGLE_DRIVE_CREDENTIALS_JSON
        and GOOGLE_DRIVE_FOLDER_ID
        and service_account is not None
        and google_api_build is not None
        and MediaIoBaseDownload is not None
        and MediaIoBaseUpload is not None
    )


def get_gcs_client():
    global _GCS_CLIENT
    if not gcs_is_configured():
        raise RuntimeError("Google Cloud Storage is not configured.")
    if _GCS_CLIENT is not None:
        return _GCS_CLIENT
    if GCS_CREDENTIALS_JSON:
        if service_account is None:
            raise RuntimeError("google-auth service account support is not available.")
        info = json.loads(GCS_CREDENTIALS_JSON)
        credentials = service_account.Credentials.from_service_account_info(info)
        _GCS_CLIENT = gcs_storage.Client(project=GCS_PROJECT or info.get("project_id"), credentials=credentials)
    else:
        _GCS_CLIENT = gcs_storage.Client(project=GCS_PROJECT or None)
    return _GCS_CLIENT


def get_drive_service():
    global _DRIVE_SERVICE
    if not drive_is_configured():
        raise RuntimeError("Google Drive archive is not configured.")
    if _DRIVE_SERVICE is not None:
        return _DRIVE_SERVICE
    info = json.loads(GOOGLE_DRIVE_CREDENTIALS_JSON)
    credentials = service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive.file"],
    )
    _DRIVE_SERVICE = google_api_build("drive", "v3", credentials=credentials, cache_discovery=False)
    return _DRIVE_SERVICE


def maybe_single_data(result) -> dict | None:
    return result.data if result is not None else None


def find_user_by_id(user_id: str) -> dict | None:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select(USER_COLUMNS)
        .eq("id", user_id)
        .maybe_single()
        .execute()
    )
    return maybe_single_data(result)


def find_user_by_email(email: str) -> dict | None:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select(USER_COLUMNS)
        .eq("email", email)
        .maybe_single()
        .execute()
    )
    return maybe_single_data(result)


def normalize_invite_code(code: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", code or "").upper()


def normalize_registration_code(code: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", code or "").upper()


def generate_invite_code() -> str:
    return uuid4().hex[:10].upper()


def generate_registration_code() -> str:
    return f"UPC{uuid4().hex[:9]}".upper()


def find_user_by_invite_code(code: str) -> dict | None:
    invite_code = normalize_invite_code(code)
    if not invite_code:
        return None
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select(USER_COLUMNS)
        .eq("invite_code", invite_code)
        .maybe_single()
        .execute()
    )
    return maybe_single_data(result)


def create_unique_invite_code() -> str:
    for _ in range(8):
        code = generate_invite_code()
        if find_user_by_invite_code(code) is None:
            return code
    return uuid4().hex.upper()


def ensure_user_invite_code(user: dict | None) -> dict | None:
    if not user or user.get("invite_code"):
        return user
    for _ in range(8):
        code = create_unique_invite_code()
        try:
            (
                get_supabase()
                .table(SUPABASE_TABLE)
                .update({"invite_code": code})
                .eq("id", user["id"])
                .execute()
            )
        except APIError:
            continue
        updated = dict(user)
        updated["invite_code"] = code
        return updated
    return user


def find_registration_code(code: str) -> dict | None:
    normalized = normalize_registration_code(code)
    if not normalized:
        return None
    result = (
        get_supabase()
        .table(REGISTRATION_CODES_TABLE)
        .select(REGISTRATION_CODE_COLUMNS)
        .eq("code", normalized)
        .maybe_single()
        .execute()
    )
    return maybe_single_data(result)


def list_registration_codes(limit: int = 12) -> list[dict]:
    try:
        result = (
            get_supabase()
            .table(REGISTRATION_CODES_TABLE)
            .select(REGISTRATION_CODE_COLUMNS)
            .order("created_at", desc=True)
            .limit(limit)
            .execute()
        )
        return result.data or []
    except Exception:
        app.logger.warning("Failed to list registration codes", exc_info=True)
        return []


def registration_code_remaining(item: dict | None) -> int:
    if not item:
        return 0
    return max(0, int(item.get("max_uses") or 0) - int(item.get("used_count") or 0))


def registration_code_is_available(item: dict | None) -> bool:
    return bool(item and item.get("is_active") and registration_code_remaining(item) > 0)


def create_registration_code(actor: dict, max_uses: int, note: str = "") -> dict:
    max_uses = min(max(max_uses, 1), 999)
    note = (note or "").strip()[:120]
    for _ in range(10):
        code = generate_registration_code()
        if find_registration_code(code) is not None:
            continue
        result = (
            get_supabase()
            .table(REGISTRATION_CODES_TABLE)
            .insert(
                {
                    "code": code,
                    "note": note,
                    "max_uses": max_uses,
                    "used_count": 0,
                    "is_active": True,
                    "created_by": actor.get("email", ""),
                }
            )
            .execute()
        )
        return result.data[0]
    raise RuntimeError("注册码生成失败，请重试。")


def consume_registration_code(code: str) -> dict:
    normalized = normalize_registration_code(code)
    if not normalized:
        raise ValueError("QQ群注册码不存在、已停用或使用次数已满。请加入官方 QQ 群 537124215，从群公告获取最新注册码。")
    result = get_supabase().rpc(
        "consume_thesis_audit_registration_code",
        {"target_code": normalized},
    ).execute()
    data = result.data or []
    if not data:
        raise ValueError("QQ群注册码不存在、已停用或使用次数已满。请加入官方 QQ 群 537124215，从群公告获取最新注册码。")
    return data[0]


def update_registration_code_status(code_id: str, is_active: bool) -> dict:
    result = (
        get_supabase()
        .table(REGISTRATION_CODES_TABLE)
        .update({"is_active": bool(is_active)})
        .eq("id", code_id)
        .execute()
    )
    if not result.data:
        raise ValueError("注册码不存在。")
    return result.data[0]


def create_user(email: str, password: str, invited_by: str | None = None) -> dict:
    trace = request_trace_payload()
    payload = {
        "email": email,
        "password_hash": generate_password_hash(password, method="pbkdf2:sha256"),
        "submission_quota": MAX_SUBMISSIONS,
        "account_status": ACCOUNT_STATUS_ACTIVE,
        "is_admin": email in SUPER_ADMIN_EMAILS,
        "invite_code": create_unique_invite_code(),
        "register_ip": trace["ip"],
        "register_user_agent": trace["user_agent"],
        "last_login_at": utc_now_iso(),
        "last_login_ip": trace["ip"],
        "last_login_user_agent": trace["user_agent"],
    }
    if invited_by:
        payload["invited_by"] = invited_by
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .insert(payload)
        .execute()
    )
    return result.data[0]


def update_user_login_trace(user_id: str) -> None:
    (
        get_supabase()
        .table(SUPABASE_TABLE)
        .update(
            {
                "last_login_at": utc_now_iso(),
                "last_login_ip": client_ip(),
                "last_login_user_agent": request_user_agent(),
            }
        )
        .eq("id", user_id)
        .execute()
    )


def update_user_audit_trace(user_id: str) -> None:
    (
        get_supabase()
        .table(SUPABASE_TABLE)
        .update(
            {
                "last_audit_at": utc_now_iso(),
                "last_audit_ip": client_ip(),
                "last_audit_user_agent": request_user_agent(),
            }
        )
        .eq("id", user_id)
        .execute()
    )


def increment_submissions(user_id: str) -> None:
    user = find_user_by_id(user_id)
    max_allowed = int(user.get("submission_quota", MAX_SUBMISSIONS)) if user else MAX_SUBMISSIONS
    result = get_supabase().rpc(
        "increment_thesis_audit_submissions",
        {"target_user_id": user_id, "max_allowed": max_allowed},
    ).execute()
    if result.data is not True:
        raise RuntimeError("Submission limit reached.")


def current_user() -> dict | None:
    user_id = session.get("user_id")
    if not user_id:
        user_id = user_id_from_token(request.values.get("auth_token", ""))
        if user_id:
            session["user_id"] = user_id
    if not user_id:
        return None
    return ensure_user_invite_code(find_user_by_id(user_id))


def remaining_submissions(user: dict | None) -> int:
    if user is None:
        return 0
    quota = int(user.get("submission_quota", MAX_SUBMISSIONS))
    return max(0, quota - int(user["submissions_used"]))


def user_quota(user: dict | None) -> int:
    if user is None:
        return MAX_SUBMISSIONS
    return int(user.get("submission_quota", MAX_SUBMISSIONS))


def is_super_admin(user: dict | None) -> bool:
    return bool(user and user.get("email", "").lower() in SUPER_ADMIN_EMAILS)


def is_admin(user: dict | None) -> bool:
    if not user:
        return False
    email = user.get("email", "").lower()
    return is_super_admin(user) or bool(user.get("is_admin")) or email in LEGACY_ADMIN_EMAILS


def list_users() -> list[dict]:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select(ADMIN_USER_COLUMNS)
        .order("created_at", desc=True)
        .execute()
    )
    return result.data or []


def list_admin_logs() -> list[dict]:
    result = (
        get_supabase()
        .table(ADMIN_LOG_TABLE)
        .select("id,actor_user_id,actor_email,action,target_user_id,target_email,summary,details,created_at")
        .order("created_at", desc=True)
        .execute()
    )
    return result.data or []


def list_reports_for_user(user_id: str) -> list[dict]:
    result = (
        get_supabase()
        .table(REPORTS_TABLE)
        .select(REPORT_COLUMNS)
        .eq("user_id", user_id)
        .order("created_at", desc=True)
        .execute()
    )
    return result.data or []


def list_reports_for_admin() -> list[dict]:
    result = (
        get_supabase()
        .table(REPORTS_TABLE)
        .select(REPORT_COLUMNS)
        .order("created_at", desc=True)
        .execute()
    )
    return result.data or []


def find_report_by_id(report_id: str) -> dict | None:
    result = (
        get_supabase()
        .table(REPORTS_TABLE)
        .select(REPORT_COLUMNS)
        .eq("id", report_id)
        .maybe_single()
        .execute()
    )
    return maybe_single_data(result)


def create_report_record(
    *,
    user: dict,
    original_filename: str,
    report_filename: str,
    status: str,
    college_name: str = UNKNOWN_COLLEGE,
    college_source: str = "",
    college_raw_text: str = "",
    report_storage_path: str = "",
    original_storage_backend: str = "",
    original_storage_path: str = "",
    original_gcs_path: str = "",
    original_drive_file_id: str = "",
    original_drive_path: str = "",
    original_size_bytes: int = 0,
    original_sha256: str = "",
    report_storage_backend: str = "",
    report_gcs_path: str = "",
    report_size_bytes: int = 0,
    report_sha256: str = "",
    error_message: str = "",
) -> dict:
    result = (
        get_supabase()
        .table(REPORTS_TABLE)
        .insert(
            {
                "user_id": user["id"],
                "user_email": user.get("email", ""),
                "original_filename": original_filename,
                "report_filename": report_filename,
                "report_storage_path": report_storage_path,
                "status": status,
                "error_message": error_message,
                "college_name": college_name or UNKNOWN_COLLEGE,
                "college_source": college_source,
                "college_raw_text": college_raw_text[:240],
                "client_ip": client_ip(),
                "user_agent": request_user_agent(),
                "original_storage_backend": original_storage_backend,
                "original_storage_path": original_storage_path,
                "original_gcs_path": original_gcs_path,
                "original_drive_file_id": original_drive_file_id,
                "original_drive_path": original_drive_path,
                "original_size_bytes": original_size_bytes,
                "original_sha256": original_sha256,
                "report_storage_backend": report_storage_backend,
                "report_gcs_path": report_gcs_path,
                "report_size_bytes": report_size_bytes,
                "report_sha256": report_sha256,
            }
        )
        .execute()
    )
    return result.data[0]


def report_storage_path_for(user: dict, report_filename: str) -> str:
    safe_email = secure_filename(user.get("email", "")) or user["id"]
    return f"{user['id']}/{uuid4().hex}_{safe_email}_{secure_filename(report_filename)}"


def original_storage_path_for(user: dict, original_filename: str) -> str:
    safe_email = secure_filename(user.get("email", "")) or user["id"]
    safe_filename = secure_filename(original_filename) or "thesis.docx"
    return f"{user['id']}/{uuid4().hex}_{safe_email}_{safe_filename}"


def gcs_object_path(kind: str, user: dict, storage_path: str) -> str:
    parts = [part for part in [GCS_PREFIX, kind, storage_path] if part]
    return "/".join(parts)


def drive_object_path(kind: str, user: dict, storage_path: str) -> str:
    parts = [part for part in [GOOGLE_DRIVE_PREFIX, kind, storage_path] if part]
    return "/".join(parts)


def drive_file_name_for_path(path: str) -> str:
    return re.sub(r"[\\/]+", "__", path).strip("_") or f"thesis-audit-{uuid4().hex}"


def sha256_hex(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def upload_to_supabase_storage(storage_path: str, content: bytes, content_type: str) -> None:
    get_supabase().storage.from_(REPORTS_BUCKET).upload(
        path=storage_path,
        file=content,
        file_options={
            "content-type": content_type,
            "upsert": "true",
        },
    )


def upload_to_gcs_storage(object_path: str, content: bytes, content_type: str) -> None:
    bucket = get_gcs_client().bucket(GCS_BUCKET)
    blob = bucket.blob(object_path)
    blob.upload_from_string(content, content_type=content_type)


def upload_to_drive_storage(drive_path: str, content: bytes, content_type: str) -> str:
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=content_type, resumable=False)
    metadata = {
        "name": drive_file_name_for_path(drive_path),
        "parents": [GOOGLE_DRIVE_FOLDER_ID],
        "description": drive_path,
    }
    created = (
        get_drive_service()
        .files()
        .create(
            body=metadata,
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        )
        .execute()
    )
    return created["id"]


def upload_report_to_storage(storage_path: str, report_bytes: bytes) -> str:
    upload_to_supabase_storage(storage_path, report_bytes, "text/html; charset=utf-8")
    return "supabase"


def upload_original_to_storage(storage_path: str, original_bytes: bytes, content_type: str = DOCX_MIMETYPE) -> str:
    upload_to_supabase_storage(
        storage_path,
        original_bytes,
        content_type,
    )
    return "supabase"


def upload_report_to_gcs(user: dict, storage_path: str, report_bytes: bytes) -> str:
    object_path = gcs_object_path("reports", user, storage_path)
    upload_to_gcs_storage(object_path, report_bytes, "text/html; charset=utf-8")
    return object_path


def upload_original_to_gcs(user: dict, storage_path: str, original_bytes: bytes, content_type: str = DOCX_MIMETYPE) -> str:
    object_path = gcs_object_path("originals", user, storage_path)
    upload_to_gcs_storage(
        object_path,
        original_bytes,
        content_type,
    )
    return object_path


def upload_original_to_drive(user: dict, storage_path: str, original_bytes: bytes, content_type: str = DOCX_MIMETYPE) -> dict:
    drive_path = drive_object_path("originals", user, storage_path)
    file_id = upload_to_drive_storage(drive_path, original_bytes, content_type)
    return {"file_id": file_id, "path": drive_path}


def download_from_supabase_storage(storage_path: str) -> bytes:
    return get_supabase().storage.from_(REPORTS_BUCKET).download(storage_path)


def download_from_gcs_storage(object_path: str) -> bytes:
    return get_gcs_client().bucket(GCS_BUCKET).blob(object_path).download_as_bytes()


def download_from_drive_storage(file_id: str) -> bytes:
    request_media = get_drive_service().files().get_media(fileId=file_id, supportsAllDrives=True)
    output = io.BytesIO()
    downloader = MediaIoBaseDownload(output, request_media)
    done = False
    while not done:
        _status, done = downloader.next_chunk()
    return output.getvalue()


def download_report_from_storage(report: dict) -> bytes:
    if report.get("report_storage_path"):
        return download_from_supabase_storage(report["report_storage_path"])
    if report.get("report_gcs_path"):
        return download_from_gcs_storage(report["report_gcs_path"])
    raise FileNotFoundError("Report storage path is empty.")


def download_original_from_storage(report: dict) -> bytes:
    if report.get("original_storage_path"):
        return download_from_supabase_storage(report["original_storage_path"])
    if report.get("original_gcs_path"):
        return download_from_gcs_storage(report["original_gcs_path"])
    if report.get("original_drive_file_id"):
        return download_from_drive_storage(report["original_drive_file_id"])
    raise FileNotFoundError("Original storage path is empty.")


def report_status_label(status: str) -> str:
    return {
        "success": "已完成",
        "audit_failed": "检测失败",
        "storage_failed": "存档失败",
    }.get(status or "success", "未知")


def format_datetime_display(value: str) -> str:
    raw = (value or "").strip()
    if not raw:
        return ""
    normalized = raw.replace("Z", "+00:00")
    if re.search(r"[+-]\d{2}$", normalized):
        normalized = f"{normalized}:00"
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return raw
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    local_time = parsed.astimezone(CHINA_TZ)
    return local_time.strftime("%Y年%m月%d日 %H:%M:%S")


def add_user_quota(user_id: str, amount: int) -> None:
    user = find_user_by_id(user_id)
    if user is None:
        raise ValueError("用户不存在。")
    new_quota = max(user_quota(user), int(user["submissions_used"])) + amount
    (
        get_supabase()
        .table(SUPABASE_TABLE)
        .update({"submission_quota": new_quota})
        .eq("id", user_id)
        .execute()
    )


def award_invite_bonus(inviter_id: str) -> None:
    add_user_quota(inviter_id, 1)


def reduce_user_quota(user_id: str, amount: int) -> None:
    user = find_user_by_id(user_id)
    if user is None:
        raise ValueError("用户不存在。")
    used = int(user["submissions_used"])
    quota = user_quota(user)
    remaining = max(quota - used, 0)
    if amount > remaining:
        raise ValueError(f"减少次数不能超过当前剩余次数（剩余 {remaining} 次）。")
    new_quota = quota - amount
    (
        get_supabase()
        .table(SUPABASE_TABLE)
        .update({"submission_quota": new_quota})
        .eq("id", user_id)
        .execute()
    )


def admin_user_quota_payload(user: dict) -> dict:
    return {
        "id": user["id"],
        "submissions_used": int(user["submissions_used"]),
        "submission_quota": user_quota(user),
        "remaining": remaining_submissions(user),
    }


def enrich_admin_user(user: dict) -> dict:
    item = dict(user)
    item["is_super_admin"] = item.get("email", "").lower() in SUPER_ADMIN_EMAILS
    item["is_admin"] = bool(item.get("is_admin")) or item["is_super_admin"] or item.get("email", "").lower() in LEGACY_ADMIN_EMAILS
    item["remaining"] = remaining_submissions(item)
    item["invite_count"] = 0
    item["invited_by_email"] = ""
    item["invite_link"] = ""
    item["created_at_display"] = format_datetime_display(item.get("created_at", ""))
    item["last_login_at_display"] = format_datetime_display(item.get("last_login_at", ""))
    item["last_audit_at_display"] = format_datetime_display(item.get("last_audit_at", ""))
    return item


def attach_invite_stats(users: list[dict]) -> list[dict]:
    email_by_id = {item.get("id"): item.get("email", "") for item in users if item.get("id")}
    invite_counts: defaultdict[str, int] = defaultdict(int)
    for item in users:
        inviter_id = item.get("invited_by")
        if inviter_id:
            invite_counts[str(inviter_id)] += 1
    for item in users:
        user_id = str(item.get("id", ""))
        invite_code = normalize_invite_code(item.get("invite_code", ""))
        item["invite_count"] = invite_counts.get(user_id, 0)
        item["invited_by_email"] = email_by_id.get(item.get("invited_by"), "") if item.get("invited_by") else ""
        item["invite_code"] = invite_code
        item["invite_link"] = url_for("index", invite=invite_code, _external=True) if invite_code else ""
    return users


def parse_positive_int(value: str | None, default: int) -> int:
    try:
        parsed = int(value or "")
    except (TypeError, ValueError):
        return default
    return parsed if parsed > 0 else default


def admin_per_page(value: str | None) -> int:
    parsed = parse_positive_int(value, 20)
    return parsed if parsed in ADMIN_PER_PAGE_OPTIONS else 20


def admin_sort_value(value: str | None) -> str:
    candidate = (value or "created_desc").strip().lower()
    return candidate if candidate in ADMIN_SORT_OPTIONS else "created_desc"


def admin_redirect_url(token: str, fallback_endpoint: str = "admin") -> str:
    candidate = (request.form.get("next") or request.values.get("next") or "").strip()
    if candidate.startswith("/admin"):
        return candidate
    return url_for(fallback_endpoint, auth_token=token)


def redirect_with_message(destination: str, message: str):
    parts = urlsplit(destination)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["message"] = message
    return redirect(urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query, doseq=True), parts.fragment)))


def build_reports_url(auth_token: str, state: dict, endpoint: str, **overrides) -> str:
    params = {
        "auth_token": auth_token,
        "q": state["q"],
        "status": state["status"],
        "college": state.get("college", "all"),
        "page": state["page"],
    }
    params.update(overrides)
    return url_for(endpoint, **params)


def admin_table_state(auth_token: str) -> dict:
    return {
        "auth_token": auth_token,
        "q": request.args.get("q", "").strip(),
        "status": (request.args.get("status", "all") or "all").strip().lower(),
        "quota": (request.args.get("quota", "all") or "all").strip().lower(),
        "sort": admin_sort_value(request.args.get("sort")),
        "per_page": admin_per_page(request.args.get("per_page")),
        "page": parse_positive_int(request.args.get("page"), 1),
    }


def build_admin_url(auth_token: str, state: dict, **overrides) -> str:
    params = {
        "auth_token": auth_token,
        "q": state["q"],
        "status": state["status"],
        "quota": state["quota"],
        "sort": state["sort"],
        "per_page": state["per_page"],
        "page": state["page"],
    }
    params.update(overrides)
    return url_for("admin", **params)


def apply_admin_user_filters(users: list[dict], state: dict) -> list[dict]:
    keyword = state["q"].lower()
    status = state["status"]
    quota = state["quota"]
    filtered: list[dict] = []
    for item in users:
        haystack = " ".join(
            [
                item.get("email", ""),
                item.get("id", ""),
                item.get("created_at", ""),
                item.get("register_ip", ""),
                item.get("last_login_ip", ""),
                item.get("last_audit_ip", ""),
                item.get("register_user_agent", ""),
                item.get("last_login_user_agent", ""),
                item.get("last_audit_user_agent", ""),
            ]
        ).lower()
        if keyword and keyword not in haystack:
            continue
        if status != "all" and item.get("account_status", ACCOUNT_STATUS_ACTIVE) != status:
            continue
        remaining = int(item.get("remaining", 0))
        if quota == "remaining" and remaining <= 0:
            continue
        if quota == "empty" and remaining > 0:
            continue
        filtered.append(item)
    return filtered


def sort_admin_users(users: list[dict], sort_key: str) -> list[dict]:
    if sort_key == "created_asc":
        return sorted(users, key=lambda item: item.get("created_at", ""))
    if sort_key == "remaining_desc":
        return sorted(users, key=lambda item: (-int(item.get("remaining", 0)), item.get("email", "").lower()))
    if sort_key == "remaining_asc":
        return sorted(users, key=lambda item: (int(item.get("remaining", 0)), item.get("email", "").lower()))
    if sort_key == "quota_desc":
        return sorted(users, key=lambda item: (-int(item.get("submission_quota", 0)), item.get("email", "").lower()))
    if sort_key == "quota_asc":
        return sorted(users, key=lambda item: (int(item.get("submission_quota", 0)), item.get("email", "").lower()))
    if sort_key == "used_desc":
        return sorted(users, key=lambda item: (-int(item.get("submissions_used", 0)), item.get("email", "").lower()))
    if sort_key == "used_asc":
        return sorted(users, key=lambda item: (int(item.get("submissions_used", 0)), item.get("email", "").lower()))
    if sort_key == "email_asc":
        return sorted(users, key=lambda item: item.get("email", "").lower())
    if sort_key == "email_desc":
        return sorted(users, key=lambda item: item.get("email", "").lower(), reverse=True)
    return sorted(users, key=lambda item: item.get("created_at", ""), reverse=True)


def paginate_items(items: list[dict], page: int, per_page: int) -> tuple[list[dict], dict]:
    total = len(items)
    total_pages = max(1, (total + per_page - 1) // per_page) if total else 1
    current_page = min(max(page, 1), total_pages)
    start_index = (current_page - 1) * per_page
    end_index = start_index + per_page
    page_items = items[start_index:end_index]
    return page_items, {
        "total": total,
        "page": current_page,
        "per_page": per_page,
        "pages": total_pages,
        "start": start_index + 1 if total else 0,
        "end": min(end_index, total),
    }


def build_page_numbers(current_page: int, total_pages: int, radius: int = 2) -> list[int]:
    start = max(1, current_page - radius)
    end = min(total_pages, current_page + radius)
    return list(range(start, end + 1))


def summarize_admin_stats(users: list[dict]) -> dict:
    return {
        "total": len(users),
        "active": sum(1 for item in users if item.get("account_status", ACCOUNT_STATUS_ACTIVE) == ACCOUNT_STATUS_ACTIVE),
        "frozen": sum(1 for item in users if item.get("account_status") == ACCOUNT_STATUS_FROZEN),
        "disabled": sum(1 for item in users if item.get("account_status") == ACCOUNT_STATUS_DISABLED),
        "admins": sum(1 for item in users if item.get("is_admin")),
        "invited": sum(1 for item in users if item.get("invited_by")),
        "invite_rewards": sum(int(item.get("invite_count", 0)) for item in users),
    }


def parse_datetime_value(value: str | None) -> datetime | None:
    if not value:
        return None
    normalized = value.replace("Z", "+00:00")
    if re.search(r"[+-]\d{2}$", normalized):
        normalized = f"{normalized}:00"
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(CHINA_TZ)


def day_key(value: str | None) -> str:
    parsed = parse_datetime_value(value)
    return parsed.strftime("%m-%d") if parsed else "未知"


def list_events_for_admin(limit: int = 3000) -> list[dict]:
    try:
        result = (
            get_supabase()
            .table(EVENTS_TABLE)
            .select(EVENT_COLUMNS)
            .order("created_at", desc=True)
            .limit(limit)
            .execute()
        )
        return result.data or []
    except Exception:
        app.logger.warning("Failed to load analytics events", exc_info=True)
        return []


def record_event(event_type: str, user: dict | None = None, metadata: dict | None = None) -> None:
    try:
        payload = {
            "event_type": event_type[:60],
            "user_id": user.get("id") if user else None,
            "user_email": user.get("email", "") if user else "",
            "path": request.path[:200] if has_request_context() else "",
            "client_ip": client_ip(),
            "user_agent": request_user_agent(),
            "metadata": metadata or {},
        }
        get_supabase().table(EVENTS_TABLE).insert(payload).execute()
    except Exception:
        app.logger.debug("Failed to record analytics event %s", event_type, exc_info=True)


def summarize_traffic_stats(users: list[dict], reports: list[dict], events: list[dict]) -> dict:
    now = datetime.now(CHINA_TZ)

    def in_days(item: dict, days: int) -> bool:
        parsed = parse_datetime_value(item.get("created_at"))
        return bool(parsed and parsed >= now - timedelta(days=days))

    success_reports = [item for item in reports if item.get("status") == "success"]
    event_types = Counter(item.get("event_type", "") for item in events)
    report_ips = {item.get("client_ip", "") for item in reports if item.get("client_ip")}
    event_ips = {item.get("client_ip", "") for item in events if item.get("client_ip")}
    report_user_ids = {item.get("user_id", "") for item in reports if item.get("user_id")}
    report_emails = {item.get("user_email", "") for item in reports if item.get("user_email")}

    daily_reports: Counter[str] = Counter()
    daily_visits: Counter[str] = Counter()
    for item in reports:
        daily_reports[day_key(item.get("created_at"))] += 1
    for item in events:
        if item.get("event_type") in ("page_view", "audit_submit", "audit_success", "audit_failed", "login_success", "register_success"):
            daily_visits[day_key(item.get("created_at"))] += 1

    recent_days = []
    daily_peak = max([daily_visits.get((now - timedelta(days=offset)).strftime("%m-%d"), 0) + daily_reports.get((now - timedelta(days=offset)).strftime("%m-%d"), 0) for offset in range(13, -1, -1)] or [0])
    for offset in range(13, -1, -1):
        label = (now - timedelta(days=offset)).strftime("%m-%d")
        visits = daily_visits.get(label, 0)
        reports_count = daily_reports.get(label, 0)
        recent_days.append(
            {
                "date": label,
                "visits": visits,
                "reports": reports_count,
                "bar_percent": round(((visits + reports_count) / daily_peak) * 100, 1) if daily_peak else 0,
            }
        )

    success_rate = round(len(success_reports) * 100 / len(reports), 1) if reports else 0
    total_original_bytes = sum(int(item.get("original_size_bytes") or 0) for item in reports)
    return {
        "events_enabled": bool(events),
        "page_views": event_types.get("page_view", 0),
        "unique_ips": len(report_ips | event_ips),
        "registered_users": len(users),
        "active_users": len(report_user_ids | report_emails),
        "total_reports": len(reports),
        "success_reports": len(success_reports),
        "success_rate": success_rate,
        "today_reports": sum(1 for item in reports if in_days(item, 1)),
        "week_reports": sum(1 for item in reports if in_days(item, 7)),
        "month_reports": sum(1 for item in reports if in_days(item, 30)),
        "today_events": sum(1 for item in events if in_days(item, 1)),
        "week_events": sum(1 for item in events if in_days(item, 7)),
        "month_events": sum(1 for item in events if in_days(item, 30)),
        "login_success": event_types.get("login_success", 0),
        "register_success": event_types.get("register_success", 0),
        "audit_submit": event_types.get("audit_submit", 0),
        "audit_success": event_types.get("audit_success", 0),
        "audit_failed": event_types.get("audit_failed", 0),
        "storage_gb": round(total_original_bytes / (1024 ** 3), 2),
        "daily": recent_days,
    }


def summarize_report_colleges(reports: list[dict]) -> dict:
    counts: defaultdict[str, int] = defaultdict(int)
    success_counts: defaultdict[str, int] = defaultdict(int)
    for item in reports:
        college = item.get("college_name") or UNKNOWN_COLLEGE
        counts[college] += 1
        if item.get("status") == "success":
            success_counts[college] += 1
    top_count = max(counts.values(), default=0)
    rows = [
        {
            "college": college,
            "count": count,
            "success": success_counts[college],
            "percent": round((count / len(reports)) * 100, 1) if reports else 0,
            "bar_percent": round((count / top_count) * 100, 1) if top_count else 0,
        }
        for college, count in counts.items()
    ]
    rows.sort(key=lambda item: (-item["count"], item["college"] == UNKNOWN_COLLEGE, item["college"]))
    return {
        "rows": rows,
        "top": rows[0] if rows else {"college": "暂无", "count": 0, "success": 0, "percent": 0},
        "unknown": counts.get(UNKNOWN_COLLEGE, 0),
        "known": len(reports) - counts.get(UNKNOWN_COLLEGE, 0),
        "options": [item["college"] for item in rows],
    }


def enrich_report_item(item: dict) -> dict:
    enriched = dict(item)
    enriched["created_at_display"] = format_datetime_display(item.get("created_at", ""))
    enriched["college_name"] = item.get("college_name") or UNKNOWN_COLLEGE
    enriched["college_source"] = item.get("college_source") or ""
    enriched["college_raw_text"] = item.get("college_raw_text") or ""
    enriched["original_size_display"] = format_bytes(int(item.get("original_size_bytes") or 0))
    enriched["report_size_display"] = format_bytes(int(item.get("report_size_bytes") or 0))
    enriched["original_sha_short"] = (item.get("original_sha256") or "")[:12]
    enriched["report_sha_short"] = (item.get("report_sha256") or "")[:12]
    return enriched


def format_bytes(size: int) -> str:
    if size <= 0:
        return ""
    units = ["B", "KB", "MB", "GB"]
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{size} B"


def report_table_state() -> dict:
    return {
        "q": request.args.get("q", "").strip(),
        "status": (request.args.get("status", "all") or "all").strip().lower(),
        "college": (request.args.get("college", "all") or "all").strip(),
        "page": parse_positive_int(request.args.get("page"), 1),
    }


def apply_report_filters(reports: list[dict], state: dict) -> list[dict]:
    keyword = state["q"].lower()
    status = state["status"]
    filtered: list[dict] = []
    for item in reports:
        haystack = " ".join(
            [
                item.get("user_email", ""),
                item.get("original_filename", ""),
                item.get("report_filename", ""),
                item.get("college_name", ""),
                item.get("college_source", ""),
                item.get("college_raw_text", ""),
                item.get("client_ip", ""),
                item.get("user_agent", ""),
                item.get("created_at", ""),
            ]
        ).lower()
        if keyword and keyword not in haystack:
            continue
        if status != "all" and item.get("status") != status:
            continue
        if state.get("college", "all") != "all" and item.get("college_name", UNKNOWN_COLLEGE) != state["college"]:
            continue
        filtered.append(item)
    return filtered


def record_admin_log(actor: dict | None, action: str, target: dict | None, summary: str, details: dict | None = None) -> None:
    try:
        (
            get_supabase()
            .table(ADMIN_LOG_TABLE)
            .insert(
                {
                    "actor_user_id": actor.get("id") if actor else None,
                    "actor_email": actor.get("email", "") if actor else "",
                    "action": action,
                    "target_user_id": target.get("id") if target else None,
                    "target_email": target.get("email", "") if target else "",
                    "summary": summary,
                    "details": details or {},
                }
            )
            .execute()
        )
    except Exception:
        app.logger.warning("Failed to record admin log for action %s", action, exc_info=True)


def update_user_status(user_id: str, status: str) -> None:
    if status not in {ACCOUNT_STATUS_ACTIVE, ACCOUNT_STATUS_FROZEN, ACCOUNT_STATUS_DISABLED}:
        raise ValueError("无效的账号状态。")
    user = find_user_by_id(user_id)
    if user is None:
        raise ValueError("用户不存在。")
    (
        get_supabase()
        .table(SUPABASE_TABLE)
        .update({"account_status": status})
        .eq("id", user_id)
        .execute()
    )


def update_user_admin(user_id: str, admin_enabled: bool) -> None:
    user = find_user_by_id(user_id)
    if user is None:
        raise ValueError("用户不存在。")
    (
        get_supabase()
        .table(SUPABASE_TABLE)
        .update({"is_admin": bool(admin_enabled)})
        .eq("id", user_id)
        .execute()
    )


def account_status_label(status: str) -> str:
    return {
        ACCOUNT_STATUS_ACTIVE: "正常",
        ACCOUNT_STATUS_FROZEN: "已冻结",
        ACCOUNT_STATUS_DISABLED: "已注销",
    }.get(status or ACCOUNT_STATUS_ACTIVE, "未知")


def is_account_active(user: dict | None) -> bool:
    return (user or {}).get("account_status", ACCOUNT_STATUS_ACTIVE) == ACCOUNT_STATUS_ACTIVE


def account_block_message(user: dict | None) -> str:
    status = (user or {}).get("account_status", ACCOUNT_STATUS_ACTIVE)
    if status == ACCOUNT_STATUS_FROZEN:
        return "这个账号已被冻结，请联系管理员处理。"
    if status == ACCOUNT_STATUS_DISABLED:
        return "这个账号已被注销，暂时不能继续使用。"
    return ""


def render_home(
    error: str = "",
    auth_error: str = "",
    auth_mode: str = "login",
    auth_values: dict | None = None,
) -> str:
    user = current_user()
    if "captcha_answer" not in session:
        refresh_captcha()
    auth_values = auth_values or {}
    invite_code = normalize_invite_code(user.get("invite_code", "")) if user else normalize_invite_code(request.values.get("invite", ""))
    invite_link = url_for("index", invite=invite_code, _external=True) if user and invite_code else ""
    if not user and invite_code and auth_mode == "login" and not auth_error:
        auth_mode = "register"
    return render_template_string(
        PAGE,
        user=user,
        configured=bool(SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY),
        remaining=remaining_submissions(user),
        max_submissions=user_quota(user),
        is_admin=is_admin(user),
        captcha_question=session.get("captcha_question", ""),
        captcha_left=session.get("captcha_left", ""),
        captcha_right=session.get("captcha_right", ""),
        auth_mode=auth_mode,
        auth_values=auth_values,
        auth_token=request.values.get("auth_token", ""),
        invite_code=invite_code,
        invite_link=invite_link,
        email_verification_enabled=email_verification_enabled(),
        email_code_target=session.get("email_code_target", ""),
        email_code_sent=email_code_remaining_seconds() > 0,
        email_code_remaining_seconds=email_code_remaining_seconds(),
        email_code_resend_seconds=EMAIL_CODE_RESEND_SECONDS,
        error=error,
        auth_error=auth_error,
    )


def refresh_captcha() -> None:
    left = random.randint(2, 9)
    right = random.randint(1, 8)
    session["captcha_left"] = str(left)
    session["captcha_right"] = str(right)
    session["captcha_question"] = f"{left} + {right} = ?"
    session["captcha_answer"] = str(left + right)


def registration_values(
    email: str,
    password: str,
    confirm_password: str,
    registration_code: str = "",
    invite_code: str = "",
    email_code: str = "",
) -> dict:
    return {
        "register_email": email,
        "register_password": password,
        "register_confirm_password": confirm_password,
        "registration_code": normalize_registration_code(registration_code),
        "invite_code": normalize_invite_code(invite_code),
        "email_code": (email_code or "").strip(),
    }


def login_values(email: str, password: str) -> dict:
    return {
        "login_email": email,
        "login_password": password,
    }


def is_valid_captcha(answer: str, left: str, right: str) -> bool:
    try:
        return int(answer.strip()) == int(left) + int(right)
    except (TypeError, ValueError):
        return False


PAGE = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>UPC本科论文格式检测工具</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #18212c;
      --muted: #657382;
      --paper: #f6f4ef;
      --surface: rgba(255, 255, 255, .82);
      --surface-strong: #ffffff;
      --line: #d6d8d2;
      --field: #fbfbf8;
      --accent: #1e7f62;
      --accent-strong: #145a45;
      --accent-soft: #dcefe7;
      --warn: #a64232;
      --shadow: rgba(24, 33, 44, .12);
      --grid: rgba(24, 33, 44, .045);
    }
    [data-theme="dark"] {
      color-scheme: dark;
      --ink: #edf2ef;
      --muted: #9cadb7;
      --paper: #0e1416;
      --surface: rgba(18, 27, 29, .82);
      --surface-strong: #151f22;
      --line: #2a393b;
      --field: #10191b;
      --accent: #62c598;
      --accent-strong: #8ce0b4;
      --accent-soft: #17362b;
      --warn: #ff9a84;
      --shadow: rgba(0, 0, 0, .34);
      --grid: rgba(237, 242, 239, .055);
    }
    * { box-sizing: border-box; }
    html { background: var(--paper); }
    body {
      margin: 0;
      min-height: 100svh;
      font-family: "Songti SC", "Noto Serif SC", "STSong", serif;
      color: var(--ink);
      background:
        radial-gradient(circle at 14% 20%, color-mix(in srgb, var(--accent) 20%, transparent), transparent 28rem),
        linear-gradient(120deg, color-mix(in srgb, var(--accent) 10%, transparent), transparent 42%),
        repeating-linear-gradient(0deg, var(--grid), var(--grid) 1px, transparent 1px, transparent 34px),
        var(--paper);
      transition: background .25s ease, color .25s ease;
    }
    main {
      width: min(1180px, calc(100% - 36px));
      margin: 0 auto;
      padding: 42px 0;
    }
    .topbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 34px;
      font: 700 13px/1.4 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .brand {
      display: flex;
      align-items: center;
      gap: 12px;
      color: var(--muted);
    }
    .brand-mark {
      width: 30px;
      height: 30px;
      display: grid;
      place-items: center;
      border: 1px solid var(--line);
      background: var(--surface);
      color: var(--accent);
      font-family: "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .theme-toggle {
      width: auto;
      margin: 0;
      padding: 10px 13px;
      border: 1px solid var(--line);
      background: var(--surface);
      color: var(--ink);
      box-shadow: none;
      font-size: 13px;
    }
    .theme-toggle:hover {
      background: var(--surface-strong);
      color: var(--accent-strong);
    }
    .shell {
      display: grid;
      grid-template-columns: minmax(0, 1.05fr) minmax(420px, .8fr);
      gap: clamp(34px, 6vw, 80px);
      align-items: center;
      min-height: calc(100svh - 146px);
    }
    .mark {
      width: 74px;
      height: 4px;
      margin-bottom: 30px;
      background: var(--accent);
    }
    h1 {
      margin: 0;
      max-width: 680px;
      font-size: clamp(48px, 8vw, 96px);
      line-height: .94;
      font-weight: 800;
      letter-spacing: 0;
    }
    .lead {
      max-width: 560px;
      margin: 28px 0 0;
      color: var(--muted);
      font-size: 18px;
      line-height: 1.8;
    }
    .panel {
      border: 1px solid var(--line);
      background: var(--surface);
      padding: 30px;
      box-shadow: 0 26px 90px var(--shadow);
      backdrop-filter: blur(18px);
      animation: rise .5s ease both;
    }
    .panel-title {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 22px;
      padding-bottom: 16px;
      border-bottom: 1px solid var(--line);
      font: 700 13px/1.4 "PingFang SC", "Noto Sans SC", sans-serif;
      color: var(--muted);
    }
    .panel-title strong {
      color: var(--ink);
      font-size: 20px;
    }
    label {
      display: block;
      margin-bottom: 12px;
      font: 700 15px/1.4 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    input[type="file"] {
      position: absolute;
      width: 1px;
      height: 1px;
      opacity: 0;
      pointer-events: none;
    }
    input[type="email"],
    input[type="password"],
    input[type="text"] {
      width: 100%;
      margin-bottom: 12px;
      padding: 14px 15px;
      border: 1px solid var(--line);
      background: var(--field);
      color: var(--ink);
      font: 15px/1.5 "PingFang SC", "Noto Sans SC", sans-serif;
      outline: 0;
      transition: border-color .18s ease, box-shadow .18s ease, background .18s ease;
    }
    input:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 4px color-mix(in srgb, var(--accent) 18%, transparent);
    }
    button {
      width: 100%;
      margin-top: 18px;
      border: 0;
      padding: 16px 18px;
      background: var(--accent);
      color: white;
      cursor: pointer;
      font: 700 16px/1 "PingFang SC", "Noto Sans SC", sans-serif;
      transition: transform .18s ease, background .18s ease;
    }
    button:hover { background: var(--accent-strong); transform: translateY(-1px); }
    button:disabled {
      cursor: wait;
      background: #7d928b;
      transform: none;
    }
    .note {
      margin: 18px 0 0;
      color: var(--muted);
      font: 13px/1.8 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .progress-wrap {
      display: none;
      margin-top: 18px;
      font: 13px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
      color: var(--muted);
    }
    .progress-wrap.active { display: block; }
    .progress-head {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 8px;
    }
    .progress-track {
      width: 100%;
      height: 8px;
      overflow: hidden;
      background: color-mix(in srgb, var(--line) 70%, transparent);
      border: 1px solid var(--line);
    }
    .progress-bar {
      width: 0%;
      height: 100%;
      background: linear-gradient(90deg, var(--accent), #54aa82);
      transition: width .45s ease;
    }
    .upload-card {
      display: grid;
      grid-template-columns: 52px 1fr;
      gap: 14px;
      align-items: center;
      min-height: 116px;
      margin-bottom: 16px;
      padding: 20px;
      border: 1px dashed color-mix(in srgb, var(--accent) 46%, var(--line));
      background:
        linear-gradient(135deg, color-mix(in srgb, var(--accent-soft) 34%, transparent), transparent 55%),
        var(--field);
      cursor: pointer;
      transition: border-color .18s ease, background .18s ease, transform .18s ease;
    }
    .upload-card:hover,
    .upload-card.dragging {
      border-color: var(--accent);
      background:
        linear-gradient(135deg, color-mix(in srgb, var(--accent-soft) 62%, transparent), transparent 55%),
        var(--surface-strong);
      transform: translateY(-1px);
    }
    .upload-icon {
      display: grid;
      place-items: center;
      width: 52px;
      height: 52px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      color: var(--accent);
      font: 900 24px/1 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .upload-title {
      margin: 0 0 5px;
      color: var(--ink);
      font: 800 16px/1.4 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .upload-meta {
      margin: 0;
      color: var(--muted);
      font: 13px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
      word-break: break-word;
    }
    .download-done {
      display: none;
      margin-top: 14px;
      padding: 12px 14px;
      border: 1px solid color-mix(in srgb, var(--accent) 36%, transparent);
      background: color-mix(in srgb, var(--accent-soft) 45%, transparent);
      color: var(--accent-strong);
      font: 700 13px/1.6 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .download-done.active { display: block; }
    .error {
      margin: 0 0 18px;
      padding: 12px 14px;
      border: 1px solid color-mix(in srgb, var(--warn) 35%, transparent);
      background: color-mix(in srgb, var(--warn) 10%, transparent);
      color: var(--warn);
      font: 700 14px/1.6 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .account-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      margin-bottom: 18px;
      padding-bottom: 14px;
      border-bottom: 1px solid var(--line);
      font: 13px/1.6 "PingFang SC", "Noto Sans SC", sans-serif;
      color: var(--muted);
    }
    .account-bar strong { color: var(--ink); }
    .logout-link {
      color: var(--accent);
      text-decoration: none;
      font-weight: 700;
      white-space: nowrap;
      margin-left: 12px;
    }
    .top-links-inline {
      display: inline-flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
    }
    .auth-grid {
      display: grid;
      grid-template-columns: 1fr;
      gap: 0;
    }
    .auth-switch {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      margin-bottom: 22px;
      padding: 5px;
      border: 1px solid var(--line);
      background: color-mix(in srgb, var(--field) 78%, transparent);
    }
    .auth-tab {
      margin: 0;
      padding: 11px 12px;
      border: 1px solid transparent;
      background: transparent;
      color: var(--muted);
      font-size: 14px;
      box-shadow: none;
    }
    .auth-tab.active {
      border-color: var(--line);
      background: var(--surface-strong);
      color: var(--ink);
    }
    .auth-tab:hover {
      transform: none;
      background: var(--surface-strong);
    }
    .auth-box {
      min-width: 0;
      display: none;
    }
    .auth-box.active {
      display: block;
      animation: rise .28s ease both;
    }
    .auth-box h2 {
      margin: 0 0 8px;
      font: 800 22px/1.3 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .auth-copy {
      margin: 0 0 18px;
      color: var(--muted);
      font: 13px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .auth-box button { margin-top: 4px; }
    .captcha-row {
      display: grid;
      grid-template-columns: 116px 1fr;
      gap: 10px;
      align-items: stretch;
    }
    .verify-row {
      display: grid;
      grid-template-columns: 1fr 148px;
      gap: 10px;
      align-items: stretch;
    }
    .verify-row input {
      margin-bottom: 0;
    }
    .verify-button {
      width: 100%;
      margin: 0;
      padding: 0 12px;
      font-size: 14px;
      box-shadow: none;
    }
    .verify-note {
      margin: 8px 0 12px;
      color: var(--muted);
      font: 12px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .captcha-chip {
      display: grid;
      place-items: center;
      margin-bottom: 12px;
      border: 1px solid var(--line);
      background: var(--accent-soft);
      color: var(--accent-strong);
      font: 800 15px/1 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .auth-rules {
      display: grid;
      gap: 10px;
      margin-top: 20px;
      padding-top: 18px;
      border-top: 1px solid var(--line);
    }
    .auth-rule {
      display: grid;
      grid-template-columns: 28px 1fr;
      gap: 10px;
      color: var(--muted);
      font: 13px/1.65 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .auth-rule span {
      display: grid;
      place-items: center;
      width: 28px;
      height: 28px;
      border: 1px solid var(--line);
      color: var(--accent);
      font-weight: 800;
    }
    .usage {
      margin: 0 0 16px;
      color: var(--muted);
      font: 14px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .quota-help {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 14px;
      align-items: center;
      margin: 0 0 20px;
      padding: 16px;
      border: 1px solid color-mix(in srgb, var(--accent) 54%, var(--line));
      background:
        linear-gradient(135deg, color-mix(in srgb, var(--accent) 22%, transparent), transparent 62%),
        color-mix(in srgb, var(--accent-soft) 58%, var(--surface));
      color: var(--accent-strong);
      box-shadow: 0 16px 42px color-mix(in srgb, var(--accent) 15%, transparent);
      font: 700 13px/1.6 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .quota-help p {
      margin: 0;
      color: var(--accent-strong);
    }
    .quota-actions {
      display: flex;
      flex-direction: column;
      align-items: flex-end;
      gap: 10px;
    }
    .quota-label {
      display: block;
      margin-bottom: 4px;
      color: var(--muted);
      font-size: 12px;
      letter-spacing: .08em;
      text-transform: uppercase;
    }
    .quota-number {
      display: inline-block;
      padding: 9px 12px;
      border: 1px solid color-mix(in srgb, var(--accent) 56%, var(--line));
      background: var(--surface-strong);
      color: var(--accent-strong);
      font: 900 22px/1 "PingFang SC", "Noto Sans SC", sans-serif;
      letter-spacing: .04em;
    }
    .copy-button {
      width: auto;
      margin: 0;
      padding: 11px 14px;
      border: 1px solid color-mix(in srgb, var(--accent) 56%, var(--line));
      background: var(--surface-strong);
      color: var(--accent-strong);
      box-shadow: none;
      font-size: 13px;
    }
    .copy-button:hover {
      background: color-mix(in srgb, var(--accent-soft) 55%, var(--surface-strong));
    }
    .group-invite {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 16px;
      align-items: center;
      margin: 0 0 22px;
      padding: 22px;
      border: 2px solid color-mix(in srgb, var(--accent) 62%, var(--line));
      background:
        linear-gradient(135deg, color-mix(in srgb, var(--accent-soft) 78%, transparent), transparent 74%),
        var(--surface-strong);
      color: var(--accent-strong);
      box-shadow: 0 20px 52px color-mix(in srgb, var(--accent) 18%, transparent);
      font: 800 14px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .group-invite strong {
      display: block;
      margin-bottom: 6px;
      color: var(--ink);
      font-size: 20px;
      line-height: 1.25;
    }
    .group-invite-head {
      display: flex;
      align-items: center;
      justify-content: flex-start;
      gap: 12px;
      margin-bottom: 6px;
    }
    .group-invite-head strong {
      margin: 0;
    }
    .group-number {
      margin: 8px 0 2px;
      color: var(--accent-strong);
      font: 900 30px/1.05 "PingFang SC", "Noto Sans SC", sans-serif;
      letter-spacing: .04em;
    }
    .group-copy-button {
      min-width: 148px;
      padding: 14px 18px;
      border-width: 2px;
      font-size: 15px;
    }
    .group-copy-button::before {
      content: "⧉";
      margin-right: 7px;
    }
    .invite-card {
      margin: 0 0 20px;
      padding: 16px;
      border: 1px solid color-mix(in srgb, var(--accent) 42%, var(--line));
      background:
        radial-gradient(circle at top right, color-mix(in srgb, var(--accent) 18%, transparent), transparent 12rem),
        var(--surface-strong);
      font: 700 13px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .invite-card p {
      margin: 0 0 12px;
      color: var(--muted);
    }
    .invite-code {
      display: inline-flex;
      align-items: center;
      min-height: 42px;
      padding: 8px 12px;
      border: 1px solid var(--line);
      background: var(--field);
      color: var(--ink);
      font: 900 20px/1 "PingFang SC", "Noto Sans SC", sans-serif;
      letter-spacing: .08em;
    }
    .invite-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
    }
    .modal {
      position: fixed;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 24px;
      background: rgba(15, 22, 27, .48);
      z-index: 30;
    }
    .modal.active {
      display: flex;
      animation: rise .22s ease both;
    }
    .modal-card {
      width: min(560px, 100%);
      border: 1px solid var(--line);
      background: var(--surface-strong);
      box-shadow: 0 24px 80px var(--shadow);
      padding: 24px;
    }
    .modal-card h3 {
      margin: 0 0 10px;
      font: 800 24px/1.2 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .modal-card p {
      margin: 0;
      color: var(--muted);
      font: 14px/1.8 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .modal-group-card {
      display: none;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 16px;
      align-items: center;
      margin-top: 18px;
      padding: 18px;
      border: 2px solid color-mix(in srgb, var(--accent) 58%, var(--line));
      background:
        radial-gradient(circle at top right, color-mix(in srgb, var(--accent) 22%, transparent), transparent 12rem),
        linear-gradient(135deg, color-mix(in srgb, var(--accent-soft) 72%, #fff), #fff 72%);
      box-shadow: inset 0 0 0 1px rgba(255, 255, 255, .7), 0 16px 42px rgba(23, 111, 88, .14);
    }
    .modal-group-card.active {
      display: grid;
    }
    .modal-group-kicker {
      color: var(--accent-strong);
      font-size: 12px;
      font-weight: 900;
      letter-spacing: .12em;
      text-transform: uppercase;
    }
    .modal-group-number {
      margin-top: 4px;
      color: var(--accent-strong);
      font: 950 36px/1 "PingFang SC", "Noto Sans SC", sans-serif;
      letter-spacing: .06em;
    }
    .modal-group-copy {
      min-width: 132px;
      padding: 14px 18px;
      border-width: 2px;
      font-size: 15px;
      box-shadow: 0 12px 30px rgba(23, 111, 88, .2);
    }
    .modal-group-note {
      grid-column: 1 / -1;
      margin: -6px 0 0;
      color: var(--muted);
      font: 13px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .modal-actions {
      display: flex;
      justify-content: flex-end;
      gap: 10px;
      margin-top: 20px;
    }
    .modal-actions button {
      width: auto;
      margin: 0;
      padding: 12px 16px;
    }
    .ghost-button {
      border: 1px solid var(--line);
      background: var(--surface);
      color: var(--ink);
    }
    .ghost-button:hover {
      background: var(--surface-strong);
      color: var(--accent-strong);
    }
    .facts {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 28px;
      padding: 0;
      list-style: none;
      font: 14px/1.5 "PingFang SC", "Noto Sans SC", sans-serif;
      color: var(--muted);
    }
    .facts li {
      border-top: 1px solid var(--line);
      padding-top: 10px;
      min-width: 142px;
    }
    @keyframes rise {
      from { opacity: 0; transform: translateY(14px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @media (max-width: 820px) {
      main { padding: 24px 0; }
      .topbar { margin-bottom: 30px; }
      .shell { grid-template-columns: 1fr; min-height: auto; }
      .panel { padding: 22px; }
      .auth-grid { grid-template-columns: 1fr; }
      .captcha-row { grid-template-columns: 1fr; }
      .verify-row { grid-template-columns: 1fr; }
      .quota-help { grid-template-columns: 1fr; }
      .quota-actions { align-items: flex-start; }
      .quota-number { width: fit-content; }
      .group-invite { grid-template-columns: 1fr; padding: 18px; }
      .group-invite-head { align-items: flex-start; flex-direction: column; }
      .group-number { font-size: 26px; }
      .group-copy-button { width: 100%; }
      .modal-card { padding: 20px; }
      .modal-group-card { grid-template-columns: 1fr; }
      .modal-group-number { font-size: 32px; }
      .modal-group-copy { width: 100%; }
      .modal-actions { flex-direction: column; }
      .modal-actions button { width: 100%; }
      h1 { font-size: clamp(46px, 16vw, 68px); }
    }
  </style>
</head>
<body>
  <main>
    <header class="topbar">
      <div class="brand">
        <span class="brand-mark">审</span>
        <span>UPC本科论文格式检测工具</span>
      </div>
      <button id="theme-toggle" class="theme-toggle" type="button" aria-label="切换夜间模式">夜间模式</button>
    </header>
    <section class="shell">
      <div>
        <div class="mark"></div>
        <h1>UPC本科论文格式检测工具</h1>
        <p class="lead">上传 Word 论文，系统会检查摘要、目录、标题、正文、图表、公式、参考文献和页码，并生成可交互的 HTML 报告。</p>
        <ul class="facts">
          <li>支持 .doc / .docx</li>
          <li>单文件 32MB 内</li>
          <li>检测过程不改原文</li>
        </ul>
      </div>
      <div class="panel">
        {% if not configured %}
          <div class="panel-title"><strong>系统配置</strong><span>Database</span></div>
          <p class="error">服务还没有配置 Supabase 数据库。</p>
          <p class="note">管理员需要设置 SUPABASE_URL 和 SUPABASE_SERVICE_ROLE_KEY 后才能开放注册登录。</p>
        {% elif user %}
          <div class="panel-title"><strong>生成报告</strong><span>{{ max_submissions }} 次额度</span></div>
          <div class="account-bar">
            <span>当前账号：<strong>{{ user["email"] }}</strong><br>剩余次数：<strong><span id="remaining-count">{{ remaining }}</span> 次</strong></span>
            <span class="top-links-inline">
              <a class="logout-link" href="{{ url_for('my_reports', auth_token=auth_token) }}">我的检测记录</a>
              {% if is_admin %}<a class="logout-link" href="{{ url_for('admin', auth_token=auth_token) }}">管理后台</a>{% endif %}
              <a class="logout-link" href="{{ url_for('logout') }}">退出登录</a>
            </span>
          </div>
          <div class="group-invite">
            <div>
              <strong>加入 QQ 群，领取更多检测机会</strong>
              <div class="group-number">537124215</div>
              进群后可联系管理员领取额外检测机会，也可以反馈检测问题。
            </div>
            <button class="copy-button group-copy-button" type="button" data-copy-group>复制群号</button>
          </div>
          <div class="invite-card">
            <p>邀请好友注册并完成创建账号，你的检测额度会自动增加 1 次。</p>
            <div class="invite-actions">
              <span class="invite-code">{{ invite_code }}</span>
              <button class="copy-button" type="button" data-copy-invite="{{ invite_link }}">复制邀请链接</button>
            </div>
          </div>
          <form id="audit-form" method="post" action="{{ url_for('audit') }}" enctype="multipart/form-data">
            {% if auth_token %}<input name="auth_token" type="hidden" value="{{ auth_token }}">{% endif %}
            {% if error %}<p class="error">{{ error }}</p>{% endif %}
            {% if remaining > 0 %}
              <p class="usage">每个账号最多可生成 {{ max_submissions }} 次报告。</p>
              <label for="docx">选择论文文件</label>
              <label id="upload-card" class="upload-card" for="docx">
                <span class="upload-icon">↑</span>
                <span>
                  <span id="upload-title" class="upload-title">点击选择 Word 论文</span>
                  <span id="upload-meta" class="upload-meta">支持 .doc / .docx，旧版 .doc 会自动转换后检测；生成完成后自动下载 HTML 报告。</span>
                </span>
              </label>
              <input id="docx" name="docx" type="file" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required>
              <button id="submit-button" type="submit">生成检测报告</button>
              <div id="progress-wrap" class="progress-wrap" role="status" aria-live="polite">
                <div class="progress-head">
                  <span id="progress-message">正在上传论文...</span>
                  <strong id="progress-percent">0%</strong>
                </div>
                <div class="progress-track" aria-hidden="true">
                  <div id="progress-bar" class="progress-bar"></div>
                </div>
              </div>
              <div id="download-done" class="download-done">报告已开始下载，可以继续选择新文件生成下一份报告。</div>
              <p class="note">报告会在浏览器中下载为 HTML 文件，可以直接打开或转发。大文件可能需要等待几十秒。</p>
            {% else %}
              <p class="error">这个账号的检测额度已经用完。</p>
            {% endif %}
          </form>
        {% else %}
          <div class="panel-title"><strong>开始使用</strong><span>账号限制</span></div>
          {% if auth_error %}<p class="error">{{ auth_error }}</p>{% endif %}
          <div class="group-invite">
            <div>
              <strong>加入 QQ 群，领取检测机会</strong>
              <div class="group-number">537124215</div>
              进群后可联系管理员领取额外检测机会，新用户也可以咨询注册和使用问题。
            </div>
            <button class="copy-button group-copy-button" type="button" data-copy-group>复制群号</button>
          </div>
          <div class="auth-switch" role="tablist" aria-label="登录或注册">
            <button class="auth-tab {% if auth_mode == 'login' %}active{% endif %}" type="button" data-auth-tab="login">登录</button>
            <button class="auth-tab {% if auth_mode == 'register' %}active{% endif %}" type="button" data-auth-tab="register">注册</button>
          </div>
          <div class="auth-grid">
            <form class="auth-box {% if auth_mode == 'login' %}active{% endif %}" method="post" action="{{ url_for('login') }}" data-auth-panel="login">
              <h2>登录</h2>
              <p class="auth-copy">使用已注册邮箱进入检测面板，系统会继续记录你的剩余次数。</p>
              <input name="email" type="email" placeholder="邮箱" autocomplete="email" value="{{ auth_values.get('login_email', '') }}" required>
              <input name="password" type="password" placeholder="密码" autocomplete="current-password" value="{{ auth_values.get('login_password', '') }}" required>
              <button type="submit">登录后检测</button>
            </form>
            <form class="auth-box {% if auth_mode == 'register' %}active{% endif %}" method="post" action="{{ url_for('register') }}" data-auth-panel="register">
              <h2>注册</h2>
              <p class="auth-copy">创建账号后可生成 {{ max_submissions }} 次报告。新注册仅支持 QQ 邮箱，并需要填写 QQ 群注册码。</p>
              <input name="registration_code" type="text" placeholder="QQ群注册码：加入官方 QQ 群 537124215，从群公告获得" value="{{ auth_values.get('registration_code', '') }}" required>
              <input name="email" type="email" placeholder="QQ 邮箱，例如 123456@qq.com" autocomplete="email" value="{{ auth_values.get('register_email', '') }}" required>
              {% if email_verification_enabled %}
                <div class="verify-row">
                  <input id="register-email-code" name="email_code" type="text" inputmode="numeric" pattern="[0-9]*" maxlength="6" placeholder="输入邮箱验证码" value="{{ auth_values.get('email_code', '') }}" required>
                  <button
                    id="send-email-code"
                    class="verify-button"
                    type="button"
                    data-email-target="{{ email_code_target }}"
                    data-resend-seconds="{{ email_code_remaining_seconds }}"
                  >
                    {% if email_code_sent %}{{ email_code_remaining_seconds }} 秒后重发{% else %}发送验证码{% endif %}
                  </button>
                </div>
                <p id="email-code-note" class="verify-note">
                  {% if email_code_sent %}
                    验证码已发送到 {{ email_code_target }}，10 分钟内有效。若收不到，请查看垃圾邮箱或广告邮件。
                  {% else %}
                    使用 QQ 邮箱接收 6 位验证码，发送前请先填写邮箱。若收不到验证码，请查看垃圾邮箱或广告邮件。
                  {% endif %}
                </p>
              {% else %}
                <p class="verify-note">当前邮箱验证码服务尚未配置，管理员配置 Gmail 后会自动启用。</p>
              {% endif %}
              <input name="password" type="password" placeholder="至少 6 位密码" autocomplete="new-password" minlength="6" value="{{ auth_values.get('register_password', '') }}" required>
              <input name="confirm_password" type="password" placeholder="再次输入密码" autocomplete="new-password" minlength="6" value="{{ auth_values.get('register_confirm_password', '') }}" required>
              <input name="invite_code" type="text" placeholder="邀请码（可选）" value="{{ auth_values.get('invite_code', invite_code) }}">
              <div class="captcha-row">
                <div class="captcha-chip">{{ captcha_question }}</div>
                <input name="captcha_left" type="hidden" value="{{ captcha_left }}">
                <input name="captcha_right" type="hidden" value="{{ captcha_right }}">
                <input name="captcha_answer" type="text" inputmode="numeric" pattern="[0-9]*" placeholder="输入计算结果" required>
              </div>
              <button type="submit">创建账号</button>
            </form>
          </div>
          <div class="auth-rules">
            <div class="auth-rule"><span>1</span><p>QQ群注册码从官方 QQ 群 537124215 的群公告中获得，用于控制注册入口。</p></div>
            <div class="auth-rule"><span>2</span><p>每个账号最多生成 {{ max_submissions }} 次报告，次数保存在数据库中。</p></div>
            <div class="auth-rule"><span>3</span><p>数字验证只用于减少自动注册，不会收集额外信息。</p></div>
          </div>
        {% endif %}
      </div>
    </section>
  </main>
  <div id="info-modal" class="modal" aria-hidden="true">
    <div class="modal-card" role="dialog" aria-modal="true" aria-labelledby="modal-title" aria-describedby="modal-body">
      <h3 id="modal-title">提示</h3>
      <p id="modal-body"></p>
      <div id="modal-group-card" class="modal-group-card">
        <div>
          <div class="modal-group-kicker">官方 QQ 群</div>
          <div class="modal-group-number">537124215</div>
        </div>
        <button id="modal-group-copy" class="copy-button modal-group-copy" type="button">复制群号</button>
        <div id="modal-group-note" class="modal-group-note">进群后可领取额外检测机会，也可以反馈检测问题。</div>
      </div>
      <div class="modal-actions">
        <button id="modal-secondary" class="ghost-button" type="button" hidden>关闭</button>
        <button id="modal-primary" type="button">我知道了</button>
      </div>
    </div>
  </div>
  <script>
    const GROUP_NUMBER = '537124215';
    const root = document.documentElement;
    const themeToggle = document.getElementById('theme-toggle');
    const savedTheme = localStorage.getItem('theme');
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const initialTheme = savedTheme || (prefersDark ? 'dark' : 'light');
    root.dataset.theme = initialTheme;
    if (themeToggle) themeToggle.textContent = initialTheme === 'dark' ? '日间模式' : '夜间模式';

    if (themeToggle) themeToggle.addEventListener('click', () => {
      const nextTheme = root.dataset.theme === 'dark' ? 'light' : 'dark';
      root.dataset.theme = nextTheme;
      localStorage.setItem('theme', nextTheme);
      themeToggle.textContent = nextTheme === 'dark' ? '日间模式' : '夜间模式';
    });

    document.querySelectorAll('[data-auth-tab]').forEach(tabButton => {
      tabButton.addEventListener('click', () => {
        const target = tabButton.dataset.authTab;
        document.querySelectorAll('[data-auth-tab]').forEach(button => {
          button.classList.toggle('active', button === tabButton);
        });
        document.querySelectorAll('[data-auth-panel]').forEach(panel => {
          panel.classList.toggle('active', panel.dataset.authPanel === target);
        });
      });
    });

    const form = document.getElementById('audit-form');
    const fileInput = document.getElementById('docx');
    const submitButton = document.getElementById('submit-button');
    const registerEmailInput = document.querySelector('[data-auth-panel="register"] input[name="email"]');
    const registerEmailCodeInput = document.getElementById('register-email-code');
    const sendEmailCodeButton = document.getElementById('send-email-code');
    const emailCodeNote = document.getElementById('email-code-note');
    const progressWrap = document.getElementById('progress-wrap');
    const progressBar = document.getElementById('progress-bar');
    const progressPercent = document.getElementById('progress-percent');
    const progressMessage = document.getElementById('progress-message');
    const uploadCard = document.getElementById('upload-card');
    const uploadTitle = document.getElementById('upload-title');
    const uploadMeta = document.getElementById('upload-meta');
    const downloadDone = document.getElementById('download-done');
    const remainingCount = document.getElementById('remaining-count');
    const infoModal = document.getElementById('info-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalBody = document.getElementById('modal-body');
    const modalPrimary = document.getElementById('modal-primary');
    const modalSecondary = document.getElementById('modal-secondary');
    const modalGroupCard = document.getElementById('modal-group-card');
    const modalGroupCopy = document.getElementById('modal-group-copy');
    const modalGroupNote = document.getElementById('modal-group-note');
    let modalPrimaryHandler = null;
    let modalSecondaryHandler = null;

    const messages = [
      [10, '正在上传论文...'],
      [28, '正在读取 Word 结构...'],
      [46, '正在检查摘要、目录和标题...'],
      [64, '正在检查正文、图表和公式...'],
      [82, '正在生成 HTML 报告...'],
      [92, '报告快好了，请稍等...']
    ];
    let resendCountdownTimer = null;

    const startEmailCodeCountdown = seconds => {
      if (!sendEmailCodeButton) return;
      window.clearInterval(resendCountdownTimer);
      let remainingSeconds = Number(seconds || 0);
      const paint = () => {
        if (remainingSeconds > 0) {
          sendEmailCodeButton.disabled = true;
          sendEmailCodeButton.textContent = `${remainingSeconds} 秒后重发`;
        } else {
          sendEmailCodeButton.disabled = false;
          sendEmailCodeButton.textContent = '发送验证码';
          window.clearInterval(resendCountdownTimer);
        }
      };
      paint();
      if (remainingSeconds <= 0) return;
      resendCountdownTimer = window.setInterval(() => {
        remainingSeconds -= 1;
        paint();
      }, 1000);
    };

    const closeModal = () => {
      if (!infoModal) return;
      infoModal.classList.remove('active');
      infoModal.setAttribute('aria-hidden', 'true');
      modalPrimaryHandler = null;
      modalSecondaryHandler = null;
    };

    const showModal = ({
      title,
      body,
      primaryText = '我知道了',
      secondaryText = '',
      showGroup = false,
      groupNote = '进群后可领取额外检测机会，也可以反馈检测问题。',
      onPrimary = null,
      onSecondary = null
    }) => {
      if (!infoModal || !modalTitle || !modalBody || !modalPrimary || !modalSecondary) return;
      modalTitle.textContent = title;
      modalBody.textContent = body;
      modalPrimary.textContent = primaryText;
      if (modalGroupCard) {
        modalGroupCard.classList.toggle('active', Boolean(showGroup));
      }
      if (modalGroupNote) {
        modalGroupNote.textContent = groupNote;
      }
      modalPrimaryHandler = onPrimary;
      modalSecondaryHandler = onSecondary;
      if (secondaryText) {
        modalSecondary.hidden = false;
        modalSecondary.textContent = secondaryText;
      } else {
        modalSecondary.hidden = true;
      }
      infoModal.classList.add('active');
      infoModal.setAttribute('aria-hidden', 'false');
    };

    const copyText = async text => {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text);
        return;
      }
      const helper = document.createElement('input');
      helper.value = text;
      document.body.appendChild(helper);
      helper.select();
      document.execCommand('copy');
      helper.remove();
    };

    const showQuotaExhaustedModal = source => {
      const body = source === 'download'
        ? '本次报告已经下载成功，但你的检测额度也已经用完。下面是补充检测机会的官方联系入口。'
        : '你的检测额度已经用完。下面是补充检测机会的官方联系入口。';
      showModal({
        title: '检测额度已用完',
        body,
        showGroup: true,
        groupNote: '复制群号后打开 QQ 搜索加入，进群可领取新的检测机会。',
        primaryText: '复制群号',
        secondaryText: '稍后再说',
        onPrimary: async () => {
          try {
            await copyText(GROUP_NUMBER);
            showModal({
              title: '群号已复制',
              body: `QQ群号 ${GROUP_NUMBER} 已复制到剪贴板，打开 QQ 搜索群号即可申请加入。`
            });
          } catch (_error) {
            showModal({
              title: '复制失败',
              body: `浏览器暂时无法自动复制，请手动复制群号 ${GROUP_NUMBER}。`
            });
          }
        }
      });
    };

    const showPostAuditReminderModal = () => {
      showModal({
        title: '下载成功',
        body: '检测报告已经下载成功，请到浏览器下载列表或下载文件夹中找到该 HTML 文件，并用浏览器打开查看结果。',
        showGroup: true,
        groupNote: '遇到使用问题或想领取更多检测机会，可以复制群号加入官方 QQ 群。',
        primaryText: '复制群号',
        secondaryText: '我知道了',
        onPrimary: async () => {
          try {
            await copyText(GROUP_NUMBER);
            showModal({
              title: '群号已复制',
              body: `QQ群号 ${GROUP_NUMBER} 已复制到剪贴板，打开 QQ 搜索群号即可申请加入。`
            });
          } catch (_error) {
            showModal({
              title: '复制失败',
              body: `浏览器暂时无法自动复制，请手动复制群号 ${GROUP_NUMBER}。`
            });
          }
        }
      });
    };

    if (modalPrimary) modalPrimary.addEventListener('click', () => {
      if (modalPrimaryHandler) modalPrimaryHandler();
      closeModal();
    });

    if (modalSecondary) modalSecondary.addEventListener('click', () => {
      if (modalSecondaryHandler) modalSecondaryHandler();
      closeModal();
    });

    if (infoModal) infoModal.addEventListener('click', event => {
      if (event.target === infoModal) closeModal();
    });

    document.addEventListener('keydown', event => {
      if (event.key === 'Escape' && infoModal && infoModal.classList.contains('active')) closeModal();
    });

    const copyGroupNumber = async () => {
      try {
        await copyText(GROUP_NUMBER);
        showModal({
          title: '群号已复制',
          body: `QQ群号 ${GROUP_NUMBER} 已复制到剪贴板，打开 QQ 搜索群号即可申请加入。`
        });
      } catch (_error) {
        showModal({
          title: '复制失败',
          body: `浏览器暂时无法自动复制，请手动复制群号 ${GROUP_NUMBER}。`
        });
      }
    };

    document.querySelectorAll('[data-copy-group]').forEach(button => {
      button.addEventListener('click', copyGroupNumber);
    });

    if (modalGroupCopy) {
      modalGroupCopy.addEventListener('click', copyGroupNumber);
    }

    document.querySelectorAll('[data-copy-invite]').forEach(button => {
      button.addEventListener('click', async () => {
        const inviteLink = button.dataset.copyInvite || '';
        try {
          await copyText(inviteLink);
          showModal({
            title: '邀请链接已复制',
            body: '好友通过你的链接注册成功后，你会自动增加 1 次检测额度。'
          });
        } catch (_error) {
          showModal({
            title: '复制失败',
            body: `浏览器暂时无法自动复制，请手动复制：${inviteLink}`
          });
        }
      });
    });

    if (sendEmailCodeButton) {
      startEmailCodeCountdown(Number(sendEmailCodeButton.dataset.resendSeconds || '0'));
      sendEmailCodeButton.addEventListener('click', async () => {
        const email = (registerEmailInput?.value || '').trim();
        if (!email) {
          showModal({
            title: '先填写邮箱',
            body: '请先输入你要注册的邮箱，再发送验证码。'
          });
          registerEmailInput?.focus();
          return;
        }

        sendEmailCodeButton.disabled = true;
        sendEmailCodeButton.textContent = '发送中...';
        try {
          const payload = new FormData();
          payload.append('email', email);
          const response = await fetch('{{ url_for("send_register_email_code") }}', {
            method: 'POST',
            body: payload,
            credentials: 'same-origin',
            headers: {
              'Accept': 'application/json',
              'X-Requested-With': 'fetch'
            }
          });
          const responseText = await response.text();
          let data = {};
          try {
            data = responseText ? JSON.parse(responseText) : {};
          } catch (_error) {
            data = {
              ok: false,
              message: response.ok
                ? '服务器返回格式异常，请稍后重试。'
                : '验证码服务暂时不可用，请稍后重试。'
            };
          }
          if (!response.ok || data.ok === false) {
            throw new Error(data.message || '发送验证码失败，请稍后再试。');
          }
          if (emailCodeNote) {
            emailCodeNote.textContent = `验证码已发送到 ${email}，10 分钟内有效。若收不到，请查看垃圾邮箱或广告邮件。`;
          }
          startEmailCodeCountdown(Number(data.resend_seconds || {{ email_code_resend_seconds }}));
          showModal({
            title: '验证码已发送',
            body: `我们已经把 6 位验证码发送到了 ${email}，请到邮箱中查看。若收不到验证码，请查看垃圾邮箱或广告邮件。`
          });
          registerEmailCodeInput?.focus();
        } catch (error) {
          sendEmailCodeButton.disabled = false;
          sendEmailCodeButton.textContent = '发送验证码';
          showModal({
            title: '发送失败',
            body: error.message || '邮箱验证码发送失败，请稍后再试。'
          });
        }
      });
    }

    if (fileInput && uploadTitle && uploadMeta) fileInput.addEventListener('change', () => {
      const file = fileInput.files[0];
      if (!file) return;
      uploadTitle.textContent = file.name;
      uploadMeta.textContent = `${(file.size / 1024 / 1024).toFixed(2)} MB · 已选择，点击下方按钮开始检测`;
      if (downloadDone) downloadDone.classList.remove('active');
    });

    if (uploadCard && fileInput) {
      ['dragenter', 'dragover'].forEach(eventName => {
        uploadCard.addEventListener(eventName, event => {
          event.preventDefault();
          uploadCard.classList.add('dragging');
        });
      });
      ['dragleave', 'drop'].forEach(eventName => {
        uploadCard.addEventListener(eventName, event => {
          event.preventDefault();
          uploadCard.classList.remove('dragging');
        });
      });
      uploadCard.addEventListener('drop', event => {
        const file = event.dataTransfer.files[0];
        if (!file) return;
        const transfer = new DataTransfer();
        transfer.items.add(file);
        fileInput.files = transfer.files;
        fileInput.dispatchEvent(new Event('change', { bubbles: true }));
      });
    }

    if (remainingCount && Number(remainingCount.textContent) <= 0) {
      window.setTimeout(() => showQuotaExhaustedModal('page'), 220);
    }

    if (form && fileInput && submitButton && progressWrap) form.addEventListener('submit', async event => {
      event.preventDefault();
      if (!fileInput.files.length) return;

      submitButton.disabled = true;
      submitButton.textContent = '检测中，请稍等...';
      progressWrap.classList.add('active');
      if (downloadDone) downloadDone.classList.remove('active');

      let progress = 0;
      let finished = false;
      const tick = () => {
        const nextLimit = progress < 30 ? 30 : progress < 70 ? 70 : 92;
        const step = progress < 30 ? 6 : progress < 70 ? 3 : 1;
        progress = Math.min(nextLimit, progress + step);
        progressBar.style.width = `${progress}%`;
        progressPercent.textContent = `${progress}%`;

        const current = [...messages].reverse().find(([limit]) => progress >= limit);
        if (current) progressMessage.textContent = current[1];
      };

      const finishDownloadState = () => {
        if (finished) return;
        finished = true;
        progress = 100;
        progressBar.style.width = '100%';
        progressPercent.textContent = '100%';
        progressMessage.textContent = '报告已下载成功';
        const noRemaining = remainingCount && Number(remainingCount.textContent) <= 0;
        submitButton.disabled = Boolean(noRemaining);
        submitButton.textContent = noRemaining ? '额度已用完' : '继续生成报告';
        if (downloadDone) {
          downloadDone.textContent = '报告已下载成功，请在浏览器中打开下载的 HTML 文件查看检测结果。';
          downloadDone.classList.add('active');
        }
      };

      const extractFilename = response => {
        const disposition = response.headers.get('Content-Disposition') || '';
        const encodedMatch = disposition.match(/filename\\*=UTF-8''([^;]+)/i);
        if (encodedMatch) return decodeURIComponent(encodedMatch[1]);
        const normalMatch = disposition.match(/filename="?([^";]+)"?/i);
        return normalMatch ? normalMatch[1] : 'thesis_format_audit_report.html';
      };

      const updateRemainingCount = response => {
        const remaining = response.headers.get('X-Remaining-Submissions');
        if (remainingCount && remaining !== null) remainingCount.textContent = remaining;
      };

      tick();
      const progressTimer = window.setInterval(() => {
        if (finished) {
          window.clearInterval(progressTimer);
          return;
        }
        tick();
      }, 900);

      try {
        const response = await fetch(form.action, {
          method: 'POST',
          body: new FormData(form),
          credentials: 'same-origin',
          headers: {
            'X-Requested-With': 'fetch'
          }
        });
        if (!response.ok) {
          const message = await response.text();
          throw new Error(message || '检测失败，请稍后再试。');
        }
        const blob = await response.blob();
        const downloadUrl = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = extractFilename(response);
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(downloadUrl);
        updateRemainingCount(response);
        finishDownloadState();
        if (remainingCount && Number(remainingCount.textContent) <= 0) {
          showQuotaExhaustedModal('download');
        } else {
          showPostAuditReminderModal();
        }
      } catch (error) {
        finished = true;
        window.clearInterval(progressTimer);
        progressBar.style.width = '0%';
        progressPercent.textContent = '0%';
        progressMessage.textContent = error.message || '检测失败，请稍后再试。';
        submitButton.disabled = false;
        submitButton.textContent = '重新生成报告';
      }
    });
  </script>
</body>
</html>
"""


ADMIN_PAGE = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>管理后台 - UPC本科论文格式检测工具</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #16202a;
      --muted: #607181;
      --paper: #f3efe6;
      --surface: rgba(255, 255, 255, .86);
      --surface-strong: #ffffff;
      --line: #d7d9d2;
      --line-strong: #c3c9bf;
      --accent: #176f58;
      --accent-strong: #0f4d3d;
      --accent-soft: #d8ece5;
      --warn: #9d4131;
      --warn-soft: #f7e0db;
      --info: #2f5f8a;
      --info-soft: #dceaf8;
      --gold: #966d1f;
      --gold-soft: #f3e7c8;
      --shadow: rgba(22, 32, 42, .12);
    }
    * { box-sizing: border-box; }
    html { background: var(--paper); }
    body {
      margin: 0;
      min-height: 100svh;
      color: var(--ink);
      background:
        radial-gradient(circle at 8% 12%, color-mix(in srgb, var(--accent) 18%, transparent), transparent 28rem),
        radial-gradient(circle at 92% 16%, color-mix(in srgb, #caa55a 14%, transparent), transparent 22rem),
        linear-gradient(180deg, rgba(255, 255, 255, .35), transparent 26%),
        var(--paper);
      font-family: "PingFang SC", "Noto Sans SC", sans-serif;
    }
    main {
      width: min(1320px, calc(100% - 40px));
      margin: 0 auto;
      padding: 36px 0 48px;
    }
    .topbar, .user-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }
    .topbar {
      align-items: flex-start;
      margin-bottom: 24px;
    }
    .eyebrow {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 16px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: .16em;
      text-transform: uppercase;
    }
    .eyebrow::before {
      content: "";
      width: 44px;
      height: 1px;
      background: var(--accent);
    }
    h1 {
      margin: 0;
      font-size: clamp(36px, 5vw, 64px);
      line-height: .98;
      letter-spacing: -.03em;
    }
    .headline-copy {
      margin: 14px 0 0;
      max-width: 720px;
      color: var(--muted);
      font-size: 16px;
      line-height: 1.8;
    }
    a {
      color: var(--accent);
      font-weight: 800;
      text-decoration: none;
    }
    .top-actions {
      display: flex;
      flex-direction: column;
      align-items: flex-end;
      gap: 12px;
    }
    .user-row {
      color: var(--muted);
      font-size: 14px;
    }
    .top-links {
      display: flex;
      gap: 14px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }
    .top-link {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 10px 14px;
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 10px 28px rgba(22, 32, 42, .06);
    }
    .stats {
      display: grid;
      grid-template-columns: repeat(7, minmax(0, 1fr));
      gap: 16px;
      margin-bottom: 20px;
    }
    .traffic-panel {
      margin-bottom: 20px;
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      overflow: hidden;
      backdrop-filter: blur(16px);
    }
    .traffic-head {
      display: flex;
      justify-content: space-between;
      gap: 18px;
      padding: 20px 22px;
      border-bottom: 1px solid var(--line);
    }
    .traffic-head h2 {
      margin: 0 0 8px;
      font-size: 24px;
      letter-spacing: -.02em;
    }
    .traffic-grid {
      display: grid;
      grid-template-columns: repeat(6, minmax(0, 1fr));
      gap: 12px;
      padding: 18px;
      border-bottom: 1px solid var(--line);
    }
    .traffic-card {
      min-width: 0;
      padding: 16px;
      border: 1px solid color-mix(in srgb, var(--line) 78%, transparent);
      background: rgba(255, 255, 255, .64);
    }
    .traffic-value {
      display: block;
      font-size: 28px;
      font-weight: 900;
      letter-spacing: -.02em;
      line-height: 1;
    }
    .traffic-label {
      display: block;
      margin-top: 9px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 900;
      letter-spacing: .08em;
      text-transform: uppercase;
    }
    .traffic-body {
      display: grid;
      grid-template-columns: minmax(0, 1.1fr) minmax(260px, .7fr);
      gap: 18px;
      padding: 18px;
    }
    .traffic-days {
      display: grid;
      gap: 8px;
    }
    .traffic-day {
      display: grid;
      grid-template-columns: 56px minmax(0, 1fr) 112px;
      gap: 10px;
      align-items: center;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
    }
    .traffic-track {
      height: 9px;
      overflow: hidden;
      background: color-mix(in srgb, var(--line) 70%, transparent);
    }
    .traffic-fill {
      width: var(--bar, 0%);
      height: 100%;
      background: linear-gradient(90deg, var(--accent), color-mix(in srgb, var(--gold) 70%, var(--accent)));
    }
    .traffic-note {
      padding: 16px;
      border: 1px solid color-mix(in srgb, var(--accent) 22%, var(--line));
      background: color-mix(in srgb, var(--accent-soft) 42%, transparent);
      color: var(--accent-strong);
      font-size: 13px;
      line-height: 1.8;
    }
    .stat-card {
      padding: 20px 22px;
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 18px 56px rgba(22, 32, 42, .08);
      backdrop-filter: blur(14px);
    }
    .stat-label {
      display: block;
      margin-bottom: 10px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: .12em;
      text-transform: uppercase;
    }
    .stat-value {
      display: block;
      font-size: 34px;
      font-weight: 900;
      line-height: 1;
    }
    .stat-hint {
      margin-top: 8px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.6;
    }
    .panel {
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      overflow: hidden;
      backdrop-filter: blur(16px);
    }
    .notice {
      margin: 0 0 18px;
      padding: 13px 15px;
      border: 1px solid color-mix(in srgb, var(--accent) 38%, var(--line));
      background: color-mix(in srgb, var(--accent-soft) 58%, transparent);
      color: var(--accent);
      font-weight: 800;
    }
    .notice[hidden] {
      display: none;
    }
    .toolbar {
      display: grid;
      grid-template-columns: minmax(0, 1.4fr) repeat(4, minmax(140px, .38fr)) auto;
      gap: 12px;
      padding: 18px;
      border-bottom: 1px solid var(--line);
      background: rgba(255, 255, 255, .5);
    }
    .field,
    .select,
    .inline-number {
      width: 100%;
      padding: 13px 14px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      color: var(--ink);
      font: 14px/1.4 "PingFang SC", "Noto Sans SC", sans-serif;
      outline: 0;
      transition: border-color .18s ease, box-shadow .18s ease;
    }
    .field:focus,
    .select:focus,
    .inline-number:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 4px color-mix(in srgb, var(--accent) 16%, transparent);
    }
    .summary-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      padding: 14px 18px;
      border-bottom: 1px solid var(--line);
      color: var(--muted);
      font-size: 13px;
    }
    .toolbar-actions {
      display: flex;
      gap: 10px;
      align-items: center;
    }
    .user-list {
      display: grid;
      gap: 14px;
      padding: 18px;
      background:
        radial-gradient(circle at 100% 0%, color-mix(in srgb, var(--accent-soft) 54%, transparent), transparent 22rem),
        linear-gradient(180deg, rgba(255, 255, 255, .4), transparent 12rem);
    }
    .user-card {
      display: grid;
      grid-template-columns: minmax(230px, .95fr) minmax(180px, .62fr) minmax(190px, .68fr) minmax(280px, 1fr) minmax(340px, 1.2fr);
      gap: 0;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 22px;
      background: rgba(255, 255, 255, .82);
      box-shadow: 0 18px 52px rgba(22, 32, 42, .08);
      transition: transform .16s ease, border-color .16s ease, box-shadow .16s ease;
    }
    .user-card:hover {
      transform: translateY(-2px);
      border-color: color-mix(in srgb, var(--accent) 26%, var(--line));
      box-shadow: 0 24px 70px rgba(22, 32, 42, .12);
    }
    .user-card-section {
      display: grid;
      align-content: center;
      gap: 12px;
      min-width: 0;
      padding: 18px;
      border-right: 1px solid color-mix(in srgb, var(--line) 78%, transparent);
    }
    .user-card-section:last-child {
      border-right: 0;
    }
    .section-label {
      color: var(--muted);
      font-size: 11px;
      font-weight: 900;
      letter-spacing: .13em;
      text-transform: uppercase;
    }
    .account-cell {
      display: grid;
      gap: 10px;
      min-width: 0;
    }
    .email {
      font-weight: 900;
      letter-spacing: -.01em;
    }
    .id-chip {
      display: inline-flex;
      width: fit-content;
      max-width: 220px;
      padding: 6px 8px;
      border: 1px solid color-mix(in srgb, var(--line) 80%, transparent);
      background: color-mix(in srgb, var(--paper) 42%, #fff);
      color: var(--muted);
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .muted { color: var(--muted); }
    .mono {
      font-family: "SFMono-Regular", "Menlo", monospace;
      font-size: 12px;
    }
    .profile-cell {
      display: grid;
      gap: 12px;
      min-width: 190px;
    }
    .badge-stack {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }
    .quota {
      display: inline-grid;
      place-items: center;
      min-width: 78px;
      padding: 8px 10px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      font-weight: 900;
    }
    .quota-card {
      display: grid;
      gap: 8px;
      min-width: 0;
    }
    .quota-head {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 12px;
    }
    .remaining-count {
      font-size: 28px;
      font-weight: 950;
      line-height: 1;
      color: var(--accent-strong);
    }
    .quota-label {
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
    }
    .quota-track {
      height: 8px;
      overflow: hidden;
      border-radius: 999px;
      background: color-mix(in srgb, var(--line) 72%, transparent);
    }
    .quota-fill {
      width: var(--quota-percent, 0%);
      height: 100%;
      border-radius: inherit;
      background: linear-gradient(90deg, var(--accent), color-mix(in srgb, var(--accent) 55%, #d1a84b));
      transition: width .2s ease;
    }
    .quota-foot {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      color: var(--muted);
      font-size: 12px;
    }
    .invite-cell {
      display: grid;
      gap: 8px;
      min-width: 0;
    }
    .trace-cell {
      display: grid;
      gap: 8px;
      min-width: 0;
    }
    .trace-line {
      display: grid;
      grid-template-columns: 72px minmax(0, 1fr);
      gap: 4px 10px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
    }
    .trace-line strong {
      color: var(--ink);
      font-size: 12px;
      white-space: nowrap;
    }
    .trace-main {
      min-width: 0;
    }
    .trace-main span {
      display: block;
    }
    .ua {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .invite-code-mini {
      display: inline-flex;
      width: fit-content;
      padding: 7px 9px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      font-size: 12px;
      font-weight: 900;
      letter-spacing: .08em;
    }
    .mini-link {
      width: fit-content;
      border: 0;
      padding: 0;
      background: transparent;
      color: var(--accent);
      font-size: 12px;
      font-weight: 900;
      cursor: pointer;
    }
    .mini-link:hover {
      filter: none;
      transform: none;
      text-decoration: underline;
    }
    .status-badge {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 78px;
      padding: 8px 10px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: .04em;
    }
    .status-active {
      border-color: color-mix(in srgb, var(--accent) 30%, var(--line));
      background: color-mix(in srgb, var(--accent-soft) 42%, transparent);
      color: var(--accent-strong);
    }
    .status-frozen {
      border-color: color-mix(in srgb, var(--info) 30%, var(--line));
      background: color-mix(in srgb, var(--info-soft) 60%, transparent);
      color: var(--info);
    }
    .status-disabled {
      border-color: color-mix(in srgb, var(--warn) 34%, var(--line));
      background: color-mix(in srgb, var(--warn-soft) 72%, transparent);
      color: var(--warn);
    }
    .role-badge {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 86px;
      padding: 8px 10px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      font-size: 12px;
      font-weight: 900;
      letter-spacing: .04em;
    }
    .role-super {
      border-color: color-mix(in srgb, var(--gold) 42%, var(--line));
      background: color-mix(in srgb, var(--gold-soft) 76%, transparent);
      color: var(--gold);
    }
    .role-admin {
      border-color: color-mix(in srgb, var(--accent) 30%, var(--line));
      background: color-mix(in srgb, var(--accent-soft) 42%, transparent);
      color: var(--accent-strong);
    }
    .role-user {
      color: var(--muted);
    }
    .actions {
      display: grid;
      gap: 10px;
      min-width: 0;
    }
    .quota-action-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .action-group {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 6px;
    }
    .action-section {
      display: grid;
      gap: 7px;
      padding: 10px;
      border: 1px solid color-mix(in srgb, var(--line) 78%, transparent);
      border-radius: 14px;
      background: color-mix(in srgb, var(--paper) 24%, #fff);
    }
    .action-title {
      color: var(--muted);
      font-size: 11px;
      font-weight: 900;
      letter-spacing: .08em;
      text-transform: uppercase;
    }
    .inline-form {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 7px;
      align-items: center;
    }
    .account-actions {
      display: flex;
      gap: 7px;
      flex-wrap: wrap;
    }
    button {
      border: 1px solid transparent;
      padding: 10px 12px;
      border-radius: 12px;
      background: var(--accent);
      color: white;
      cursor: pointer;
      font-weight: 800;
      transition: transform .16s ease, filter .16s ease, background .16s ease, box-shadow .16s ease;
    }
    button:hover {
      filter: brightness(.95);
      transform: translateY(-1px);
      box-shadow: 0 10px 24px rgba(22, 32, 42, .12);
    }
    button:disabled {
      cursor: not-allowed;
      filter: grayscale(.25);
      opacity: .62;
      transform: none;
    }
    .ghost-button {
      background: var(--surface-strong);
      color: var(--ink);
      border-color: var(--line);
    }
    .ghost-button:hover {
      color: var(--accent-strong);
      background: color-mix(in srgb, var(--accent-soft) 36%, var(--surface-strong));
    }
    .warn-button {
      background: var(--warn);
    }
    .info-button {
      background: var(--info);
    }
    .compact-button {
      padding: 9px 10px;
      font-size: 13px;
    }
    .inline-number {
      width: 100%;
      padding: 9px 10px;
    }
    .pager {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      padding: 18px;
      border-top: 1px solid var(--line);
      background: rgba(255, 255, 255, .36);
    }
    .pager-info {
      color: var(--muted);
      font-size: 13px;
    }
    .pager-links {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 8px;
    }
    .pager-link {
      display: inline-flex;
      min-width: 40px;
      align-items: center;
      justify-content: center;
      padding: 10px 12px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      color: var(--ink);
      font-size: 13px;
      font-weight: 800;
      text-decoration: none;
    }
    .pager-link.active {
      border-color: var(--accent);
      background: var(--accent);
      color: #fff;
    }
    .log-preview {
      margin-top: 18px;
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      backdrop-filter: blur(16px);
    }
    .code-panel {
      margin-bottom: 18px;
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      backdrop-filter: blur(16px);
    }
    .code-panel-body {
      display: grid;
      grid-template-columns: minmax(280px, .9fr) minmax(0, 1.4fr);
      gap: 18px;
      padding: 18px;
    }
    .code-create {
      display: grid;
      gap: 10px;
      align-content: start;
      padding: 16px;
      border: 1px solid color-mix(in srgb, var(--line) 78%, transparent);
      border-radius: 18px;
      background: rgba(255, 255, 255, .58);
    }
    .code-create .inline-form {
      grid-template-columns: 110px minmax(0, 1fr);
    }
    .code-list {
      display: grid;
      gap: 10px;
    }
    .code-item {
      display: grid;
      grid-template-columns: minmax(160px, .8fr) minmax(0, 1fr) auto;
      gap: 12px;
      align-items: center;
      padding: 14px;
      border: 1px solid color-mix(in srgb, var(--line) 78%, transparent);
      border-radius: 18px;
      background: rgba(255, 255, 255, .68);
    }
    .code-token {
      display: inline-flex;
      width: fit-content;
      padding: 9px 11px;
      border: 1px solid color-mix(in srgb, var(--accent) 32%, var(--line));
      border-radius: 12px;
      background: color-mix(in srgb, var(--accent-soft) 50%, #fff);
      color: var(--accent-strong);
      font-family: "SFMono-Regular", "Menlo", monospace;
      font-size: 14px;
      font-weight: 900;
      letter-spacing: .08em;
    }
    .code-meta {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.7;
    }
    .code-status {
      display: inline-flex;
      width: fit-content;
      padding: 6px 9px;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: var(--surface-strong);
      font-size: 12px;
      font-weight: 900;
    }
    .code-status.active {
      border-color: color-mix(in srgb, var(--accent) 34%, var(--line));
      background: color-mix(in srgb, var(--accent-soft) 52%, transparent);
      color: var(--accent-strong);
    }
    .code-status.inactive {
      border-color: color-mix(in srgb, var(--warn) 34%, var(--line));
      background: color-mix(in srgb, var(--warn-soft) 72%, transparent);
      color: var(--warn);
    }
    .section-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 18px;
      border-bottom: 1px solid var(--line);
    }
    .section-head h2 {
      margin: 0;
      font-size: 22px;
      line-height: 1.2;
    }
    .log-list {
      display: grid;
      gap: 0;
    }
    .log-item {
      display: grid;
      grid-template-columns: minmax(0, 1.2fr) minmax(0, .8fr) auto;
      gap: 16px;
      padding: 16px 18px;
      border-bottom: 1px solid var(--line);
    }
    .log-item:last-child {
      border-bottom: 0;
    }
    .log-item strong {
      display: block;
      margin-bottom: 4px;
    }
    .log-meta {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.7;
    }
    .empty-state {
      padding: 42px 18px;
      color: var(--muted);
      text-align: center;
      font-size: 15px;
      line-height: 1.8;
    }
    @media (max-width: 1180px) {
      .stats {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
      .traffic-grid {
        grid-template-columns: repeat(3, minmax(0, 1fr));
      }
      .traffic-body {
        grid-template-columns: 1fr;
      }
      .toolbar {
        grid-template-columns: 1fr 1fr 1fr;
      }
      .user-card {
        grid-template-columns: minmax(220px, 1fr) minmax(180px, .8fr);
      }
      .user-card-section {
        border-right: 0;
        border-bottom: 1px solid color-mix(in srgb, var(--line) 78%, transparent);
      }
      .user-card-section:nth-child(odd) {
        border-right: 1px solid color-mix(in srgb, var(--line) 78%, transparent);
      }
      .user-card-section:last-child {
        grid-column: 1 / -1;
        border-bottom: 0;
        border-right: 0;
      }
      .actions {
        grid-template-columns: minmax(0, 1fr);
      }
    }
    @media (max-width: 820px) {
      main {
        width: min(100%, calc(100% - 24px));
        padding-top: 24px;
      }
      .topbar { align-items: flex-start; flex-direction: column; }
      .top-actions {
        width: 100%;
        align-items: flex-start;
      }
      .top-links {
        justify-content: flex-start;
      }
      .stats {
        grid-template-columns: 1fr;
      }
      .traffic-head,
      .traffic-body {
        grid-template-columns: 1fr;
      }
      .traffic-head {
        flex-direction: column;
      }
      .traffic-grid {
        grid-template-columns: 1fr 1fr;
      }
      .traffic-day {
        grid-template-columns: 52px minmax(0, 1fr);
      }
      .traffic-day span:last-child {
        grid-column: 2;
      }
      .toolbar {
        grid-template-columns: 1fr;
      }
      .summary-bar {
        align-items: flex-start;
        flex-direction: column;
      }
      .code-panel-body,
      .code-item {
        grid-template-columns: 1fr;
      }
      .toolbar-actions,
      .pager,
      .section-head,
      .log-item {
        align-items: flex-start;
        flex-direction: column;
        grid-template-columns: 1fr;
      }
      .user-list {
        padding: 12px;
      }
      .user-card {
        grid-template-columns: 1fr;
        border-radius: 18px;
      }
      .user-card-section,
      .user-card-section:nth-child(odd) {
        border-right: 0;
        border-bottom: 1px solid color-mix(in srgb, var(--line) 78%, transparent);
      }
      .user-card-section:last-child {
        border-bottom: 0;
      }
      .quota-action-grid {
        grid-template-columns: 1fr;
      }
      .actions,
      .action-group,
      .inline-form {
        width: 100%;
      }
      .inline-number {
        width: 100%;
      }
    }
  </style>
</head>
<body>
  <main>
    <div class="topbar">
      <div>
        <div class="eyebrow">Admin Console</div>
        <h1>管理后台</h1>
        <p class="headline-copy">在这里集中管理用户额度、账号状态和使用情况。支持搜索、筛选、冻结、恢复、注销以及快速发放检测次数。</p>
      </div>
      <div class="top-actions">
        <div class="user-row">
          <span class="muted">管理员：{{ admin_user["email"] }}</span>
        </div>
        <div class="top-links">
          <a class="top-link" href="{{ url_for('admin_reports', auth_token=auth_token) }}">检测记录</a>
          <a class="top-link" href="{{ url_for('admin_logs', auth_token=auth_token) }}">操作日志</a>
          <a class="top-link" href="{{ url_for('index', auth_token=auth_token) }}">返回检测页</a>
          <a class="top-link" href="{{ url_for('logout') }}">退出登录</a>
        </div>
      </div>
    </div>
    <div id="admin-notice" class="notice" {% if not message %}hidden{% endif %}>{{ message }}</div>
    <section class="traffic-panel">
      <div class="traffic-head">
        <div>
          <h2>流量与使用频率</h2>
          <div class="log-meta">面向广告合作的数据看板：访问、人流、检测量、活跃用户和转化情况。新访问日志上线前的历史 PV 会偏少，检测量和用户量来自已有记录。</div>
        </div>
        <div class="traffic-note">
          {% if traffic_stats["events_enabled"] %}
            已启用访问事件追踪，后续会持续记录页面访问、注册、登录和检测转化。
          {% else %}
            访问事件表暂无数据；当前先展示历史检测和用户数据，部署后会开始积累真实 PV/UV。
          {% endif %}
        </div>
      </div>
      <div class="traffic-grid">
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["page_views"] }}</span><span class="traffic-label">页面访问 PV</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["unique_ips"] }}</span><span class="traffic-label">独立 IP / 人流</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["registered_users"] }}</span><span class="traffic-label">注册用户</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["active_users"] }}</span><span class="traffic-label">检测用户</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["total_reports"] }}</span><span class="traffic-label">累计检测</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["success_rate"] }}%</span><span class="traffic-label">检测成功率</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["today_reports"] }}</span><span class="traffic-label">近 24 小时检测</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["week_reports"] }}</span><span class="traffic-label">近 7 天检测</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["month_reports"] }}</span><span class="traffic-label">近 30 天检测</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["login_success"] }}</span><span class="traffic-label">登录成功</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["register_success"] }}</span><span class="traffic-label">注册成功</span></div>
        <div class="traffic-card"><span class="traffic-value">{{ traffic_stats["storage_gb"] }}GB</span><span class="traffic-label">原文归档量</span></div>
      </div>
      <div class="traffic-body">
        <div>
          <div class="section-label">近 14 天趋势</div>
          <div class="traffic-days">
            {% for day in traffic_stats["daily"] %}
              <div class="traffic-day">
                <span>{{ day["date"] }}</span>
                <div class="traffic-track"><div class="traffic-fill" style="--bar: {{ day['bar_percent'] }}%;"></div></div>
                <span>访问 {{ day["visits"] }} · 检测 {{ day["reports"] }}</span>
              </div>
            {% endfor %}
          </div>
        </div>
        <div class="traffic-note">
          <strong>给广告商可用口径</strong><br>
          注册用户 {{ traffic_stats["registered_users"] }} 人，累计检测 {{ traffic_stats["total_reports"] }} 次，近 7 天检测 {{ traffic_stats["week_reports"] }} 次，近 30 天检测 {{ traffic_stats["month_reports"] }} 次。独立 IP 当前按检测记录和访问事件合并统计，为 {{ traffic_stats["unique_ips"] }}。
        </div>
      </div>
    </section>
    <section class="stats">
      <div class="stat-card">
        <span class="stat-label">总用户数</span>
        <span class="stat-value">{{ stats["total"] }}</span>
        <div class="stat-hint">当前数据库中的全部账号</div>
      </div>
      <div class="stat-card">
        <span class="stat-label">正常账号</span>
        <span class="stat-value">{{ stats["active"] }}</span>
        <div class="stat-hint">可正常登录并生成检测报告</div>
      </div>
      <div class="stat-card">
        <span class="stat-label">冻结账号</span>
        <span class="stat-value">{{ stats["frozen"] }}</span>
        <div class="stat-hint">已被临时停用，可随时恢复</div>
      </div>
      <div class="stat-card">
        <span class="stat-label">已注销账号</span>
        <span class="stat-value">{{ stats["disabled"] }}</span>
        <div class="stat-hint">已停用，不允许继续使用</div>
      </div>
      <div class="stat-card">
        <span class="stat-label">管理员</span>
        <span class="stat-value">{{ stats["admins"] }}</span>
        <div class="stat-hint">含最高管理员和普通管理员</div>
      </div>
      <div class="stat-card">
        <span class="stat-label">邀请注册</span>
        <span class="stat-value">{{ stats["invited"] }}</span>
        <div class="stat-hint">通过邀请码注册的新用户</div>
      </div>
      <div class="stat-card">
        <span class="stat-label">奖励次数</span>
        <span class="stat-value">{{ stats["invite_rewards"] }}</span>
        <div class="stat-hint">已发放的邀请奖励次数</div>
      </div>
    </section>
    <section class="code-panel">
      <div class="section-head">
        <div>
          <h2>QQ群注册码</h2>
          <div class="log-meta">用户注册必须填写注册码。把可用注册码发布到官方 QQ 群 537124215 的群公告中，用户从群公告获取后再注册。</div>
        </div>
      </div>
      <div class="code-panel-body">
        <form class="code-create" method="post" action="{{ url_for('admin_create_registration_code') }}">
          <input name="auth_token" type="hidden" value="{{ auth_token }}">
          <input name="next" type="hidden" value="{{ current_url }}">
          <strong>生成新注册码</strong>
          <div class="log-meta">建议每次群公告放 1 个码，可设置使用次数；用完自动不可用。</div>
          <label class="log-meta" for="code-max-uses">可用次数</label>
          <div class="inline-form">
            <input id="code-max-uses" class="inline-number" name="max_uses" type="number" min="1" max="999" step="1" value="20" required>
            <button class="compact-button" type="submit">生成注册码</button>
          </div>
          <input class="field" name="note" type="text" maxlength="120" placeholder="备注，例如：QQ群公告 5月30日">
        </form>
        <div class="code-list">
          {% if registration_codes %}
            {% for code in registration_codes %}
              {% set remaining_uses = (code["max_uses"]|int) - (code["used_count"]|int) %}
              <div class="code-item">
                <div>
                  <button class="code-token mini-link" type="button" data-copy-text="{{ code['code'] }}">{{ code["code"] }}</button>
                  <div class="code-meta">点击复制，发布到 QQ 群公告</div>
                </div>
                <div class="code-meta">
                  <strong>{{ code["note"] or "未填写备注" }}</strong><br>
                  使用 {{ code["used_count"] }} / {{ code["max_uses"] }}，剩余 {{ remaining_uses if remaining_uses > 0 else 0 }}<br>
                  创建者：{{ code["created_by"] or "未知" }} · {{ code["created_at"] }}
                </div>
                <form method="post" action="{{ url_for('admin_update_registration_code_status') }}">
                  <input name="auth_token" type="hidden" value="{{ auth_token }}">
                  <input name="next" type="hidden" value="{{ current_url }}">
                  <input name="code_id" type="hidden" value="{{ code["id"] }}">
                  {% if code["is_active"] %}
                    <input name="is_active" type="hidden" value="0">
                    <button class="ghost-button compact-button" type="submit">停用</button>
                  {% else %}
                    <input name="is_active" type="hidden" value="1">
                    <button class="compact-button" type="submit">启用</button>
                  {% endif %}
                  <span class="code-status {% if code['is_active'] %}active{% else %}inactive{% endif %}">
                    {% if code["is_active"] %}启用中{% else %}已停用{% endif %}
                  </span>
                </form>
              </div>
            {% endfor %}
          {% else %}
            <div class="empty-state">还没有注册码。生成一个后放到 QQ 群公告中，用户注册时必须填写。</div>
          {% endif %}
        </div>
      </div>
    </section>
    <section class="panel">
      <form class="toolbar" method="get" action="{{ url_for('admin') }}">
        <input name="auth_token" type="hidden" value="{{ auth_token }}">
        <input name="page" type="hidden" value="1">
        <input name="q" class="field" type="search" placeholder="搜索邮箱、用户 ID 或注册时间" value="{{ table_state['q'] }}">
        <select name="status" class="select">
          <option value="all">全部状态</option>
          <option value="active" {% if table_state['status'] == 'active' %}selected{% endif %}>仅看正常</option>
          <option value="frozen" {% if table_state['status'] == 'frozen' %}selected{% endif %}>仅看冻结</option>
          <option value="disabled" {% if table_state['status'] == 'disabled' %}selected{% endif %}>仅看已注销</option>
        </select>
        <select name="quota" class="select">
          <option value="all">全部额度情况</option>
          <option value="remaining" {% if table_state['quota'] == 'remaining' %}selected{% endif %}>仅看仍有次数</option>
          <option value="empty" {% if table_state['quota'] == 'empty' %}selected{% endif %}>仅看额度用完</option>
        </select>
        <select name="sort" class="select">
          {% for option in sort_options %}
            <option value="{{ option['value'] }}" {% if table_state['sort'] == option['value'] %}selected{% endif %}>{{ option['label'] }}</option>
          {% endfor %}
        </select>
        <select name="per_page" class="select">
          {% for size in per_page_options %}
            <option value="{{ size }}" {% if table_state['per_page'] == size %}selected{% endif %}>每页 {{ size }} 条</option>
          {% endfor %}
        </select>
        <div class="toolbar-actions">
          <button type="submit">应用筛选</button>
          <a class="top-link" href="{{ reset_url }}">重置</a>
        </div>
      </form>
      <div class="summary-bar">
        <span>共筛选出 <strong>{{ pagination['total'] }}</strong> 个账号，当前显示第 {{ pagination['start'] }} - {{ pagination['end'] }} 条</span>
        <span>支持快速额度发放/扣减、自定义调整、冻结、恢复、注销和权限管理</span>
      </div>
      <div id="user-table" class="user-list">
        {% for item in users %}
          {% set used = item["submissions_used"]|int %}
          {% set total = item["submission_quota"]|int %}
          {% set quota_percent = (used * 100 / total) if total > 0 else 0 %}
          <article class="user-card" data-user-id="{{ item['id'] }}">
            <section class="user-card-section">
              <div class="section-label">账号</div>
                  <div class="account-cell">
                    <div>
                      <div class="email">{{ item["email"] }}</div>
                      <div class="mono id-chip" title="{{ item['id'] }}">{{ item["id"] }}</div>
                    </div>
                    <div class="badge-stack">
                      {% if item["is_super_admin"] %}
                        <span class="role-badge role-super">最高管理员</span>
                      {% elif item["is_admin"] %}
                        <span class="role-badge role-admin">管理员</span>
                      {% else %}
                        <span class="role-badge role-user">普通用户</span>
                      {% endif %}
                      <span class="status-badge status-{{ item['account_status'] }}">
                        {{ status_labels.get(item["account_status"], "未知") }}
                      </span>
                    </div>
                  </div>
            </section>
            <section class="user-card-section">
              <div class="section-label">状态与额度</div>
                  <div class="quota-card">
                    <div class="quota-head">
                      <div>
                        <div class="quota-label">剩余次数</div>
                        <div class="remaining-count" data-remaining-text>{{ item["remaining"] }}</div>
                      </div>
                      <span class="quota" data-quota-text>{{ used }} / {{ total }}</span>
                    </div>
                    <div class="quota-track" aria-label="额度使用进度">
                      <div class="quota-fill" data-quota-fill style="--quota-percent: {{ quota_percent if quota_percent <= 100 else 100 }}%;"></div>
                    </div>
                    <div class="quota-foot">
                      <span>已用 {{ used }}</span>
                      <span>总额 {{ total }}</span>
                    </div>
                  </div>
            </section>
            <section class="user-card-section">
              <div class="section-label">邀请</div>
                  <div class="invite-cell">
                    {% if item["invite_code"] %}
                      <span class="invite-code-mini">{{ item["invite_code"] }}</span>
                      <span class="muted">邀请 {{ item["invite_count"] }} 人</span>
                      {% if item["invited_by_email"] %}<span class="muted">来自：{{ item["invited_by_email"] }}</span>{% endif %}
                      <button class="mini-link" type="button" data-copy-text="{{ item['invite_link'] }}">复制邀请链接</button>
                    {% else %}
                      <span class="muted">暂无邀请码</span>
                    {% endif %}
                  </div>
            </section>
            <section class="user-card-section">
              <div class="section-label">来源追踪</div>
                  <div class="trace-cell">
                    <div class="trace-line">
                      <strong>注册</strong>
                      <div class="trace-main">
                        <span>{{ item["register_ip"] or "旧账号未记录" }}</span>
                        <span>{{ item["created_at_display"] or item["created_at"] }}</span>
                        {% if item["register_user_agent"] %}<span class="ua" title="{{ item['register_user_agent'] }}">{{ item["register_user_agent"] }}</span>{% endif %}
                      </div>
                    </div>
                    <div class="trace-line">
                      <strong>登录</strong>
                      <div class="trace-main">
                        <span>{{ item["last_login_ip"] or "暂无" }}</span>
                        <span>{{ item["last_login_at_display"] or "暂无登录记录" }}</span>
                        {% if item["last_login_user_agent"] %}<span class="ua" title="{{ item['last_login_user_agent'] }}">{{ item["last_login_user_agent"] }}</span>{% endif %}
                      </div>
                    </div>
                    <div class="trace-line">
                      <strong>检测</strong>
                      <div class="trace-main">
                        <span>{{ item["last_audit_ip"] or "暂无" }}</span>
                        <span>{{ item["last_audit_at_display"] or "暂无检测记录" }}</span>
                        {% if item["last_audit_user_agent"] %}<span class="ua" title="{{ item['last_audit_user_agent'] }}">{{ item["last_audit_user_agent"] }}</span>{% endif %}
                      </div>
                    </div>
                  </div>
                  <div class="muted">注册时间：{{ item["created_at_display"] or item["created_at"] }}</div>
            </section>
            <section class="user-card-section">
              <div class="section-label">管理操作</div>
                  <div class="actions">
                    <div class="quota-action-grid">
                      <div class="action-section">
                        <div class="action-title">快速加次</div>
                        <div class="action-group">
                          {% for amount in [1, 3, 10] %}
                            <form class="quota-form" method="post" action="{{ url_for('admin_add_quota') }}">
                              <input name="auth_token" type="hidden" value="{{ auth_token }}">
                              <input name="next" type="hidden" value="{{ current_url }}">
                              <input name="user_id" type="hidden" value="{{ item["id"] }}">
                              <input name="amount" type="hidden" value="{{ amount }}">
                              <button class="compact-button" type="submit">+{{ amount }}</button>
                            </form>
                          {% endfor %}
                        </div>
                      </div>
                      <div class="action-section">
                        <div class="action-title">快速扣减</div>
                        <div class="action-group">
                          {% for amount in [1, 3, 10] %}
                            <form class="quota-form" method="post" action="{{ url_for('admin_reduce_quota') }}">
                              <input name="auth_token" type="hidden" value="{{ auth_token }}">
                              <input name="next" type="hidden" value="{{ current_url }}">
                              <input name="user_id" type="hidden" value="{{ item["id"] }}">
                              <input name="amount" type="hidden" value="{{ amount }}">
                              <button class="ghost-button compact-button" type="submit">-{{ amount }}</button>
                            </form>
                          {% endfor %}
                        </div>
                      </div>
                    </div>
                    <div class="action-section">
                      <div class="action-title">自定义额度</div>
                      <form class="inline-form quota-form" method="post" action="{{ url_for('admin_add_quota') }}">
                        <input name="auth_token" type="hidden" value="{{ auth_token }}">
                        <input name="next" type="hidden" value="{{ current_url }}">
                        <input name="user_id" type="hidden" value="{{ item["id"] }}">
                        <input class="inline-number" name="amount" type="number" min="1" step="1" placeholder="输入次数">
                        <button class="ghost-button compact-button" type="submit">增加</button>
                      </form>
                      <form class="inline-form quota-form" method="post" action="{{ url_for('admin_reduce_quota') }}">
                        <input name="auth_token" type="hidden" value="{{ auth_token }}">
                        <input name="next" type="hidden" value="{{ current_url }}">
                        <input name="user_id" type="hidden" value="{{ item["id"] }}">
                        <input class="inline-number" name="amount" type="number" min="1" step="1" placeholder="输入次数">
                        <button class="ghost-button compact-button" type="submit">减少</button>
                      </form>
                    </div>
                    <div class="action-section">
                      <div class="action-title">账号状态</div>
                      <div class="account-actions">
                        {% if can_manage_admins %}
                          {% if item["is_super_admin"] %}
                            <button class="ghost-button compact-button" type="button" disabled>最高权限</button>
                          {% elif item["is_admin"] %}
                            <form method="post" action="{{ url_for('admin_update_role') }}" data-confirm="确认取消 {{ item['email'] }} 的管理员权限吗？">
                              <input name="auth_token" type="hidden" value="{{ auth_token }}">
                              <input name="next" type="hidden" value="{{ current_url }}">
                              <input name="user_id" type="hidden" value="{{ item["id"] }}">
                              <input name="is_admin" type="hidden" value="0">
                              <button class="ghost-button compact-button" type="submit">取消管理员</button>
                            </form>
                          {% else %}
                            <form method="post" action="{{ url_for('admin_update_role') }}" data-confirm="确认把 {{ item['email'] }} 设为管理员吗？管理员可以进入后台管理用户额度和账号状态。">
                              <input name="auth_token" type="hidden" value="{{ auth_token }}">
                              <input name="next" type="hidden" value="{{ current_url }}">
                              <input name="user_id" type="hidden" value="{{ item["id"] }}">
                              <input name="is_admin" type="hidden" value="1">
                              <button class="compact-button" type="submit">设为管理员</button>
                            </form>
                          {% endif %}
                        {% endif %}
                        {% if item["account_status"] == "active" %}
                          <form method="post" action="{{ url_for('admin_update_status') }}" data-confirm="确认冻结 {{ item['email'] }} 吗？冻结后该账号将不能登录和检测。">
                            <input name="auth_token" type="hidden" value="{{ auth_token }}">
                            <input name="next" type="hidden" value="{{ current_url }}">
                            <input name="user_id" type="hidden" value="{{ item["id"] }}">
                            <input name="status" type="hidden" value="frozen">
                            <button class="info-button compact-button" type="submit">冻结</button>
                          </form>
                        {% else %}
                          <form method="post" action="{{ url_for('admin_update_status') }}" data-confirm="确认恢复 {{ item['email'] }} 吗？恢复后账号可继续使用。">
                            <input name="auth_token" type="hidden" value="{{ auth_token }}">
                            <input name="next" type="hidden" value="{{ current_url }}">
                            <input name="user_id" type="hidden" value="{{ item["id"] }}">
                            <input name="status" type="hidden" value="active">
                            <button class="ghost-button compact-button" type="submit">恢复</button>
                          </form>
                        {% endif %}
                        {% if item["account_status"] != "disabled" %}
                          <form method="post" action="{{ url_for('admin_update_status') }}" data-confirm-email="{{ item['email'] }}">
                            <input name="auth_token" type="hidden" value="{{ auth_token }}">
                            <input name="next" type="hidden" value="{{ current_url }}">
                            <input name="user_id" type="hidden" value="{{ item["id"] }}">
                            <input name="status" type="hidden" value="disabled">
                            <input name="confirm_email" type="hidden" value="">
                            <button class="warn-button compact-button" type="submit">注销</button>
                          </form>
                        {% endif %}
                      </div>
                    </div>
                  </div>
            </section>
          </article>
        {% endfor %}
      </div>
      {% if not users %}
        <div class="empty-state">没有匹配到符合条件的账号，试试清空搜索词或调整筛选条件。</div>
      {% endif %}
      <div class="pager">
        <div class="pager-info">第 {{ pagination['page'] }} / {{ pagination['pages'] }} 页</div>
        <div class="pager-links">
          {% if pagination['page'] > 1 %}
            <a class="pager-link" href="{{ build_admin_url(auth_token, table_state, page=pagination['page'] - 1) }}">上一页</a>
          {% endif %}
          {% for page_number in page_numbers %}
            <a class="pager-link {% if page_number == pagination['page'] %}active{% endif %}" href="{{ build_admin_url(auth_token, table_state, page=page_number) }}">{{ page_number }}</a>
          {% endfor %}
          {% if pagination['page'] < pagination['pages'] %}
            <a class="pager-link" href="{{ build_admin_url(auth_token, table_state, page=pagination['page'] + 1) }}">下一页</a>
          {% endif %}
        </div>
      </div>
    </section>
    <section class="log-preview">
      <div class="section-head">
        <div>
          <h2>最近操作</h2>
          <div class="log-meta">最近几次关键后台动作，方便快速核对是否有误操作。</div>
        </div>
        <a class="top-link" href="{{ url_for('admin_logs', auth_token=auth_token) }}">查看全部日志</a>
      </div>
      {% if recent_logs %}
        <div class="log-list">
          {% for log in recent_logs %}
            <div class="log-item">
              <div>
                <strong>{{ log['summary'] }}</strong>
                <div class="log-meta">操作者：{{ log['actor_email'] or '未知管理员' }}</div>
              </div>
              <div class="log-meta">
                目标账号：{{ log['target_email'] or '无' }}<br>
                动作类型：{{ log['action'] }}
              </div>
              <div class="log-meta">{{ log['created_at'] }}</div>
            </div>
          {% endfor %}
        </div>
      {% else %}
        <div class="empty-state">还没有后台操作日志。</div>
      {% endif %}
    </section>
  </main>
  <script>
    const userTable = document.getElementById('user-table');
    const adminNotice = document.getElementById('admin-notice');

    const copyText = async text => {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text);
        return;
      }
      const helper = document.createElement('input');
      helper.value = text;
      document.body.appendChild(helper);
      helper.select();
      document.execCommand('copy');
      helper.remove();
    };

    const showAdminNotice = (message, isError = false) => {
      if (!adminNotice) return;
      adminNotice.textContent = message;
      adminNotice.hidden = false;
      adminNotice.style.color = isError ? 'var(--warn)' : 'var(--accent)';
      adminNotice.style.borderColor = isError
        ? 'color-mix(in srgb, var(--warn) 38%, var(--line))'
        : 'color-mix(in srgb, var(--accent) 38%, var(--line))';
      adminNotice.style.background = isError
        ? 'color-mix(in srgb, var(--warn-soft) 72%, transparent)'
        : 'color-mix(in srgb, var(--accent-soft) 58%, transparent)';
    };

    document.querySelectorAll('.quota-form').forEach(form => {
      form.addEventListener('submit', async event => {
        event.preventDefault();
        const button = form.querySelector('button[type="submit"]');
        const originalText = button?.textContent || '';
        if (button) {
          button.disabled = true;
          button.textContent = '处理中';
        }

        try {
          const response = await fetch(form.action, {
            method: 'POST',
            body: new FormData(form),
            credentials: 'same-origin',
            headers: {
              'Accept': 'application/json',
              'X-Requested-With': 'fetch'
            }
          });
          const contentType = response.headers.get('content-type') || '';
          const payload = contentType.includes('application/json')
            ? await response.json()
            : { ok: response.ok, message: await response.text() };
          if (!response.ok || payload.ok === false) {
            throw new Error(payload.message || '调整次数失败，请稍后再试。');
          }

          const row = userTable?.querySelector(`[data-user-id="${payload.user.id}"]`);
          if (row) {
            const quotaText = row.querySelector('[data-quota-text]');
            const remainingText = row.querySelector('[data-remaining-text]');
            const quotaFill = row.querySelector('[data-quota-fill]');
            if (quotaText) quotaText.textContent = `${payload.user.submissions_used} / ${payload.user.submission_quota}`;
            if (remainingText) remainingText.textContent = String(payload.user.remaining);
            if (quotaFill) {
              const total = Number(payload.user.submission_quota || 0);
              const used = Number(payload.user.submissions_used || 0);
              const percent = total > 0 ? Math.min(100, Math.max(0, (used / total) * 100)) : 0;
              quotaFill.style.setProperty('--quota-percent', `${percent}%`);
            }
          }
          const customAmount = form.querySelector('input[name="amount"][type="number"]');
          if (customAmount) customAmount.value = '';
          showAdminNotice(payload.message || '次数已调整。');
        } catch (error) {
          showAdminNotice(error.message || '调整次数失败，请稍后再试。', true);
        } finally {
          if (button) {
            button.disabled = false;
            button.textContent = originalText;
          }
        }
      });
    });

    document.querySelectorAll('[data-copy-text]').forEach(button => {
      button.addEventListener('click', async () => {
        try {
          await copyText(button.dataset.copyText || '');
          showAdminNotice('邀请链接已复制。');
        } catch (_error) {
          showAdminNotice('复制失败，请手动复制邀请码。', true);
        }
      });
    });

    document.querySelectorAll('form[data-confirm]').forEach(form => {
      form.addEventListener('submit', event => {
        const message = form.dataset.confirm || '确认继续吗？';
        if (!window.confirm(message)) event.preventDefault();
      });
    });

    document.querySelectorAll('form[data-confirm-email]').forEach(form => {
      form.addEventListener('submit', event => {
        const email = form.dataset.confirmEmail || '';
        const typed = window.prompt(`这是高风险操作。请输入目标邮箱 ${email} 确认注销：`, '');
        if (typed === null) {
          event.preventDefault();
          return;
        }
        if (typed.trim().toLowerCase() !== email.toLowerCase()) {
          event.preventDefault();
          showAdminNotice('输入的邮箱与目标账号不一致，已取消注销操作。', true);
          return;
        }
        const field = form.querySelector('input[name="confirm_email"]');
        if (field) field.value = typed.trim();
      });
    });
  </script>
</body>
</html>
"""


ADMIN_LOG_PAGE = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>操作日志 - 管理后台</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #16202a;
      --muted: #607181;
      --paper: #f3efe6;
      --surface: rgba(255, 255, 255, .86);
      --surface-strong: #ffffff;
      --line: #d7d9d2;
      --accent: #176f58;
      --shadow: rgba(22, 32, 42, .12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100svh;
      color: var(--ink);
      background:
        radial-gradient(circle at 8% 12%, color-mix(in srgb, var(--accent) 18%, transparent), transparent 28rem),
        var(--paper);
      font-family: "PingFang SC", "Noto Sans SC", sans-serif;
    }
    main {
      width: min(1120px, calc(100% - 32px));
      margin: 0 auto;
      padding: 34px 0 48px;
    }
    .topbar, .section-head, .pager {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
    }
    h1, h2 {
      margin: 0;
    }
    .copy {
      margin: 12px 0 0;
      color: var(--muted);
      line-height: 1.8;
    }
    a {
      color: var(--accent);
      font-weight: 800;
      text-decoration: none;
    }
    .panel {
      margin-top: 22px;
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      overflow: hidden;
      backdrop-filter: blur(16px);
    }
    .section-head {
      padding: 18px;
      border-bottom: 1px solid var(--line);
    }
    .log-list {
      display: grid;
    }
    .log-item {
      display: grid;
      grid-template-columns: minmax(0, 1.2fr) minmax(0, .8fr) auto;
      gap: 16px;
      padding: 18px;
      border-bottom: 1px solid var(--line);
    }
    .log-item:last-child {
      border-bottom: 0;
    }
    .log-item strong {
      display: block;
      margin-bottom: 4px;
    }
    .meta {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.8;
    }
    .details {
      margin-top: 8px;
      padding: 10px 12px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      color: var(--muted);
      font-family: "SFMono-Regular", "Menlo", monospace;
      font-size: 12px;
      white-space: pre-wrap;
      word-break: break-word;
    }
    .empty-state {
      padding: 40px 18px;
      color: var(--muted);
      text-align: center;
    }
    .pager {
      padding: 18px;
      border-top: 1px solid var(--line);
    }
    .pager-links {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .pager-link {
      display: inline-flex;
      min-width: 40px;
      align-items: center;
      justify-content: center;
      padding: 10px 12px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      color: var(--ink);
      font-size: 13px;
      font-weight: 800;
    }
    .pager-link.active {
      background: var(--accent);
      border-color: var(--accent);
      color: #fff;
    }
    @media (max-width: 820px) {
      .topbar, .section-head, .pager, .log-item {
        align-items: flex-start;
        flex-direction: column;
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>
  <main>
    <div class="topbar">
      <div>
        <h1>操作日志</h1>
        <p class="copy">这里会记录管理员的关键动作，包括额度调整、权限变更、冻结、恢复和注销。</p>
      </div>
      <div>
        <a href="{{ url_for('admin', auth_token=auth_token) }}">返回管理后台</a>
      </div>
    </div>
    <section class="panel">
      <div class="section-head">
        <h2>最近记录</h2>
        <div class="meta">当前显示第 {{ pagination['start'] }} - {{ pagination['end'] }} 条，共 {{ pagination['total'] }} 条</div>
      </div>
      {% if logs %}
        <div class="log-list">
          {% for log in logs %}
            <div class="log-item">
              <div>
                <strong>{{ log['summary'] }}</strong>
                <div class="meta">操作者：{{ log['actor_email'] or '未知管理员' }}</div>
                {% if log['details_pretty'] %}
                  <div class="details">{{ log['details_pretty'] }}</div>
                {% endif %}
              </div>
              <div class="meta">
                目标账号：{{ log['target_email'] or '无' }}<br>
                动作类型：{{ log['action'] }}
              </div>
              <div class="meta">{{ log['created_at'] }}</div>
            </div>
          {% endfor %}
        </div>
      {% else %}
        <div class="empty-state">还没有后台操作日志。</div>
      {% endif %}
      <div class="pager">
        <div class="meta">第 {{ pagination['page'] }} / {{ pagination['pages'] }} 页</div>
        <div class="pager-links">
          {% if pagination['page'] > 1 %}
            <a class="pager-link" href="{{ url_for('admin_logs', auth_token=auth_token, page=pagination['page'] - 1) }}">上一页</a>
          {% endif %}
          {% for page_number in page_numbers %}
            <a class="pager-link {% if page_number == pagination['page'] %}active{% endif %}" href="{{ url_for('admin_logs', auth_token=auth_token, page=page_number) }}">{{ page_number }}</a>
          {% endfor %}
          {% if pagination['page'] < pagination['pages'] %}
            <a class="pager-link" href="{{ url_for('admin_logs', auth_token=auth_token, page=pagination['page'] + 1) }}">下一页</a>
          {% endif %}
        </div>
      </div>
    </section>
  </main>
</body>
</html>
"""


MY_REPORTS_PAGE = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>我的检测记录 - UPC本科论文格式检测工具</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #16202a;
      --muted: #607181;
      --paper: #f3efe6;
      --surface: rgba(255, 255, 255, .88);
      --surface-strong: #ffffff;
      --line: #d7d9d2;
      --accent: #176f58;
      --accent-soft: #d8ece5;
      --warn: #9d4131;
      --warn-soft: #f7e0db;
      --shadow: rgba(22, 32, 42, .12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100svh;
      color: var(--ink);
      background:
        radial-gradient(circle at 8% 12%, color-mix(in srgb, var(--accent) 18%, transparent), transparent 28rem),
        var(--paper);
      font-family: "PingFang SC", "Noto Sans SC", sans-serif;
    }
    main {
      width: min(1160px, calc(100% - 32px));
      margin: 0 auto;
      padding: 34px 0 48px;
    }
    .topbar, .summary, .empty-state, .record-item {
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      backdrop-filter: blur(16px);
    }
    .topbar {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 18px;
      padding: 22px;
      margin-bottom: 20px;
    }
    h1 {
      margin: 0;
      font-size: clamp(32px, 5vw, 56px);
      line-height: 1;
    }
    .copy {
      margin: 12px 0 0;
      color: var(--muted);
      line-height: 1.8;
    }
    .top-links {
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
    }
    a {
      color: var(--accent);
      font-weight: 800;
      text-decoration: none;
    }
    .link-chip {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 10px 14px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 16px;
      padding: 18px 22px;
      margin-bottom: 18px;
    }
    .toolbar, .pager {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 18px 22px;
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      backdrop-filter: blur(16px);
      margin-bottom: 18px;
    }
    .toolbar {
      flex-wrap: wrap;
    }
    .toolbar form {
      display: grid;
      grid-template-columns: minmax(240px, 1fr) minmax(160px, 220px) auto auto;
      gap: 12px;
      width: 100%;
    }
    .field, .select {
      width: 100%;
      padding: 12px 14px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      color: var(--ink);
      font: 14px/1.4 "PingFang SC", "Noto Sans SC", sans-serif;
      outline: 0;
    }
    .pager-links {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .pager-link {
      display: inline-flex;
      min-width: 40px;
      align-items: center;
      justify-content: center;
      padding: 10px 12px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
    }
    .pager-link.active {
      background: var(--accent);
      border-color: var(--accent);
      color: #fff;
    }
    .summary strong {
      display: block;
      margin-top: 8px;
      font-size: 28px;
      line-height: 1;
    }
    .summary span {
      color: var(--muted);
      font-size: 13px;
      font-weight: 800;
      letter-spacing: .08em;
      text-transform: uppercase;
    }
    .records {
      display: grid;
      gap: 14px;
    }
    .record-item {
      padding: 18px 20px;
    }
    .record-head {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 12px;
    }
    .record-title {
      margin: 0;
      font-size: 18px;
      line-height: 1.4;
    }
    .meta {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.8;
    }
    .ua {
      max-width: 280px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .status {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 84px;
      padding: 8px 10px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: .04em;
    }
    .status-success {
      border-color: color-mix(in srgb, var(--accent) 30%, var(--line));
      background: color-mix(in srgb, var(--accent-soft) 42%, transparent);
      color: var(--accent);
    }
    .status-failed,
    .status-storage {
      border-color: color-mix(in srgb, var(--warn) 34%, var(--line));
      background: color-mix(in srgb, var(--warn-soft) 72%, transparent);
      color: var(--warn);
    }
    .actions {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-top: 12px;
    }
    .action-link {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 10px 14px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
    }
    .error-note {
      margin-top: 10px;
      padding: 10px 12px;
      border: 1px solid color-mix(in srgb, var(--warn) 35%, var(--line));
      background: color-mix(in srgb, var(--warn-soft) 72%, transparent);
      color: var(--warn);
      font-size: 13px;
      line-height: 1.7;
    }
    .empty-state {
      padding: 42px 24px;
      color: var(--muted);
      text-align: center;
      line-height: 1.9;
    }
    @media (max-width: 820px) {
      .topbar, .record-head, .toolbar, .pager {
        flex-direction: column;
      }
      .toolbar form {
        grid-template-columns: 1fr;
      }
      .summary {
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>
  <main>
    <section class="topbar">
      <div>
        <h1>我的检测记录</h1>
        <p class="copy">你可以在这里查看历史检测结果，并重新下载已经生成成功的 HTML 报告。</p>
      </div>
      <div class="top-links">
        <a class="link-chip" href="{{ url_for('index', auth_token=auth_token) }}">返回检测页</a>
        {% if is_admin %}<a class="link-chip" href="{{ url_for('admin_reports', auth_token=auth_token) }}">后台检测记录</a>{% endif %}
        <a class="link-chip" href="{{ url_for('logout') }}">退出登录</a>
      </div>
    </section>
    <section class="summary">
      <div><span>总记录数</span><strong>{{ stats['total'] }}</strong></div>
      <div><span>成功报告</span><strong>{{ stats['success'] }}</strong></div>
      <div><span>失败/异常</span><strong>{{ stats['failed'] }}</strong></div>
    </section>
    <section class="toolbar">
      <form method="get" action="{{ url_for('my_reports') }}">
        <input name="auth_token" type="hidden" value="{{ auth_token }}">
        <input name="page" type="hidden" value="1">
        <input class="field" name="q" type="search" value="{{ table_state['q'] }}" placeholder="搜索原文件名、报告名或时间">
        <select class="select" name="status">
          <option value="all">全部状态</option>
          <option value="success" {% if table_state['status'] == 'success' %}selected{% endif %}>仅看成功</option>
          <option value="audit_failed" {% if table_state['status'] == 'audit_failed' %}selected{% endif %}>仅看检测失败</option>
          <option value="storage_failed" {% if table_state['status'] == 'storage_failed' %}selected{% endif %}>仅看存档失败</option>
        </select>
        <button class="link-chip" type="submit">应用筛选</button>
        <a class="link-chip" href="{{ url_for('my_reports', auth_token=auth_token) }}">重置</a>
      </form>
    </section>
    {% if reports %}
      <section class="records">
        {% for item in reports %}
          <article class="record-item">
            <div class="record-head">
              <div>
                <h2 class="record-title">{{ item['original_filename'] }}</h2>
                <div class="meta">生成时间：{{ item['created_at_display'] }}</div>
                <div class="meta">报告文件：{{ item['report_filename'] or '未生成' }}</div>
                {% if item['report_size_display'] %}
                  <div class="meta">报告大小：{{ item['report_size_display'] }}{% if item['report_sha_short'] %} · SHA256 {{ item['report_sha_short'] }}...{% endif %}</div>
                {% endif %}
              </div>
              <span class="status {% if item['status'] == 'success' %}status-success{% elif item['status'] == 'storage_failed' %}status-storage{% else %}status-failed{% endif %}">
                {{ report_status_labels.get(item['status'], item['status']) }}
              </span>
            </div>
            {% if item['error_message'] %}
              <div class="error-note">{{ item['error_message'] }}</div>
            {% endif %}
            <div class="actions">
              {% if item['status'] == 'success' and (item['report_storage_path'] or item['report_gcs_path']) %}
                <a class="action-link" href="{{ url_for('download_report', report_id=item['id'], auth_token=auth_token) }}">重新下载报告</a>
              {% endif %}
            </div>
          </article>
        {% endfor %}
      </section>
      <section class="pager">
        <div class="meta">当前显示第 {{ pagination['start'] }} - {{ pagination['end'] }} 条，共 {{ pagination['total'] }} 条</div>
        <div class="pager-links">
          {% if pagination['page'] > 1 %}
            <a class="pager-link" href="{{ build_reports_url(auth_token, table_state, page=pagination['page'] - 1) }}">上一页</a>
          {% endif %}
          {% for page_number in page_numbers %}
            <a class="pager-link {% if page_number == pagination['page'] %}active{% endif %}" href="{{ build_reports_url(auth_token, table_state, page=page_number) }}">{{ page_number }}</a>
          {% endfor %}
          {% if pagination['page'] < pagination['pages'] %}
            <a class="pager-link" href="{{ build_reports_url(auth_token, table_state, page=pagination['page'] + 1) }}">下一页</a>
          {% endif %}
        </div>
      </section>
    {% else %}
      <section class="empty-state">
        你还没有检测记录。上传一篇论文生成报告后，这里会显示历史记录。
      </section>
    {% endif %}
  </main>
</body>
</html>
"""


ADMIN_REPORTS_PAGE = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>检测记录 - 管理后台</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #16202a;
      --muted: #607181;
      --paper: #f3efe6;
      --surface: rgba(255, 255, 255, .88);
      --surface-strong: #ffffff;
      --line: #d7d9d2;
      --accent: #176f58;
      --accent-soft: #d8ece5;
      --warn: #9d4131;
      --warn-soft: #f7e0db;
      --shadow: rgba(22, 32, 42, .12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100svh;
      color: var(--ink);
      background:
        radial-gradient(circle at 8% 12%, color-mix(in srgb, var(--accent) 18%, transparent), transparent 28rem),
        var(--paper);
      font-family: "PingFang SC", "Noto Sans SC", sans-serif;
    }
    main {
      width: min(1240px, calc(100% - 32px));
      margin: 0 auto;
      padding: 34px 0 48px;
    }
    .topbar, .summary, .panel {
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      backdrop-filter: blur(16px);
    }
    .topbar {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 18px;
      padding: 22px;
      margin-bottom: 20px;
    }
    h1 {
      margin: 0;
      font-size: clamp(32px, 5vw, 56px);
      line-height: 1;
    }
    .copy {
      margin: 12px 0 0;
      color: var(--muted);
      line-height: 1.8;
    }
    .top-links {
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
    }
    a {
      color: var(--accent);
      font-weight: 800;
      text-decoration: none;
    }
    .link-chip {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 10px 14px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
      gap: 16px;
      padding: 18px 22px;
      margin-bottom: 18px;
    }
    .toolbar, .pager {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 18px 22px;
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      backdrop-filter: blur(16px);
      margin-bottom: 18px;
    }
    .toolbar {
      flex-wrap: wrap;
    }
    .toolbar form {
      display: grid;
      grid-template-columns: minmax(240px, 1fr) minmax(160px, 210px) minmax(180px, 240px) auto auto;
      gap: 12px;
      width: 100%;
    }
    .field, .select {
      width: 100%;
      padding: 12px 14px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      color: var(--ink);
      font: 14px/1.4 "PingFang SC", "Noto Sans SC", sans-serif;
      outline: 0;
    }
    .summary strong {
      display: block;
      margin-top: 8px;
      font-size: 28px;
      line-height: 1;
    }
    .summary span {
      color: var(--muted);
      font-size: 13px;
      font-weight: 800;
      letter-spacing: .08em;
      text-transform: uppercase;
    }
    .panel {
      overflow: hidden;
    }
    .college-board {
      display: grid;
      grid-template-columns: minmax(0, 1.15fr) minmax(280px, .85fr);
      gap: 18px;
      margin-bottom: 18px;
    }
    .board-card {
      border: 1px solid var(--line);
      background: var(--surface);
      box-shadow: 0 24px 72px var(--shadow);
      backdrop-filter: blur(16px);
      padding: 20px 22px;
    }
    .board-title {
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 16px;
    }
    .board-title h2 {
      margin: 0;
      font-size: 22px;
      line-height: 1.2;
    }
    .board-title .meta {
      text-align: right;
    }
    .college-bars {
      display: grid;
      gap: 12px;
    }
    .college-row {
      display: grid;
      grid-template-columns: minmax(150px, 220px) minmax(120px, 1fr) 74px;
      align-items: center;
      gap: 12px;
    }
    .college-name {
      color: var(--ink);
      font-weight: 800;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .bar-track {
      height: 14px;
      border: 1px solid color-mix(in srgb, var(--accent) 24%, var(--line));
      background: color-mix(in srgb, var(--accent-soft) 46%, transparent);
      overflow: hidden;
    }
    .bar-fill {
      height: 100%;
      min-width: 4px;
      background: linear-gradient(90deg, var(--accent), #d29b48);
    }
    .college-count {
      color: var(--muted);
      font-size: 13px;
      text-align: right;
      white-space: nowrap;
    }
    .insight-list {
      display: grid;
      gap: 12px;
      margin-top: 16px;
    }
    .insight-item {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 12px 14px;
      border: 1px solid var(--line);
      background: color-mix(in srgb, var(--surface-strong) 82%, transparent);
    }
    .pager-links {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .pager-link {
      display: inline-flex;
      min-width: 40px;
      align-items: center;
      justify-content: center;
      padding: 10px 12px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
    }
    .pager-link.active {
      background: var(--accent);
      border-color: var(--accent);
      color: #fff;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }
    th, td {
      padding: 16px 18px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: middle;
    }
    th {
      color: var(--muted);
      font-size: 12px;
      letter-spacing: .08em;
      text-transform: uppercase;
      background: color-mix(in srgb, var(--accent-soft) 28%, transparent);
    }
    .meta {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.8;
    }
    .status {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 84px;
      padding: 8px 10px;
      border: 1px solid var(--line);
      background: var(--surface-strong);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: .04em;
    }
    .status-success {
      border-color: color-mix(in srgb, var(--accent) 30%, var(--line));
      background: color-mix(in srgb, var(--accent-soft) 42%, transparent);
      color: var(--accent);
    }
    .status-failed,
    .status-storage {
      border-color: color-mix(in srgb, var(--warn) 34%, var(--line));
      background: color-mix(in srgb, var(--warn-soft) 72%, transparent);
      color: var(--warn);
    }
    .error-note {
      color: var(--warn);
      font-size: 13px;
      line-height: 1.7;
    }
    .empty-state {
      padding: 40px 22px;
      color: var(--muted);
      text-align: center;
      line-height: 1.9;
    }
    @media (max-width: 820px) {
      .topbar, .toolbar, .pager {
        flex-direction: column;
      }
      .toolbar form {
        grid-template-columns: 1fr;
      }
      .college-board {
        grid-template-columns: 1fr;
      }
      .college-row {
        grid-template-columns: 1fr;
        gap: 7px;
      }
      .college-count,
      .board-title .meta {
        text-align: left;
      }
      .summary {
        grid-template-columns: 1fr;
      }
      table, thead, tbody, tr, th, td { display: block; }
      thead { display: none; }
      tr { border-bottom: 1px solid var(--line); padding: 14px; }
      td { border: 0; padding: 8px 4px; }
      td::before {
        content: attr(data-label);
        display: block;
        color: var(--muted);
        font-size: 12px;
        margin-bottom: 3px;
      }
    }
  </style>
</head>
<body>
  <main>
    <section class="topbar">
      <div>
        <h1>检测记录</h1>
        <p class="copy">查看所有用户的检测结果、生成状态与重新下载记录入口。</p>
      </div>
      <div class="top-links">
        <a class="link-chip" href="{{ url_for('admin', auth_token=auth_token) }}">返回管理后台</a>
        <a class="link-chip" href="{{ url_for('admin_logs', auth_token=auth_token) }}">操作日志</a>
      </div>
    </section>
    <section class="summary">
      <div><span>总记录数</span><strong>{{ stats['total'] }}</strong></div>
      <div><span>成功报告</span><strong>{{ stats['success'] }}</strong></div>
      <div><span>检测失败</span><strong>{{ stats['audit_failed'] }}</strong></div>
      <div><span>存档失败</span><strong>{{ stats['storage_failed'] }}</strong></div>
      <div><span>识别出学院</span><strong>{{ college_stats['known'] }}</strong></div>
      <div><span>未识别学院</span><strong>{{ college_stats['unknown'] }}</strong></div>
    </section>
    <section class="college-board">
      <div class="board-card">
        <div class="board-title">
          <div>
            <h2>学院分布看板</h2>
            <div class="meta">按每一次检测记录归类，同一个账号的不同报告会分别统计。</div>
          </div>
          <div class="meta">Top：{{ college_stats['top']['college'] }} · {{ college_stats['top']['count'] }} 次</div>
        </div>
        {% if college_stats['rows'] %}
          <div class="college-bars">
            {% for row in college_stats['rows'][:12] %}
              <div class="college-row">
                <div class="college-name" title="{{ row['college'] }}">{{ row['college'] }}</div>
                <div class="bar-track" aria-label="{{ row['college'] }} {{ row['count'] }} 次">
                  <div class="bar-fill" style="width: {{ row['bar_percent'] }}%;"></div>
                </div>
                <div class="college-count">{{ row['count'] }} 次 · {{ row['percent'] }}%</div>
              </div>
            {% endfor %}
          </div>
        {% else %}
          <div class="empty-state">还没有可统计的检测记录。</div>
        {% endif %}
      </div>
      <div class="board-card">
        <div class="board-title">
          <h2>识别说明</h2>
          <div class="meta">封面优先</div>
        </div>
        <div class="insight-list">
          <div class="insight-item"><span>统计口径</span><strong>按报告</strong></div>
          <div class="insight-item"><span>优先来源</span><strong>封面/表格</strong></div>
          <div class="insight-item"><span>筛选结果</span><strong>{{ pagination['total'] }} 条</strong></div>
          <div class="insight-item"><span>未识别处理</span><strong>单独归类</strong></div>
        </div>
      </div>
    </section>
    <section class="toolbar">
      <form method="get" action="{{ url_for('admin_reports') }}">
        <input name="auth_token" type="hidden" value="{{ auth_token }}">
        <input name="page" type="hidden" value="1">
        <input class="field" name="q" type="search" value="{{ table_state['q'] }}" placeholder="搜索用户邮箱、原文件名、报告名、学院或时间">
        <select class="select" name="status">
          <option value="all">全部状态</option>
          <option value="success" {% if table_state['status'] == 'success' %}selected{% endif %}>仅看成功</option>
          <option value="audit_failed" {% if table_state['status'] == 'audit_failed' %}selected{% endif %}>仅看检测失败</option>
          <option value="storage_failed" {% if table_state['status'] == 'storage_failed' %}selected{% endif %}>仅看存档失败</option>
        </select>
        <select class="select" name="college">
          <option value="all">全部学院</option>
          {% for college in college_stats['options'] %}
            <option value="{{ college }}" {% if table_state['college'] == college %}selected{% endif %}>{{ college }}</option>
          {% endfor %}
        </select>
        <button class="link-chip" type="submit">应用筛选</button>
        <a class="link-chip" href="{{ url_for('admin_reports', auth_token=auth_token) }}">重置</a>
      </form>
    </section>
    <section class="panel">
      {% if reports %}
        <table>
          <thead>
            <tr>
              <th>用户</th>
              <th>原文件名</th>
              <th>学院</th>
              <th>状态</th>
              <th>生成时间</th>
              <th>来源</th>
              <th>操作</th>
            </tr>
          </thead>
          <tbody>
            {% for item in reports %}
              <tr>
                <td data-label="用户">
                  <div>{{ item['user_email'] }}</div>
                  <div class="meta">{{ item['user_id'] }}</div>
                </td>
                <td data-label="原文件名">
                  <div>{{ item['original_filename'] }}</div>
                  <div class="meta">{{ item['report_filename'] or '未生成报告名' }}</div>
                  {% if item['original_size_display'] %}
                    <div class="meta">原文 {{ item['original_size_display'] }}{% if item['original_sha_short'] %} · SHA256 {{ item['original_sha_short'] }}...{% endif %}</div>
                  {% endif %}
                  <div class="meta">
                    原文：{% if item['original_storage_path'] %}Supabase{% else %}未存 Supabase{% endif %}
                    {% if item['original_gcs_path'] %} / GCS{% endif %}
                    {% if item['original_drive_file_id'] %} / Drive{% endif %}
                  </div>
                  <div class="meta">
                    报告：{% if item['report_storage_path'] %}Supabase{% else %}未存 Supabase{% endif %}
                    {% if item['report_gcs_path'] %} / GCS{% endif %}
                  </div>
                </td>
                <td data-label="学院">
                  <div>{{ item['college_name'] }}</div>
                  {% if item['college_source'] %}
                    <div class="meta">{{ item['college_source'] }}</div>
                  {% endif %}
                  {% if item['college_raw_text'] %}
                    <div class="meta" title="{{ item['college_raw_text'] }}">{{ item['college_raw_text'] }}</div>
                  {% endif %}
                </td>
                <td data-label="状态">
                  <span class="status {% if item['status'] == 'success' %}status-success{% elif item['status'] == 'storage_failed' %}status-storage{% else %}status-failed{% endif %}">
                    {{ report_status_labels.get(item['status'], item['status']) }}
                  </span>
                  {% if item['error_message'] %}
                    <div class="error-note">{{ item['error_message'] }}</div>
                  {% endif %}
                </td>
                <td data-label="生成时间" class="meta">{{ item['created_at_display'] }}</td>
                <td data-label="来源" class="meta">
                  <div>{{ item['client_ip'] or '旧记录未记录 IP' }}</div>
                  {% if item['user_agent'] %}
                    <div class="ua" title="{{ item['user_agent'] }}">{{ item['user_agent'] }}</div>
                  {% else %}
                    <div>旧记录未记录浏览器</div>
                  {% endif %}
                </td>
                <td data-label="操作">
                  {% if item['status'] == 'success' and (item['report_storage_path'] or item['report_gcs_path']) %}
                    <a class="link-chip" href="{{ url_for('download_report', report_id=item['id'], auth_token=auth_token) }}">下载报告</a>
                  {% endif %}
                  {% if item['original_storage_path'] or item['original_gcs_path'] or item['original_drive_file_id'] %}
                    <a class="link-chip" href="{{ url_for('download_original', report_id=item['id'], auth_token=auth_token) }}">下载原文</a>
                  {% endif %}
                  {% if not ((item['status'] == 'success' and (item['report_storage_path'] or item['report_gcs_path'])) or item['original_storage_path'] or item['original_gcs_path'] or item['original_drive_file_id']) %}
                    <span class="meta">暂无可下载文件</span>
                  {% endif %}
                </td>
              </tr>
            {% endfor %}
          </tbody>
        </table>
        <section class="pager">
          <div class="meta">当前显示第 {{ pagination['start'] }} - {{ pagination['end'] }} 条，共 {{ pagination['total'] }} 条</div>
          <div class="pager-links">
            {% if pagination['page'] > 1 %}
              <a class="pager-link" href="{{ build_reports_url(auth_token, table_state, page=pagination['page'] - 1) }}">上一页</a>
            {% endif %}
            {% for page_number in page_numbers %}
              <a class="pager-link {% if page_number == pagination['page'] %}active{% endif %}" href="{{ build_reports_url(auth_token, table_state, page=page_number) }}">{{ page_number }}</a>
            {% endfor %}
            {% if pagination['page'] < pagination['pages'] %}
              <a class="pager-link" href="{{ build_reports_url(auth_token, table_state, page=pagination['page'] + 1) }}">下一页</a>
            {% endif %}
          </div>
        </section>
      {% else %}
        <div class="empty-state">还没有任何检测记录。</div>
      {% endif %}
    </section>
  </main>
</body>
</html>
"""


@app.get("/")
def index() -> str:
    record_event("page_view", current_user(), {"auth_mode": request.args.get("auth", "")})
    return render_home()


@app.post("/register")
def register():
    limited = rate_limit("register")
    if limited:
        return limited
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        return render_home(auth_error="服务还没有配置 Supabase 数据库。", auth_mode="register"), 503
    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")
    confirm_password = request.form.get("confirm_password", "")
    registration_code = normalize_registration_code(request.form.get("registration_code", ""))
    invite_code = normalize_invite_code(request.form.get("invite_code", ""))
    email_code = request.form.get("email_code", "").strip()
    captcha_answer = request.form.get("captcha_answer", "").strip()
    captcha_left = request.form.get("captcha_left", "")
    captcha_right = request.form.get("captcha_right", "")
    auth_values = registration_values(email, password, confirm_password, registration_code, invite_code, email_code)
    if not is_valid_registration_email(email):
        refresh_captcha()
        return render_home(auth_error="新注册仅支持 QQ 邮箱，请使用类似 123456@qq.com 的邮箱地址。", auth_mode="register", auth_values=auth_values), 400
    if not email_verification_enabled():
        refresh_captcha()
        return render_home(auth_error="邮箱验证码服务尚未配置，暂时无法注册新账号。", auth_mode="register", auth_values=auth_values), 503
    if not is_valid_email_code(email, email_code):
        refresh_captcha()
        return render_home(auth_error="邮箱验证码不正确，或已经过期，请重新发送后再试。", auth_mode="register", auth_values=auth_values), 400
    if len(password) < 6:
        refresh_captcha()
        return render_home(auth_error="密码至少需要 6 位。", auth_mode="register", auth_values=auth_values), 400
    if password != confirm_password:
        refresh_captcha()
        return render_home(auth_error="两次输入的密码不一致。", auth_mode="register", auth_values=auth_values), 400
    if not is_valid_captcha(captcha_answer, captcha_left, captcha_right):
        refresh_captcha()
        return render_home(auth_error="数字验证不正确，请重新计算。", auth_mode="register", auth_values=auth_values), 400
    access_code_record = find_registration_code(registration_code)
    if not registration_code_is_available(access_code_record):
        refresh_captcha()
        return render_home(auth_error="QQ群注册码不存在、已停用或使用次数已满。请加入官方 QQ 群 537124215，从群公告获取最新注册码。", auth_mode="register", auth_values=auth_values), 400
    if find_user_by_email(email) is not None:
        refresh_captcha()
        return render_home(auth_error="这个邮箱已经注册，请直接登录。", auth_mode="register", auth_values=auth_values), 400

    inviter = find_user_by_invite_code(invite_code) if invite_code else None
    if invite_code and inviter is None:
        refresh_captcha()
        return render_home(auth_error="邀请码不存在，请检查后再注册。", auth_mode="register", auth_values=auth_values), 400
    if inviter and inviter.get("email", "").lower() == email:
        refresh_captcha()
        return render_home(auth_error="不能使用自己的邀请码注册。", auth_mode="register", auth_values=auth_values), 400

    try:
        consumed_code = consume_registration_code(registration_code)
    except ValueError as exc:
        refresh_captcha()
        return render_home(auth_error=str(exc), auth_mode="register", auth_values=auth_values), 400
    try:
        user = create_user(email, password, invited_by=inviter["id"] if inviter else None)
    except APIError:
        refresh_captcha()
        return render_home(auth_error="这个邮箱已经注册，请直接登录。", auth_mode="register", auth_values=auth_values), 400
    try:
        record_admin_log(
            None,
            "registration_code_use",
            user,
            f"{email} 使用 QQ 群注册码 {consumed_code['code']} 完成注册",
            {"code": consumed_code["code"], "used_count": consumed_code.get("used_count"), "max_uses": consumed_code.get("max_uses")},
        )
    except Exception:
        app.logger.warning("Failed to consume registration code %s for user %s", registration_code, user["id"], exc_info=True)
    if inviter:
        try:
            award_invite_bonus(inviter["id"])
        except Exception:
            app.logger.warning("Failed to award invite bonus for inviter %s", inviter["id"], exc_info=True)
    clear_email_code()
    session["user_id"] = user["id"]
    record_event("register_success", user, {"registration_code": consumed_code.get("code", ""), "invited": bool(inviter)})
    return redirect(url_for("index", auth_token=generate_auth_token(user["id"])))


@app.post("/register/email-code")
def send_register_email_code():
    wants_json = request.headers.get("X-Requested-With") == "fetch" or "application/json" in request.headers.get("Accept", "")
    limited = rate_limit("email_code")
    if limited:
        message = limited.get_data(as_text=True)
        if wants_json:
            return jsonify({"ok": False, "message": message}), limited.status_code
        return Response(message, status=limited.status_code, mimetype="text/plain; charset=utf-8")
    if not email_verification_enabled():
        message = "邮箱验证码服务尚未配置。"
        if wants_json:
            return jsonify({"ok": False, "message": message}), 503
        return Response(message, status=503, mimetype="text/plain; charset=utf-8")
    resend_seconds = email_code_remaining_seconds()
    if resend_seconds > 0:
        message = f"请在 {resend_seconds} 秒后再重新发送。"
        if wants_json:
            return jsonify({"ok": False, "message": message}), 429
        return Response(message, status=429, mimetype="text/plain; charset=utf-8")

    email = request.form.get("email", "").strip().lower()
    if not is_valid_registration_email(email):
        message = "新注册仅支持 QQ 邮箱，请使用类似 123456@qq.com 的邮箱地址。"
        if wants_json:
            return jsonify({"ok": False, "message": message}), 400
        return Response(message, status=400, mimetype="text/plain; charset=utf-8")
    if find_user_by_email(email) is not None:
        message = "这个邮箱已经注册，请直接登录。"
        if wants_json:
            return jsonify({"ok": False, "message": message}), 400
        return Response(message, status=400, mimetype="text/plain; charset=utf-8")

    code = generate_email_code()
    try:
        send_registration_email_code(email, code)
    except Exception:
        app.logger.exception("Failed to send registration email code to %s", email)
        message = "验证码发送失败，请稍后再试。"
        if wants_json:
            return jsonify({"ok": False, "message": message}), 500
        return Response(message, status=500, mimetype="text/plain; charset=utf-8")

    store_email_code(email, code)
    message = "验证码已发送，请到邮箱中查看。若收不到验证码，请查看垃圾邮箱或广告邮件。"
    if wants_json:
        return jsonify({"ok": True, "message": message, "resend_seconds": EMAIL_CODE_RESEND_SECONDS})
    return Response(message, mimetype="text/plain; charset=utf-8")


@app.post("/login")
def login():
    limited = rate_limit("login")
    if limited:
        return limited
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        return render_home(auth_error="服务还没有配置 Supabase 数据库。"), 503
    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")
    auth_values = login_values(email, password)
    user = find_user_by_email(email)
    if user is None or not check_password_hash(user["password_hash"], password):
        return render_home(auth_error="邮箱或密码不正确。请确认这个邮箱已经注册，并且密码没有输错。", auth_mode="login", auth_values=auth_values), 400
    if not is_account_active(user):
        return render_home(auth_error=account_block_message(user), auth_mode="login", auth_values=auth_values), 403
    try:
        update_user_login_trace(user["id"])
    except Exception:
        app.logger.warning("Failed to update login trace for user %s", user["id"], exc_info=True)
    session["user_id"] = user["id"]
    record_event("login_success", user)
    return redirect(url_for("index", auth_token=generate_auth_token(user["id"])))


@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.get("/admin")
def admin():
    user = current_user()
    if not is_admin(user):
        return redirect(url_for("index"))
    token = request.values.get("auth_token") or generate_auth_token(user["id"])
    table_state = admin_table_state(token)
    all_users = attach_invite_stats([enrich_admin_user(item) for item in list_users()])
    stats = summarize_admin_stats(all_users)
    all_reports = list_reports_for_admin()
    traffic_stats = summarize_traffic_stats(all_users, all_reports, list_events_for_admin())
    filtered_users = apply_admin_user_filters(all_users, table_state)
    sorted_users = sort_admin_users(filtered_users, table_state["sort"])
    users, pagination = paginate_items(sorted_users, table_state["page"], table_state["per_page"])
    page_numbers = build_page_numbers(pagination["page"], pagination["pages"])
    current_url = build_admin_url(token, {**table_state, "page": pagination["page"]})
    recent_logs = list_admin_logs()[:6]
    registration_codes = list_registration_codes()
    return render_template_string(
        ADMIN_PAGE,
        admin_user=user,
        users=users,
        stats=stats,
        traffic_stats=traffic_stats,
        recent_logs=recent_logs,
        registration_codes=registration_codes,
        table_state={**table_state, "page": pagination["page"]},
        pagination=pagination,
        page_numbers=page_numbers,
        current_url=current_url,
        reset_url=url_for("admin", auth_token=token),
        sort_options=[
            {"value": "created_desc", "label": "按注册时间倒序"},
            {"value": "created_asc", "label": "按注册时间正序"},
            {"value": "remaining_desc", "label": "按剩余次数从高到低"},
            {"value": "remaining_asc", "label": "按剩余次数从低到高"},
            {"value": "quota_desc", "label": "按总额度从高到低"},
            {"value": "quota_asc", "label": "按总额度从低到高"},
            {"value": "used_desc", "label": "按已用次数从高到低"},
            {"value": "used_asc", "label": "按已用次数从低到高"},
            {"value": "email_asc", "label": "按邮箱 A-Z"},
            {"value": "email_desc", "label": "按邮箱 Z-A"},
        ],
        per_page_options=sorted(ADMIN_PER_PAGE_OPTIONS),
        build_admin_url=build_admin_url,
        can_manage_admins=is_super_admin(user),
        status_labels={
            ACCOUNT_STATUS_ACTIVE: account_status_label(ACCOUNT_STATUS_ACTIVE),
            ACCOUNT_STATUS_FROZEN: account_status_label(ACCOUNT_STATUS_FROZEN),
            ACCOUNT_STATUS_DISABLED: account_status_label(ACCOUNT_STATUS_DISABLED),
        },
        auth_token=token,
        message=request.args.get("message", ""),
    )


@app.post("/admin/quota")
def admin_add_quota():
    wants_json = request.headers.get("X-Requested-With") == "fetch" or "application/json" in request.headers.get("Accept", "")
    limited = rate_limit("admin")
    if limited:
        if wants_json:
            return jsonify({"ok": False, "message": limited.get_data(as_text=True)}), limited.status_code
        return limited
    user = current_user()
    if not is_admin(user):
        if wants_json:
            return jsonify({"ok": False, "message": "没有权限。"}), 403
        return Response("没有权限。", status=403, mimetype="text/plain; charset=utf-8")
    token = request.form.get("auth_token") or generate_auth_token(user["id"])
    target_user_id = request.form.get("user_id", "")
    try:
        amount = int(request.form.get("amount", "0"))
    except ValueError:
        amount = 0
    if amount <= 0:
        if wants_json:
            return jsonify({"ok": False, "message": "增加次数必须大于 0。"}), 400
        return redirect(url_for("admin", auth_token=token, message="增加次数必须大于 0。"))
    try:
        add_user_quota(target_user_id, amount)
        target_user = find_user_by_id(target_user_id)
    except ValueError as exc:
        if wants_json:
            return jsonify({"ok": False, "message": str(exc)}), 400
        return redirect_with_message(admin_redirect_url(token), str(exc))
    message = f"已增加 {amount} 次额度。"
    record_admin_log(
        user,
        "quota_add",
        target_user,
        f"给 {target_user['email']} 增加了 {amount} 次额度",
        {"amount": amount, "submission_quota": user_quota(target_user), "remaining": remaining_submissions(target_user)},
    )
    if wants_json:
        return jsonify({"ok": True, "message": message, "user": admin_user_quota_payload(target_user)})
    return redirect_with_message(admin_redirect_url(token), message)


@app.post("/admin/quota/reduce")
def admin_reduce_quota():
    wants_json = request.headers.get("X-Requested-With") == "fetch" or "application/json" in request.headers.get("Accept", "")
    limited = rate_limit("admin")
    if limited:
        if wants_json:
            return jsonify({"ok": False, "message": limited.get_data(as_text=True)}), limited.status_code
        return limited
    user = current_user()
    if not is_admin(user):
        if wants_json:
            return jsonify({"ok": False, "message": "没有权限。"}), 403
        return Response("没有权限。", status=403, mimetype="text/plain; charset=utf-8")
    token = request.form.get("auth_token") or generate_auth_token(user["id"])
    target_user_id = request.form.get("user_id", "")
    try:
        amount = int(request.form.get("amount", "0"))
    except ValueError:
        amount = 0
    if amount <= 0:
        if wants_json:
            return jsonify({"ok": False, "message": "减少次数必须大于 0。"}), 400
        return redirect(url_for("admin", auth_token=token, message="减少次数必须大于 0。"))
    try:
        reduce_user_quota(target_user_id, amount)
        target_user = find_user_by_id(target_user_id)
    except ValueError as exc:
        if wants_json:
            return jsonify({"ok": False, "message": str(exc)}), 400
        return redirect_with_message(admin_redirect_url(token), str(exc))
    message = f"已减少 {amount} 次额度。"
    record_admin_log(
        user,
        "quota_reduce",
        target_user,
        f"给 {target_user['email']} 减少了 {amount} 次额度",
        {"amount": amount, "submission_quota": user_quota(target_user), "remaining": remaining_submissions(target_user)},
    )
    if wants_json:
        return jsonify({"ok": True, "message": message, "user": admin_user_quota_payload(target_user)})
    return redirect_with_message(admin_redirect_url(token), message)


@app.post("/admin/status")
def admin_update_status():
    limited = rate_limit("admin")
    if limited:
        return limited
    user = current_user()
    if not is_admin(user):
        return Response("没有权限。", status=403, mimetype="text/plain; charset=utf-8")
    token = request.form.get("auth_token") or generate_auth_token(user["id"])
    target_user_id = request.form.get("user_id", "")
    status = request.form.get("status", "").strip().lower()
    confirm_email = request.form.get("confirm_email", "").strip().lower()
    target_user = find_user_by_id(target_user_id)
    if target_user is None:
        return redirect_with_message(admin_redirect_url(token), "用户不存在。")
    target_is_privileged = is_super_admin(target_user) or is_admin(target_user)
    if target_is_privileged and not is_super_admin(user):
        return redirect_with_message(admin_redirect_url(token), "普通管理员不能修改其他管理员账号状态。")
    if is_super_admin(target_user) and status != ACCOUNT_STATUS_ACTIVE:
        return redirect_with_message(admin_redirect_url(token), "最高管理员账号不能被冻结或注销。")
    if target_user_id == user.get("id") and status == ACCOUNT_STATUS_DISABLED:
        return redirect_with_message(admin_redirect_url(token), "不能注销当前管理员账号。")
    if status == ACCOUNT_STATUS_DISABLED and confirm_email != target_user.get("email", "").strip().lower():
        return redirect_with_message(admin_redirect_url(token), "注销确认失败，请输入目标邮箱后再试。")
    try:
        update_user_status(target_user_id, status)
    except ValueError as exc:
        return redirect_with_message(admin_redirect_url(token), str(exc))
    refreshed_target = find_user_by_id(target_user_id)
    record_admin_log(
        user,
        "status_update",
        refreshed_target,
        f"将 {target_user['email']} 的账号状态更新为 {account_status_label(status)}",
        {"status": status},
    )
    return redirect_with_message(admin_redirect_url(token), f"账号状态已更新为：{account_status_label(status)}。")


@app.post("/admin/role")
def admin_update_role():
    limited = rate_limit("admin")
    if limited:
        return limited
    user = current_user()
    if not is_super_admin(user):
        return Response("只有最高管理员可以设置管理员权限。", status=403, mimetype="text/plain; charset=utf-8")
    token = request.form.get("auth_token") or generate_auth_token(user["id"])
    target_user_id = request.form.get("user_id", "")
    admin_enabled = request.form.get("is_admin", "") == "1"
    target_user = find_user_by_id(target_user_id)
    if target_user is None:
        return redirect_with_message(admin_redirect_url(token), "用户不存在。")
    if is_super_admin(target_user):
        return redirect_with_message(admin_redirect_url(token), "最高管理员权限不能在后台取消。")
    try:
        update_user_admin(target_user_id, admin_enabled)
    except ValueError as exc:
        return redirect_with_message(admin_redirect_url(token), str(exc))
    action = "设为管理员" if admin_enabled else "取消管理员"
    refreshed_target = find_user_by_id(target_user_id)
    record_admin_log(
        user,
        "role_update",
        refreshed_target,
        f"已将 {target_user['email']} {action}",
        {"is_admin": admin_enabled},
    )
    return redirect_with_message(admin_redirect_url(token), f"已将 {target_user['email']} {action}。")


@app.post("/admin/registration-codes")
def admin_create_registration_code():
    limited = rate_limit("admin")
    if limited:
        return limited
    user = current_user()
    if not is_admin(user):
        return Response("没有权限。", status=403, mimetype="text/plain; charset=utf-8")
    token = request.form.get("auth_token") or generate_auth_token(user["id"])
    try:
        max_uses = int(request.form.get("max_uses", "20"))
    except ValueError:
        max_uses = 20
    note = request.form.get("note", "").strip()
    try:
        code = create_registration_code(user, max_uses=max_uses, note=note)
    except Exception as exc:
        app.logger.warning("Failed to create registration code", exc_info=True)
        return redirect_with_message(admin_redirect_url(token), f"生成注册码失败：{exc}")
    record_admin_log(
        user,
        "registration_code_create",
        None,
        f"生成 QQ 群注册码 {code['code']}，可用 {code['max_uses']} 次",
        {"code": code["code"], "max_uses": code["max_uses"], "note": code.get("note", "")},
    )
    return redirect_with_message(admin_redirect_url(token), f"已生成 QQ 群注册码：{code['code']}。请复制到 QQ 群 537124215 的群公告中。")


@app.post("/admin/registration-codes/status")
def admin_update_registration_code_status():
    limited = rate_limit("admin")
    if limited:
        return limited
    user = current_user()
    if not is_admin(user):
        return Response("没有权限。", status=403, mimetype="text/plain; charset=utf-8")
    token = request.form.get("auth_token") or generate_auth_token(user["id"])
    code_id = request.form.get("code_id", "").strip()
    is_active = request.form.get("is_active", "") == "1"
    try:
        code = update_registration_code_status(code_id, is_active)
    except ValueError as exc:
        return redirect_with_message(admin_redirect_url(token), str(exc))
    action = "启用" if is_active else "停用"
    record_admin_log(
        user,
        "registration_code_status",
        None,
        f"{action} QQ 群注册码 {code['code']}",
        {"code": code["code"], "is_active": is_active},
    )
    return redirect_with_message(admin_redirect_url(token), f"已{action} QQ 群注册码：{code['code']}。")


@app.get("/admin/logs")
def admin_logs():
    user = current_user()
    if not is_admin(user):
        return redirect(url_for("index"))
    token = request.values.get("auth_token") or generate_auth_token(user["id"])
    page = parse_positive_int(request.args.get("page"), 1)
    logs = list_admin_logs()
    for item in logs:
        details = item.get("details")
        if details in (None, "", {}):
            item["details_pretty"] = ""
        else:
            try:
                item["details_pretty"] = json.dumps(details, ensure_ascii=False, indent=2)
            except TypeError:
                item["details_pretty"] = str(details)
    page_logs, pagination = paginate_items(logs, page, 20)
    page_numbers = build_page_numbers(pagination["page"], pagination["pages"])
    return render_template_string(
        ADMIN_LOG_PAGE,
        auth_token=token,
        logs=page_logs,
        pagination=pagination,
        page_numbers=page_numbers,
    )


@app.get("/reports")
def my_reports():
    user = current_user()
    if user is None:
        return redirect(url_for("index"))
    token = request.values.get("auth_token") or generate_auth_token(user["id"])
    table_state = report_table_state()
    table_state["college"] = "all"
    reports = [enrich_report_item(item) for item in list_reports_for_user(user["id"])]
    filtered_reports = apply_report_filters(reports, table_state)
    page_reports, pagination = paginate_items(filtered_reports, table_state["page"], 10)
    page_numbers = build_page_numbers(pagination["page"], pagination["pages"])
    stats = {
        "total": len(reports),
        "success": sum(1 for item in reports if item.get("status") == "success"),
        "failed": sum(1 for item in reports if item.get("status") != "success"),
    }
    return render_template_string(
        MY_REPORTS_PAGE,
        auth_token=token,
        user=user,
        reports=page_reports,
        stats=stats,
        is_admin=is_admin(user),
        table_state={**table_state, "page": pagination["page"]},
        pagination=pagination,
        page_numbers=page_numbers,
        build_reports_url=lambda auth, state, **overrides: build_reports_url(auth, state, "my_reports", **overrides),
        report_status_labels={
            "success": report_status_label("success"),
            "audit_failed": report_status_label("audit_failed"),
            "storage_failed": report_status_label("storage_failed"),
        },
    )


@app.get("/admin/reports")
def admin_reports():
    user = current_user()
    if not is_admin(user):
        return redirect(url_for("index"))
    token = request.values.get("auth_token") or generate_auth_token(user["id"])
    table_state = report_table_state()
    reports = [enrich_report_item(item) for item in list_reports_for_admin()]
    college_stats = summarize_report_colleges(reports)
    filtered_reports = apply_report_filters(reports, table_state)
    page_reports, pagination = paginate_items(filtered_reports, table_state["page"], 20)
    page_numbers = build_page_numbers(pagination["page"], pagination["pages"])
    stats = {
        "total": len(reports),
        "success": sum(1 for item in reports if item.get("status") == "success"),
        "audit_failed": sum(1 for item in reports if item.get("status") == "audit_failed"),
        "storage_failed": sum(1 for item in reports if item.get("status") == "storage_failed"),
    }
    return render_template_string(
        ADMIN_REPORTS_PAGE,
        auth_token=token,
        reports=page_reports,
        stats=stats,
        college_stats=college_stats,
        table_state={**table_state, "page": pagination["page"]},
        pagination=pagination,
        page_numbers=page_numbers,
        build_reports_url=lambda auth, state, **overrides: build_reports_url(auth, state, "admin_reports", **overrides),
        report_status_labels={
            "success": report_status_label("success"),
            "audit_failed": report_status_label("audit_failed"),
            "storage_failed": report_status_label("storage_failed"),
        },
    )


@app.get("/reports/<report_id>/download")
def download_report(report_id: str):
    user = current_user()
    if user is None:
        return redirect(url_for("index"))
    report = find_report_by_id(report_id)
    if report is None:
        return Response("报告不存在。", status=404, mimetype="text/plain; charset=utf-8")
    if not is_admin(user) and report.get("user_id") != user.get("id"):
        return Response("没有权限下载这个报告。", status=403, mimetype="text/plain; charset=utf-8")
    if report.get("status") != "success" or not (report.get("report_storage_path") or report.get("report_gcs_path")):
        return Response("这个报告暂时不可下载。", status=404, mimetype="text/plain; charset=utf-8")
    try:
        report_bytes = download_report_from_storage(report)
    except Exception:
        app.logger.exception("Failed to download stored report %s", report_id)
        return Response("下载报告失败，请稍后再试。", status=500, mimetype="text/plain; charset=utf-8")
    return send_file(
        io.BytesIO(report_bytes),
        as_attachment=True,
        download_name=report.get("report_filename") or "thesis_format_audit_report.html",
        mimetype="text/html",
    )


@app.get("/reports/<report_id>/original")
def download_original(report_id: str):
    user = current_user()
    if user is None:
        return redirect(url_for("index"))
    report = find_report_by_id(report_id)
    if report is None:
        return Response("记录不存在。", status=404, mimetype="text/plain; charset=utf-8")
    if not is_admin(user):
        return Response("只有管理员可以下载原始论文文件。", status=403, mimetype="text/plain; charset=utf-8")
    if not (report.get("original_storage_path") or report.get("original_gcs_path") or report.get("original_drive_file_id")):
        return Response("这条记录没有保存原始文件。", status=404, mimetype="text/plain; charset=utf-8")
    try:
        original_bytes = download_original_from_storage(report)
    except Exception:
        app.logger.exception("Failed to download original upload %s", report_id)
        return Response("下载原始文件失败，请稍后再试。", status=500, mimetype="text/plain; charset=utf-8")
    return send_file(
        io.BytesIO(original_bytes),
        as_attachment=True,
        download_name=report.get("original_filename") or "thesis.docx",
        mimetype=original_content_type(report.get("original_filename") or "thesis.docx"),
    )


@app.post("/audit")
def audit():
    limited = rate_limit("audit")
    if limited:
        return limited
    user = current_user()
    if user is None:
        return audit_reject("请先注册或登录后再生成报告。", 401)
    if not is_account_active(user):
        return audit_reject(account_block_message(user), 403)
    if remaining_submissions(user) <= 0:
        return audit_reject("这个账号的检测额度已经用完。", 403)

    upload = request.files.get("docx")
    if not upload or not upload.filename:
        return audit_reject("请先选择一个 .doc 或 .docx 文件。", 400)

    names = uploaded_word_names(upload.filename)
    _safe_name, download_name = names
    original_filename = upload.filename.strip() or "thesis.docx"
    record_event("audit_submit", user, {"filename": original_filename})
    try:
        update_user_audit_trace(user["id"])
    except Exception:
        app.logger.warning("Failed to update audit trace for user %s", user["id"], exc_info=True)

    with tempfile.TemporaryDirectory(prefix="thesis-audit-") as tmp:
        tmp_path = Path(tmp)
        upload_path = tmp_path / f"{uuid4().hex}_{secure_filename(original_filename) or 'thesis'}"
        report_path = tmp_path / download_name
        upload.save(upload_path)

        try:
            audit_docx_path = prepare_docx_for_audit(upload_path, original_filename, tmp_path)
        except Exception as exc:
            record_event("audit_failed", user, {"filename": original_filename, "stage": "prepare", "error": str(exc)[:240]})
            return audit_reject(f"{exc}", 400)

        original_archive = archive_original_upload(user, upload_path, original_filename)
        college_info = {"college_name": UNKNOWN_COLLEGE, "college_source": "", "college_raw_text": ""}

        try:
            college_info = safe_extract_college_from_docx(audit_docx_path)
            run_audit_with_timeout(audit_docx_path, report_path)
        except Exception as exc:
            error_message = f"{exc}"
            try:
                create_report_record(
                    user=user,
                    original_filename=original_filename,
                    report_filename=download_name,
                    status="audit_failed",
                    college_name=college_info["college_name"],
                    college_source=college_info["college_source"],
                    college_raw_text=college_info["college_raw_text"],
                    original_storage_backend=original_archive["original_storage_backend"],
                    original_storage_path=original_archive["original_storage_path"],
                    original_gcs_path=original_archive["original_gcs_path"],
                    original_drive_file_id=original_archive["original_drive_file_id"],
                    original_drive_path=original_archive["original_drive_path"],
                    original_size_bytes=original_archive["original_size_bytes"],
                    original_sha256=original_archive["original_sha256"],
                    error_message=error_message,
                )
            except Exception:
                app.logger.warning("Failed to create failed report record for user %s", user["id"], exc_info=True)
            record_event("audit_failed", user, {"filename": original_filename, "stage": "audit", "error": error_message[:240], "college": college_info["college_name"]})
            return audit_error_response(exc)

        report_bytes = report_path.read_bytes()
        storage_path = report_storage_path_for(user, download_name)
        report_archive_path = storage_path
        report_status = "success"
        report_error_message = ""
        report_storage_backend = ""
        report_gcs_path = ""
        report_size_bytes = len(report_bytes)
        report_sha256 = sha256_hex(report_bytes)
        try:
            report_storage_backend = upload_report_to_storage(storage_path, report_bytes)
        except Exception as exc:
            storage_path = ""
            report_status = "storage_failed"
            report_error_message = f"报告已生成，但 Supabase 存档失败：{exc}"
            app.logger.warning("Failed to upload report to Supabase for user %s", user["id"], exc_info=True)
        if gcs_is_configured():
            try:
                report_gcs_path = upload_report_to_gcs(user, report_archive_path, report_bytes)
                if not storage_path:
                    report_status = "success"
                    report_error_message = report_error_message or "Supabase 存档失败，但 GCS 归档成功。"
            except Exception as exc:
                app.logger.warning("Failed to upload report to GCS for user %s", user["id"], exc_info=True)
                if report_error_message:
                    report_error_message += f"；GCS 归档失败：{exc}"
                else:
                    report_error_message = f"GCS 归档失败：{exc}"

        try:
            create_report_record(
                user=user,
                original_filename=original_filename,
                report_filename=download_name,
                status=report_status,
                college_name=college_info["college_name"],
                college_source=college_info["college_source"],
                college_raw_text=college_info["college_raw_text"],
                report_storage_path=storage_path,
                original_storage_backend=original_archive["original_storage_backend"],
                original_storage_path=original_archive["original_storage_path"],
                original_gcs_path=original_archive["original_gcs_path"],
                original_drive_file_id=original_archive["original_drive_file_id"],
                original_drive_path=original_archive["original_drive_path"],
                original_size_bytes=original_archive["original_size_bytes"],
                original_sha256=original_archive["original_sha256"],
                report_storage_backend=report_storage_backend,
                report_gcs_path=report_gcs_path,
                report_size_bytes=report_size_bytes,
                report_sha256=report_sha256,
                error_message=report_error_message,
            )
        except Exception:
            app.logger.warning("Failed to create report record for user %s", user["id"], exc_info=True)

        increment_submissions(user["id"])
        fresh_user = find_user_by_id(user["id"])
        remaining = remaining_submissions(fresh_user)
        record_event(
            "audit_success",
            user,
            {
                "filename": original_filename,
                "status": report_status,
                "college": college_info["college_name"],
                "remaining": remaining,
                "original_size_bytes": original_archive["original_size_bytes"],
            },
        )

        response = send_file(io.BytesIO(report_bytes), as_attachment=True, download_name=download_name, mimetype="text/html")
        response.headers["X-Remaining-Submissions"] = str(remaining)
        return response


@app.get("/health")
def health() -> str:
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=True)
