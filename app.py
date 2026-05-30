from __future__ import annotations

import logging
import multiprocessing
import os
import queue
import random
import tempfile
import time
from collections import defaultdict
from pathlib import Path
from uuid import uuid4

from flask import Flask, Response, redirect, render_template_string, request, send_file, session, url_for
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from postgrest.exceptions import APIError
from supabase import Client, create_client
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
from zipfile import BadZipFile, ZipFile

from thesis_format_audit import run_audit


app = Flask(__name__)
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "32"))
MAX_DOCX_ENTRIES = int(os.environ.get("MAX_DOCX_ENTRIES", "1500"))
MAX_DOCX_UNCOMPRESSED_MB = int(os.environ.get("MAX_DOCX_UNCOMPRESSED_MB", "180"))
AUDIT_TIMEOUT_SECONDS = int(os.environ.get("AUDIT_TIMEOUT_SECONDS", "105"))

app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["SESSION_COOKIE_SAMESITE"] = "None"
app.config["SESSION_COOKIE_SECURE"] = True
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")
logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO"))

MAX_SUBMISSIONS = 3
AUTH_TOKEN_MAX_AGE = 7 * 24 * 60 * 60
ACCOUNT_STATUS_ACTIVE = "active"
ACCOUNT_STATUS_FROZEN = "frozen"
ACCOUNT_STATUS_DISABLED = "disabled"
RATE_LIMITS = {
    "login": (10, 5 * 60),
    "register": (5, 60 * 60),
    "audit": (8, 60 * 60),
    "admin": (30, 5 * 60),
}
RATE_BUCKETS: defaultdict[tuple[str, str], list[float]] = defaultdict(list)
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
SUPABASE_TABLE = "thesis_audit_users"
ADMIN_EMAILS = {
    email.strip().lower()
    for email in os.environ.get("ADMIN_EMAILS", "2818242447@qq.com").split(",")
    if email.strip()
}


def uploaded_docx_names(filename: str) -> tuple[str, str] | None:
    raw_name = (filename or "").strip()
    if not raw_name.lower().endswith(".docx"):
        return None
    safe_name = secure_filename(raw_name)
    if not safe_name:
        safe_name = "thesis.docx"
    elif not safe_name.lower().endswith(".docx"):
        # Werkzeug may strip non-ASCII filenames down to "docx"; keep the
        # original extension decision but use a safe server-side name.
        safe_name = f"{safe_name}.docx"
    display_name = raw_name.replace("\\", "/").rsplit("/", 1)[-1]
    display_stem = display_name[:-5].strip() or Path(safe_name).stem or "thesis"
    return safe_name, f"{display_stem}_format_audit_report.html"


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
        raise ValueError("这个文件扩展名是 .docx，但内部不是有效的 Word 文档包。请重新另存为 .docx 后上传。") from exc


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


def find_user_by_id(user_id: str) -> dict | None:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select("id,email,password_hash,submissions_used,submission_quota,account_status,created_at")
        .eq("id", user_id)
        .maybe_single()
        .execute()
    )
    return result.data


def find_user_by_email(email: str) -> dict | None:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select("id,email,password_hash,submissions_used,submission_quota,account_status,created_at")
        .eq("email", email)
        .maybe_single()
        .execute()
    )
    return result.data


def create_user(email: str, password: str) -> dict:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .insert(
            {
                "email": email,
                "password_hash": generate_password_hash(password),
                "account_status": ACCOUNT_STATUS_ACTIVE,
            }
        )
        .execute()
    )
    return result.data[0]


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
    return find_user_by_id(user_id)


def remaining_submissions(user: dict | None) -> int:
    if user is None:
        return 0
    quota = int(user.get("submission_quota", MAX_SUBMISSIONS))
    return max(0, quota - int(user["submissions_used"]))


def user_quota(user: dict | None) -> int:
    if user is None:
        return MAX_SUBMISSIONS
    return int(user.get("submission_quota", MAX_SUBMISSIONS))


def is_admin(user: dict | None) -> bool:
    return bool(user and user.get("email", "").lower() in ADMIN_EMAILS)


def list_users() -> list[dict]:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select("id,email,submissions_used,submission_quota,account_status,created_at")
        .order("created_at", desc=True)
        .execute()
    )
    return result.data or []


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


def registration_values(email: str, password: str, confirm_password: str) -> dict:
    return {
        "register_email": email,
        "register_password": password,
        "register_confirm_password": confirm_password,
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
      margin-top: 18px;
      padding: 15px 16px;
      border: 1px solid color-mix(in srgb, var(--accent) 44%, var(--line));
      background:
        linear-gradient(135deg, color-mix(in srgb, var(--accent-soft) 54%, transparent), transparent 70%),
        var(--surface-strong);
      color: var(--accent-strong);
      font: 700 13px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .group-invite strong {
      display: block;
      margin-bottom: 4px;
      color: var(--ink);
      font-size: 15px;
    }
    .group-invite-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 8px;
    }
    .group-invite-head strong {
      margin: 0;
    }
    .group-number {
      margin-top: 6px;
      font: 900 18px/1.2 "PingFang SC", "Noto Sans SC", sans-serif;
      letter-spacing: .05em;
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
      width: min(480px, 100%);
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
      .quota-help { grid-template-columns: 1fr; }
      .quota-actions { align-items: flex-start; }
      .quota-number { width: fit-content; }
      .group-invite-head { align-items: flex-start; flex-direction: column; }
      .modal-card { padding: 20px; }
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
          <li>仅支持 .docx</li>
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
            <span>
              {% if is_admin %}<a class="logout-link" href="{{ url_for('admin', auth_token=auth_token) }}">管理后台</a>{% endif %}
              <a class="logout-link" href="{{ url_for('logout') }}">退出登录</a>
            </span>
          </div>
          <form id="audit-form" method="post" action="{{ url_for('audit') }}" enctype="multipart/form-data">
            {% if auth_token %}<input name="auth_token" type="hidden" value="{{ auth_token }}">{% endif %}
            {% if error %}<p class="error">{{ error }}</p>{% endif %}
            {% if remaining > 0 %}
              <p class="usage">每个账号最多可生成 {{ max_submissions }} 次报告。</p>
              <div class="quota-help">
                <p><span class="quota-label">增加检测次数</span>加入官方 QQ 群可领取检测机会，也可以联系管理员增加账号额度。</p>
                <div class="quota-actions">
                  <strong class="quota-number">537124215</strong>
                  <button class="copy-button" type="button" data-copy-group>一键复制群号</button>
                </div>
              </div>
              <label for="docx">选择论文文件</label>
              <label id="upload-card" class="upload-card" for="docx">
                <span class="upload-icon">↑</span>
                <span>
                  <span id="upload-title" class="upload-title">点击选择 Word 论文</span>
                  <span id="upload-meta" class="upload-meta">支持 .docx，文件选择后会显示名称；生成完成后自动下载 HTML 报告。</span>
                </span>
              </label>
              <input id="docx" name="docx" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required>
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
              <div class="quota-help">
                <p><span class="quota-label">额度已用完</span>加入官方 QQ 群可领取检测机会，也可以联系管理员增加检测次数。</p>
                <div class="quota-actions">
                  <strong class="quota-number">537124215</strong>
                  <button class="copy-button" type="button" data-copy-group>一键复制群号</button>
                </div>
              </div>
            {% endif %}
          </form>
        {% else %}
          <div class="panel-title"><strong>开始使用</strong><span>账号限制</span></div>
          {% if auth_error %}<p class="error">{{ auth_error }}</p>{% endif %}
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
              <p class="auth-copy">创建账号后可生成 {{ max_submissions }} 次报告。请确认密码并完成数字验证。</p>
              <input name="email" type="email" placeholder="邮箱" autocomplete="email" value="{{ auth_values.get('register_email', '') }}" required>
              <input name="password" type="password" placeholder="至少 6 位密码" autocomplete="new-password" minlength="6" value="{{ auth_values.get('register_password', '') }}" required>
              <input name="confirm_password" type="password" placeholder="再次输入密码" autocomplete="new-password" minlength="6" value="{{ auth_values.get('register_confirm_password', '') }}" required>
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
            <div class="auth-rule"><span>1</span><p>每个账号最多生成 {{ max_submissions }} 次报告，次数保存在数据库中。</p></div>
            <div class="auth-rule"><span>2</span><p>数字验证只用于减少自动注册，不会收集额外信息。</p></div>
          </div>
          <div class="group-invite">
            <div class="group-invite-head">
              <strong>加入 QQ 群可领取检测机会</strong>
              <button class="copy-button" type="button" data-copy-group>一键复制群号</button>
            </div>
            <div class="group-number">QQ 群号：537124215</div>
            进群后可联系管理员领取额外检测机会。
          </div>
        {% endif %}
      </div>
    </section>
  </main>
  <div id="info-modal" class="modal" aria-hidden="true">
    <div class="modal-card" role="dialog" aria-modal="true" aria-labelledby="modal-title" aria-describedby="modal-body">
      <h3 id="modal-title">提示</h3>
      <p id="modal-body"></p>
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
      onPrimary = null,
      onSecondary = null
    }) => {
      if (!infoModal || !modalTitle || !modalBody || !modalPrimary || !modalSecondary) return;
      modalTitle.textContent = title;
      modalBody.textContent = body;
      modalPrimary.textContent = primaryText;
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
        if (navigator.clipboard?.writeText) {
          await navigator.clipboard.writeText(GROUP_NUMBER);
        } else {
          const helper = document.createElement('input');
          helper.value = GROUP_NUMBER;
          document.body.appendChild(helper);
          helper.select();
          document.execCommand('copy');
          helper.remove();
        }
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
          credentials: 'same-origin'
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
        showModal({
          title: '下载成功',
          body: '检测报告已经下载成功，请到浏览器下载列表或下载文件夹中找到该 HTML 文件，并用浏览器打开查看结果。'
        });
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
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 16px;
      margin-bottom: 20px;
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
    .toolbar {
      display: grid;
      grid-template-columns: minmax(0, 1.2fr) repeat(3, minmax(160px, .42fr));
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
    .table-wrap {
      overflow-x: auto;
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
    tbody tr {
      transition: background .16s ease;
    }
    tbody tr:hover {
      background: rgba(255, 255, 255, .55);
    }
    .email { font-weight: 800; }
    .muted { color: var(--muted); }
    .mono {
      font-family: "SFMono-Regular", "Menlo", monospace;
      font-size: 12px;
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
    .actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }
    .action-group {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
    }
    button {
      border: 1px solid transparent;
      padding: 10px 12px;
      background: var(--accent);
      color: white;
      cursor: pointer;
      font-weight: 800;
      transition: transform .16s ease, filter .16s ease, background .16s ease;
    }
    button:hover {
      filter: brightness(.95);
      transform: translateY(-1px);
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
      padding: 9px 11px;
      font-size: 13px;
    }
    .inline-form {
      display: flex;
      gap: 8px;
      align-items: center;
      flex-wrap: wrap;
    }
    .inline-number {
      width: 88px;
      padding: 9px 10px;
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
      .toolbar {
        grid-template-columns: 1fr 1fr;
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
      .toolbar {
        grid-template-columns: 1fr;
      }
      .summary-bar {
        align-items: flex-start;
        flex-direction: column;
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
          <a class="top-link" href="{{ url_for('index', auth_token=auth_token) }}">返回检测页</a>
          <a class="top-link" href="{{ url_for('logout') }}">退出登录</a>
        </div>
      </div>
    </div>
    {% if message %}<div class="notice">{{ message }}</div>{% endif %}
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
    </section>
    <section class="panel">
      <div class="toolbar">
        <input id="search-input" class="field" type="search" placeholder="搜索邮箱、用户 ID 或注册时间">
        <select id="status-filter" class="select">
          <option value="all">全部状态</option>
          <option value="active">仅看正常</option>
          <option value="frozen">仅看冻结</option>
          <option value="disabled">仅看已注销</option>
        </select>
        <select id="quota-filter" class="select">
          <option value="all">全部额度情况</option>
          <option value="remaining">仅看仍有次数</option>
          <option value="empty">仅看额度用完</option>
        </select>
        <button id="reset-filter" class="ghost-button" type="button">重置筛选</button>
      </div>
      <div class="summary-bar">
        <span>共 <strong id="visible-count">{{ users|length }}</strong> 个账号正在显示</span>
        <span>支持快速额度发放、自定义加额、冻结、恢复和注销操作</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>账号</th>
              <th>状态</th>
              <th>已用 / 总额度</th>
              <th>剩余</th>
              <th>注册时间</th>
              <th>增加次数</th>
              <th>账号操作</th>
            </tr>
          </thead>
          <tbody id="user-table">
            {% for item in users %}
              {% set remaining = [item["submission_quota"] - item["submissions_used"], 0] | max %}
              <tr
                data-search="{{ item['email'] }} {{ item['id'] }} {{ item['created_at'] }}"
                data-status="{{ item['account_status'] }}"
                data-remaining="{{ remaining }}"
              >
                <td data-label="账号">
                  <div class="email">{{ item["email"] }}</div>
                  <div class="muted mono">{{ item["id"] }}</div>
                </td>
                <td data-label="状态">
                  <span class="status-badge status-{{ item['account_status'] }}">
                    {{ status_labels.get(item["account_status"], "未知") }}
                  </span>
                </td>
                <td data-label="已用 / 总额度"><span class="quota">{{ item["submissions_used"] }} / {{ item["submission_quota"] }}</span></td>
                <td data-label="剩余"><strong>{{ remaining }}</strong></td>
                <td data-label="注册时间" class="muted">{{ item["created_at"] }}</td>
                <td data-label="增加次数">
                  <div class="actions">
                    <div class="action-group">
                      {% for amount in [1, 3, 10] %}
                        <form method="post" action="{{ url_for('admin_add_quota') }}">
                          <input name="auth_token" type="hidden" value="{{ auth_token }}">
                          <input name="user_id" type="hidden" value="{{ item["id"] }}">
                          <input name="amount" type="hidden" value="{{ amount }}">
                          <button class="compact-button" type="submit">+{{ amount }}</button>
                        </form>
                      {% endfor %}
                    </div>
                    <form class="inline-form" method="post" action="{{ url_for('admin_add_quota') }}">
                      <input name="auth_token" type="hidden" value="{{ auth_token }}">
                      <input name="user_id" type="hidden" value="{{ item["id"] }}">
                      <input class="inline-number" name="amount" type="number" min="1" step="1" placeholder="自定义">
                      <button class="ghost-button compact-button" type="submit">增加</button>
                    </form>
                  </div>
                </td>
                <td data-label="账号操作">
                  <div class="actions">
                    {% if item["account_status"] == "active" %}
                      <form method="post" action="{{ url_for('admin_update_status') }}" data-confirm="确认冻结 {{ item['email'] }} 吗？冻结后该账号将不能登录和检测。">
                        <input name="auth_token" type="hidden" value="{{ auth_token }}">
                        <input name="user_id" type="hidden" value="{{ item["id"] }}">
                        <input name="status" type="hidden" value="frozen">
                        <button class="info-button compact-button" type="submit">冻结</button>
                      </form>
                    {% else %}
                      <form method="post" action="{{ url_for('admin_update_status') }}" data-confirm="确认恢复 {{ item['email'] }} 吗？恢复后账号可继续使用。">
                        <input name="auth_token" type="hidden" value="{{ auth_token }}">
                        <input name="user_id" type="hidden" value="{{ item["id"] }}">
                        <input name="status" type="hidden" value="active">
                        <button class="ghost-button compact-button" type="submit">恢复</button>
                      </form>
                    {% endif %}
                    {% if item["account_status"] != "disabled" %}
                      <form method="post" action="{{ url_for('admin_update_status') }}" data-confirm="确认注销 {{ item['email'] }} 吗？注销后该账号将被停用。">
                        <input name="auth_token" type="hidden" value="{{ auth_token }}">
                        <input name="user_id" type="hidden" value="{{ item["id"] }}">
                        <input name="status" type="hidden" value="disabled">
                        <button class="warn-button compact-button" type="submit">注销</button>
                      </form>
                    {% endif %}
                  </div>
                </td>
              </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
      <div id="empty-state" class="empty-state" hidden>没有匹配到符合条件的账号，试试清空搜索词或调整筛选条件。</div>
    </section>
  </main>
  <script>
    const searchInput = document.getElementById('search-input');
    const statusFilter = document.getElementById('status-filter');
    const quotaFilter = document.getElementById('quota-filter');
    const resetFilter = document.getElementById('reset-filter');
    const userTable = document.getElementById('user-table');
    const visibleCount = document.getElementById('visible-count');
    const emptyState = document.getElementById('empty-state');

    const applyFilters = () => {
      if (!userTable) return;
      const keyword = (searchInput?.value || '').trim().toLowerCase();
      const status = statusFilter?.value || 'all';
      const quota = quotaFilter?.value || 'all';
      let shown = 0;

      userTable.querySelectorAll('tr').forEach(row => {
        const searchText = (row.dataset.search || '').toLowerCase();
        const rowStatus = row.dataset.status || 'active';
        const remaining = Number(row.dataset.remaining || '0');
        const matchesKeyword = !keyword || searchText.includes(keyword);
        const matchesStatus = status === 'all' || rowStatus === status;
        const matchesQuota = quota === 'all' || (quota === 'remaining' ? remaining > 0 : remaining <= 0);
        const visible = matchesKeyword && matchesStatus && matchesQuota;
        row.hidden = !visible;
        if (visible) shown += 1;
      });

      if (visibleCount) visibleCount.textContent = String(shown);
      if (emptyState) emptyState.hidden = shown !== 0;
    };

    [searchInput, statusFilter, quotaFilter].forEach(element => {
      if (!element) return;
      element.addEventListener('input', applyFilters);
      element.addEventListener('change', applyFilters);
    });

    if (resetFilter) resetFilter.addEventListener('click', () => {
      if (searchInput) searchInput.value = '';
      if (statusFilter) statusFilter.value = 'all';
      if (quotaFilter) quotaFilter.value = 'all';
      applyFilters();
    });

    document.querySelectorAll('form[data-confirm]').forEach(form => {
      form.addEventListener('submit', event => {
        const message = form.dataset.confirm || '确认继续吗？';
        if (!window.confirm(message)) event.preventDefault();
      });
    });

    applyFilters();
  </script>
</body>
</html>
"""


@app.get("/")
def index() -> str:
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
    captcha_answer = request.form.get("captcha_answer", "").strip()
    captcha_left = request.form.get("captcha_left", "")
    captcha_right = request.form.get("captcha_right", "")
    auth_values = registration_values(email, password, confirm_password)
    if not email or "@" not in email:
        refresh_captcha()
        return render_home(auth_error="请输入有效邮箱。", auth_mode="register", auth_values=auth_values), 400
    if len(password) < 6:
        refresh_captcha()
        return render_home(auth_error="密码至少需要 6 位。", auth_mode="register", auth_values=auth_values), 400
    if password != confirm_password:
        refresh_captcha()
        return render_home(auth_error="两次输入的密码不一致。", auth_mode="register", auth_values=auth_values), 400
    if not is_valid_captcha(captcha_answer, captcha_left, captcha_right):
        refresh_captcha()
        return render_home(auth_error="数字验证不正确，请重新计算。", auth_mode="register", auth_values=auth_values), 400

    try:
        user = create_user(email, password)
        session["user_id"] = user["id"]
    except APIError:
        refresh_captcha()
        return render_home(auth_error="这个邮箱已经注册，请直接登录。", auth_mode="register", auth_values=auth_values), 400
    return redirect(url_for("index", auth_token=generate_auth_token(user["id"])))


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
    session["user_id"] = user["id"]
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
    users = list_users()
    stats = {
        "total": len(users),
        "active": sum(1 for item in users if item.get("account_status", ACCOUNT_STATUS_ACTIVE) == ACCOUNT_STATUS_ACTIVE),
        "frozen": sum(1 for item in users if item.get("account_status") == ACCOUNT_STATUS_FROZEN),
        "disabled": sum(1 for item in users if item.get("account_status") == ACCOUNT_STATUS_DISABLED),
    }
    return render_template_string(
        ADMIN_PAGE,
        admin_user=user,
        users=users,
        stats=stats,
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
    limited = rate_limit("admin")
    if limited:
        return limited
    user = current_user()
    if not is_admin(user):
        return Response("没有权限。", status=403, mimetype="text/plain; charset=utf-8")
    token = request.form.get("auth_token") or generate_auth_token(user["id"])
    target_user_id = request.form.get("user_id", "")
    try:
        amount = int(request.form.get("amount", "0"))
    except ValueError:
        amount = 0
    if amount <= 0:
        return redirect(url_for("admin", auth_token=token, message="增加次数必须大于 0。"))
    try:
        add_user_quota(target_user_id, amount)
    except ValueError as exc:
        return redirect(url_for("admin", auth_token=token, message=str(exc)))
    return redirect(url_for("admin", auth_token=token, message=f"已增加 {amount} 次额度。"))


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
    if target_user_id == user.get("id") and status == ACCOUNT_STATUS_DISABLED:
        return redirect(url_for("admin", auth_token=token, message="不能注销当前管理员账号。"))
    try:
        update_user_status(target_user_id, status)
    except ValueError as exc:
        return redirect(url_for("admin", auth_token=token, message=str(exc)))
    return redirect(url_for("admin", auth_token=token, message=f"账号状态已更新为：{account_status_label(status)}。"))


@app.post("/audit")
def audit():
    limited = rate_limit("audit")
    if limited:
        return limited
    user = current_user()
    if user is None:
        return render_home(error="请先注册或登录后再生成报告。"), 401
    if not is_account_active(user):
        return render_home(error=account_block_message(user)), 403
    if remaining_submissions(user) <= 0:
        return render_home(error="这个账号的检测额度已经用完。"), 403

    upload = request.files.get("docx")
    if not upload or not upload.filename:
        return render_home(error="请先选择一个 .docx 文件。"), 400

    names = uploaded_docx_names(upload.filename)
    if names is None:
        return render_home(error="当前只支持 .docx 文件。"), 400
    _safe_name, download_name = names

    with tempfile.TemporaryDirectory(prefix="thesis-audit-") as tmp:
        tmp_path = Path(tmp)
        docx_path = tmp_path / f"{uuid4().hex}.docx"
        report_path = tmp_path / download_name
        upload.save(docx_path)

        try:
            validate_docx_package(docx_path)
            run_audit_with_timeout(docx_path, report_path)
        except Exception as exc:
            return audit_error_response(exc)

        increment_submissions(user["id"])
        fresh_user = find_user_by_id(user["id"])
        remaining = remaining_submissions(fresh_user)

        response = send_file(report_path, as_attachment=True, download_name=download_name, mimetype="text/html")
        response.headers["X-Remaining-Submissions"] = str(remaining)
        return response


@app.get("/health")
def health() -> str:
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=True)
