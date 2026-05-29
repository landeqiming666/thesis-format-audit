from __future__ import annotations

import os
import tempfile
from pathlib import Path
from uuid import uuid4

from flask import Flask, Response, redirect, render_template_string, request, send_file, session, url_for
from postgrest.exceptions import APIError
from supabase import Client, create_client
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from thesis_format_audit import run_audit


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")

MAX_SUBMISSIONS = 3
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
SUPABASE_TABLE = "thesis_audit_users"


def get_supabase() -> Client:
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        raise RuntimeError("Supabase is not configured. Set SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY.")
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)


def find_user_by_id(user_id: str) -> dict | None:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select("id,email,password_hash,submissions_used")
        .eq("id", user_id)
        .maybe_single()
        .execute()
    )
    return result.data


def find_user_by_email(email: str) -> dict | None:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .select("id,email,password_hash,submissions_used")
        .eq("email", email)
        .maybe_single()
        .execute()
    )
    return result.data


def create_user(email: str, password: str) -> dict:
    result = (
        get_supabase()
        .table(SUPABASE_TABLE)
        .insert({"email": email, "password_hash": generate_password_hash(password)})
        .execute()
    )
    return result.data[0]


def increment_submissions(user_id: str) -> None:
    result = get_supabase().rpc(
        "increment_thesis_audit_submissions",
        {"target_user_id": user_id, "max_allowed": MAX_SUBMISSIONS},
    ).execute()
    if result.data is not True:
        raise RuntimeError("Submission limit reached.")


def current_user() -> dict | None:
    user_id = session.get("user_id")
    if not user_id:
        return None
    return find_user_by_id(user_id)


def remaining_submissions(user: dict | None) -> int:
    if user is None:
        return 0
    return max(0, MAX_SUBMISSIONS - int(user["submissions_used"]))


def render_home(error: str = "", auth_error: str = "") -> str:
    user = current_user()
    return render_template_string(
        PAGE,
        user=user,
        configured=bool(SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY),
        remaining=remaining_submissions(user),
        max_submissions=MAX_SUBMISSIONS,
        error=error,
        auth_error=auth_error,
    )


PAGE = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>论文格式检测</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #17202a;
      --muted: #5f6b76;
      --paper: #fbfaf7;
      --line: #d8ded9;
      --field: #ffffff;
      --accent: #1f7a5c;
      --accent-strong: #0f5f45;
      --warn: #a1432f;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100svh;
      font-family: "Songti SC", "Noto Serif SC", "STSong", serif;
      color: var(--ink);
      background:
        linear-gradient(120deg, rgba(31, 122, 92, .10), transparent 34%),
        repeating-linear-gradient(0deg, rgba(23, 32, 42, .035), rgba(23, 32, 42, .035) 1px, transparent 1px, transparent 34px),
        var(--paper);
    }
    main {
      width: min(1080px, calc(100% - 32px));
      margin: 0 auto;
      padding: 48px 0;
    }
    .shell {
      display: grid;
      grid-template-columns: minmax(0, 1.05fr) minmax(320px, .95fr);
      gap: 44px;
      align-items: center;
      min-height: calc(100svh - 96px);
    }
    .mark {
      width: 72px;
      height: 5px;
      margin-bottom: 28px;
      background: var(--accent);
    }
    h1 {
      margin: 0;
      max-width: 740px;
      font-size: clamp(42px, 7vw, 86px);
      line-height: .98;
      font-weight: 800;
      letter-spacing: 0;
    }
    .lead {
      max-width: 560px;
      margin: 24px 0 0;
      color: var(--muted);
      font-size: 19px;
      line-height: 1.8;
    }
    .panel {
      border: 1px solid var(--line);
      background: rgba(255, 255, 255, .82);
      padding: 26px;
      box-shadow: 0 24px 80px rgba(23, 32, 42, .10);
      backdrop-filter: blur(14px);
    }
    label {
      display: block;
      margin-bottom: 12px;
      font: 700 15px/1.4 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    input[type="file"] {
      width: 100%;
      padding: 18px;
      border: 1px dashed #8fa099;
      background: var(--field);
      color: var(--muted);
      font: 15px/1.5 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    input[type="email"],
    input[type="password"] {
      width: 100%;
      margin-bottom: 12px;
      padding: 13px 14px;
      border: 1px solid var(--line);
      background: var(--field);
      color: var(--ink);
      font: 15px/1.5 "PingFang SC", "Noto Sans SC", sans-serif;
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
      background: #dfe7e2;
      border: 1px solid #cad5cf;
    }
    .progress-bar {
      width: 0%;
      height: 100%;
      background: linear-gradient(90deg, var(--accent), #54aa82);
      transition: width .45s ease;
    }
    .error {
      margin: 0 0 16px;
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
    }
    .auth-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 18px;
    }
    .auth-box h2 {
      margin: 0 0 14px;
      font: 800 20px/1.3 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .auth-box button { margin-top: 4px; }
    .usage {
      margin: 0 0 16px;
      color: var(--muted);
      font: 14px/1.7 "PingFang SC", "Noto Sans SC", sans-serif;
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
    @media (max-width: 820px) {
      main { padding: 28px 0; }
      .shell { grid-template-columns: 1fr; min-height: auto; }
      .panel { padding: 20px; }
      .auth-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main>
    <section class="shell">
      <div>
        <div class="mark"></div>
        <h1>本科论文格式检测</h1>
        <p class="lead">上传 Word 论文，系统会检查摘要、目录、标题、正文、图表、公式、参考文献和页码，并生成可交互的 HTML 报告。</p>
        <ul class="facts">
          <li>仅支持 .docx</li>
          <li>单文件 32MB 内</li>
          <li>检测过程不改原文</li>
        </ul>
      </div>
      <div class="panel">
        {% if not configured %}
          <p class="error">服务还没有配置 Supabase 数据库。</p>
          <p class="note">管理员需要设置 SUPABASE_URL 和 SUPABASE_SERVICE_ROLE_KEY 后才能开放注册登录。</p>
        {% elif user %}
          <div class="account-bar">
            <span>当前账号：<strong>{{ user["email"] }}</strong><br>剩余次数：{{ remaining }} / {{ max_submissions }}</span>
            <a class="logout-link" href="{{ url_for('logout') }}">退出登录</a>
          </div>
          <form id="audit-form" method="post" action="{{ url_for('audit') }}" enctype="multipart/form-data">
            {% if error %}<p class="error">{{ error }}</p>{% endif %}
            {% if remaining > 0 %}
              <p class="usage">每个账号最多可生成 {{ max_submissions }} 次报告。</p>
              <label for="docx">选择论文文件</label>
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
              <p class="note">报告会在浏览器中下载为 HTML 文件，可以直接打开或转发。大文件可能需要等待几十秒。</p>
            {% else %}
              <p class="error">这个账号的 3 次检测额度已经用完。</p>
              <p class="note">如果需要继续使用，请联系管理员增加额度或更换账号。</p>
            {% endif %}
          </form>
        {% else %}
          {% if auth_error %}<p class="error">{{ auth_error }}</p>{% endif %}
          <div class="auth-grid">
            <form class="auth-box" method="post" action="{{ url_for('login') }}">
              <h2>登录</h2>
              <input name="email" type="email" placeholder="邮箱" autocomplete="email" required>
              <input name="password" type="password" placeholder="密码" autocomplete="current-password" required>
              <button type="submit">登录后检测</button>
            </form>
            <form class="auth-box" method="post" action="{{ url_for('register') }}">
              <h2>注册</h2>
              <input name="email" type="email" placeholder="邮箱" autocomplete="email" required>
              <input name="password" type="password" placeholder="至少 6 位密码" autocomplete="new-password" minlength="6" required>
              <button type="submit">创建账号</button>
            </form>
          </div>
          <p class="note">每个账号最多可生成 {{ max_submissions }} 次报告。</p>
        {% endif %}
      </div>
    </section>
  </main>
  <script>
    const form = document.getElementById('audit-form');
    const fileInput = document.getElementById('docx');
    const submitButton = document.getElementById('submit-button');
    const progressWrap = document.getElementById('progress-wrap');
    const progressBar = document.getElementById('progress-bar');
    const progressPercent = document.getElementById('progress-percent');
    const progressMessage = document.getElementById('progress-message');

    const messages = [
      [10, '正在上传论文...'],
      [28, '正在读取 Word 结构...'],
      [46, '正在检查摘要、目录和标题...'],
      [64, '正在检查正文、图表和公式...'],
      [82, '正在生成 HTML 报告...'],
      [92, '报告快好了，请稍等...']
    ];

    if (form && fileInput && submitButton && progressWrap) form.addEventListener('submit', () => {
      if (!fileInput.files.length) return;

      submitButton.disabled = true;
      submitButton.textContent = '检测中，请稍等...';
      progressWrap.classList.add('active');

      let progress = 0;
      const tick = () => {
        const nextLimit = progress < 30 ? 30 : progress < 70 ? 70 : 92;
        const step = progress < 30 ? 6 : progress < 70 ? 3 : 1;
        progress = Math.min(nextLimit, progress + step);
        progressBar.style.width = `${progress}%`;
        progressPercent.textContent = `${progress}%`;

        const current = [...messages].reverse().find(([limit]) => progress >= limit);
        if (current) progressMessage.textContent = current[1];
      };

      tick();
      window.setInterval(tick, 900);
    });
  </script>
</body>
</html>
"""


@app.get("/")
def index() -> str:
    return render_home()


@app.post("/register")
def register():
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        return render_home(auth_error="服务还没有配置 Supabase 数据库。"), 503
    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")
    if not email or "@" not in email:
        return render_home(auth_error="请输入有效邮箱。"), 400
    if len(password) < 6:
        return render_home(auth_error="密码至少需要 6 位。"), 400

    try:
        user = create_user(email, password)
        session["user_id"] = user["id"]
    except APIError:
        return render_home(auth_error="这个邮箱已经注册，请直接登录。"), 400
    return redirect(url_for("index"))


@app.post("/login")
def login():
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        return render_home(auth_error="服务还没有配置 Supabase 数据库。"), 503
    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")
    user = find_user_by_email(email)
    if user is None or not check_password_hash(user["password_hash"], password):
        return render_home(auth_error="邮箱或密码不正确。"), 400
    session["user_id"] = user["id"]
    return redirect(url_for("index"))


@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.post("/audit")
def audit():
    user = current_user()
    if user is None:
        return render_home(error="请先注册或登录后再生成报告。"), 401
    if remaining_submissions(user) <= 0:
        return render_home(error="这个账号的 3 次检测额度已经用完。"), 403

    upload = request.files.get("docx")
    if not upload or not upload.filename:
        return render_home(error="请先选择一个 .docx 文件。"), 400

    original_name = secure_filename(upload.filename) or "thesis.docx"
    if not original_name.lower().endswith(".docx"):
        return render_home(error="当前只支持 .docx 文件。"), 400

    with tempfile.TemporaryDirectory(prefix="thesis-audit-") as tmp:
        tmp_path = Path(tmp)
        docx_path = tmp_path / f"{uuid4().hex}.docx"
        report_path = tmp_path / f"{Path(original_name).stem}_format_audit_report.html"
        upload.save(docx_path)

        try:
            run_audit(docx_path, report_path)
        except Exception as exc:
            return Response(f"检测失败：{exc}", status=500, mimetype="text/plain; charset=utf-8")

        increment_submissions(user["id"])

        download_name = f"{Path(original_name).stem}_format_audit_report.html"
        return send_file(report_path, as_attachment=True, download_name=download_name, mimetype="text/html")


@app.get("/health")
def health() -> str:
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=True)
