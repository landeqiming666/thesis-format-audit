from __future__ import annotations

import os
import random
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


def render_home(error: str = "", auth_error: str = "", auth_mode: str = "login") -> str:
    user = current_user()
    if "captcha_answer" not in session:
        refresh_captcha()
    return render_template_string(
        PAGE,
        user=user,
        configured=bool(SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY),
        remaining=remaining_submissions(user),
        max_submissions=MAX_SUBMISSIONS,
        captcha_question=session.get("captcha_question", ""),
        auth_mode=auth_mode,
        error=error,
        auth_error=auth_error,
    )


def refresh_captcha() -> None:
    left = random.randint(2, 9)
    right = random.randint(1, 8)
    session["captcha_question"] = f"{left} + {right} = ?"
    session["captcha_answer"] = str(left + right)


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
      width: 100%;
      padding: 18px;
      border: 1px dashed #8fa099;
      background: var(--field);
      color: var(--muted);
      font: 15px/1.5 "PingFang SC", "Noto Sans SC", sans-serif;
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
      h1 { font-size: clamp(46px, 16vw, 68px); }
    }
  </style>
</head>
<body>
  <main>
    <header class="topbar">
      <div class="brand">
        <span class="brand-mark">审</span>
        <span>Thesis Format Audit</span>
      </div>
      <button id="theme-toggle" class="theme-toggle" type="button" aria-label="切换夜间模式">夜间模式</button>
    </header>
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
          <div class="panel-title"><strong>系统配置</strong><span>Database</span></div>
          <p class="error">服务还没有配置 Supabase 数据库。</p>
          <p class="note">管理员需要设置 SUPABASE_URL 和 SUPABASE_SERVICE_ROLE_KEY 后才能开放注册登录。</p>
        {% elif user %}
          <div class="panel-title"><strong>生成报告</strong><span>3 次额度</span></div>
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
              <input name="email" type="email" placeholder="邮箱" autocomplete="email" required>
              <input name="password" type="password" placeholder="密码" autocomplete="current-password" required>
              <button type="submit">登录后检测</button>
            </form>
            <form class="auth-box {% if auth_mode == 'register' %}active{% endif %}" method="post" action="{{ url_for('register') }}" data-auth-panel="register">
              <h2>注册</h2>
              <p class="auth-copy">创建账号后可生成 {{ max_submissions }} 次报告。请确认密码并完成数字验证。</p>
              <input name="email" type="email" placeholder="邮箱" autocomplete="email" required>
              <input name="password" type="password" placeholder="至少 6 位密码" autocomplete="new-password" minlength="6" required>
              <input name="confirm_password" type="password" placeholder="再次输入密码" autocomplete="new-password" minlength="6" required>
              <div class="captcha-row">
                <div class="captcha-chip">{{ captcha_question }}</div>
                <input name="captcha_answer" type="text" inputmode="numeric" pattern="[0-9]*" placeholder="输入计算结果" required>
              </div>
              <button type="submit">创建账号</button>
            </form>
          </div>
          <div class="auth-rules">
            <div class="auth-rule"><span>1</span><p>每个账号最多生成 {{ max_submissions }} 次报告，次数保存在数据库中。</p></div>
            <div class="auth-rule"><span>2</span><p>数字验证只用于减少自动注册，不会收集额外信息。</p></div>
          </div>
        {% endif %}
      </div>
    </section>
  </main>
  <script>
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
        return render_home(auth_error="服务还没有配置 Supabase 数据库。", auth_mode="register"), 503
    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")
    confirm_password = request.form.get("confirm_password", "")
    captcha_answer = request.form.get("captcha_answer", "").strip()
    if not email or "@" not in email:
        refresh_captcha()
        return render_home(auth_error="请输入有效邮箱。", auth_mode="register"), 400
    if len(password) < 6:
        refresh_captcha()
        return render_home(auth_error="密码至少需要 6 位。", auth_mode="register"), 400
    if password != confirm_password:
        refresh_captcha()
        return render_home(auth_error="两次输入的密码不一致。", auth_mode="register"), 400
    if captcha_answer != session.get("captcha_answer"):
        refresh_captcha()
        return render_home(auth_error="数字验证不正确，请重新计算。", auth_mode="register"), 400

    try:
        user = create_user(email, password)
        session["user_id"] = user["id"]
    except APIError:
        refresh_captcha()
        return render_home(auth_error="这个邮箱已经注册，请直接登录。", auth_mode="register"), 400
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
