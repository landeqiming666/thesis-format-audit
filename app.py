from __future__ import annotations

import tempfile
from pathlib import Path
from uuid import uuid4

from flask import Flask, Response, redirect, render_template_string, request, send_file, url_for
from werkzeug.utils import secure_filename

from thesis_format_audit import run_audit


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


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
    .note {
      margin: 18px 0 0;
      color: var(--muted);
      font: 13px/1.8 "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .error {
      margin: 0 0 16px;
      color: var(--warn);
      font: 700 14px/1.6 "PingFang SC", "Noto Sans SC", sans-serif;
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
      <form class="panel" method="post" action="{{ url_for('audit') }}" enctype="multipart/form-data">
        {% if error %}<p class="error">{{ error }}</p>{% endif %}
        <label for="docx">选择论文文件</label>
        <input id="docx" name="docx" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required>
        <button type="submit">生成检测报告</button>
        <p class="note">报告会在浏览器中下载为 HTML 文件，可以直接打开或转发。</p>
      </form>
    </section>
  </main>
</body>
</html>
"""


@app.get("/")
def index() -> str:
    return render_template_string(PAGE)


@app.post("/audit")
def audit():
    upload = request.files.get("docx")
    if not upload or not upload.filename:
        return render_template_string(PAGE, error="请先选择一个 .docx 文件。"), 400

    original_name = secure_filename(upload.filename) or "thesis.docx"
    if not original_name.lower().endswith(".docx"):
        return render_template_string(PAGE, error="当前只支持 .docx 文件。"), 400

    with tempfile.TemporaryDirectory(prefix="thesis-audit-") as tmp:
        tmp_path = Path(tmp)
        docx_path = tmp_path / f"{uuid4().hex}.docx"
        report_path = tmp_path / f"{Path(original_name).stem}_format_audit_report.html"
        upload.save(docx_path)

        try:
            run_audit(docx_path, report_path)
        except Exception as exc:
            return Response(f"检测失败：{exc}", status=500, mimetype="text/plain; charset=utf-8")

        download_name = f"{Path(original_name).stem}_format_audit_report.html"
        return send_file(report_path, as_attachment=True, download_name=download_name, mimetype="text/html")


@app.get("/health")
def health() -> str:
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
