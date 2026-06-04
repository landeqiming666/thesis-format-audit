---
title: UPC本科论文格式检测工具
emoji: 📄
colorFrom: green
colorTo: gray
sdk: docker
pinned: false
---

# UPC本科论文格式检测工具

这个项目把原本的 macOS 命令行检测脚本包装成一个网页服务。用户上传 `.docx` 或旧版 `.doc` 后，服务会调用 `thesis_format_audit.py` 生成 HTML 检测报告，并把报告返回给浏览器下载；其中 `.doc` 会先通过 LibreOffice 自动转换为临时 `.docx` 再检测。

## 本地运行

```bash
uv venv .venv
source .venv/bin/activate
uv pip install -r requirements.txt
python app.py
```

项目根目录的 `uv.toml` 已配置清华 PyPI 镜像；如果不用 `uv`，也可以继续用
`pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt`。

打开：

```text
http://127.0.0.1:8000
```

## Render 部署

1. 把这个目录推到 GitHub、GitLab 或 Bitbucket。
2. 在 Render 新建 Blueprint，选择这个仓库。
3. Render 会读取 `render.yaml` 并创建免费 Python Web Service。

## Supabase 配置

注册、提交次数和后台检测记录使用 Supabase。

1. 在 Supabase SQL Editor 执行 `supabase_schema.sql`。
2. 确认 Supabase Storage 里存在私有 bucket：`thesis-audit-reports`。
3. 在部署平台设置环境变量：

```text
SUPABASE_URL=你的 Supabase Project URL
SUPABASE_SERVICE_ROLE_KEY=你的 Supabase service_role key
SECRET_KEY=任意一段随机长字符串
REPORTS_BUCKET=thesis-audit-reports
```

`SUPABASE_SERVICE_ROLE_KEY` 只能放在服务端环境变量里，不要写进前端页面或公开仓库。

## Gmail 注册验证码

注册邮箱验证码使用 Gmail SMTP。请使用 Google App Password，不要使用谷歌账号登录密码。

1. 打开 Google 账号安全设置。
2. 开启两步验证。
3. 在 App Passwords 里生成一个应用专用密码。
4. 在 `.env` 或部署平台环境变量里配置：

```text
GMAIL_SMTP_USER=你的 Gmail 邮箱
GMAIL_SMTP_APP_PASSWORD=Google App Password
GMAIL_SMTP_HOST=smtp.gmail.com
GMAIL_SMTP_PORT=465
EMAIL_FROM_NAME=UPC论文格式检测工具
```

本地可以先运行测试脚本确认发信正常：

```bash
python check_email_smtp.py 收件邮箱@example.com
```

## Google Cloud Storage 可选归档

如果要启用 Google Cloud Storage 作为第二存储后端，配置这些环境变量即可。没有配置时系统会继续只用 Supabase，不影响正常检测。

```text
GCS_BUCKET=你的 Google Cloud Storage bucket 名
GCS_PREFIX=thesis-audit
GCS_PROJECT=你的 Google Cloud project id
GOOGLE_APPLICATION_CREDENTIALS_JSON=服务账号 JSON 内容
```

建议给服务账号最小权限，只授予目标 bucket 的对象读写权限。`GOOGLE_APPLICATION_CREDENTIALS_JSON` 是密钥，不能提交到 GitHub 或 Hugging Face 仓库。

## Hugging Face Spaces 部署

这个项目也可以部署到 Hugging Face Spaces，适合没有信用卡、只想给别人一个公开访问链接的场景。

使用 Docker SDK 时，Spaces 会运行 `Dockerfile` 中的启动命令。
