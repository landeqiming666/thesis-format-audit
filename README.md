---
title: UPC本科论文格式检测工具
emoji: 📄
colorFrom: green
colorTo: gray
sdk: docker
pinned: false
---

# UPC本科论文格式检测工具

这个项目把原本的 macOS 命令行检测脚本包装成一个网页服务。用户上传 `.docx` 后，服务会调用 `thesis_format_audit.py` 生成 HTML 检测报告，并把报告返回给浏览器下载。

## 本地运行

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

打开：

```text
http://127.0.0.1:8000
```

## Render 部署

1. 把这个目录推到 GitHub、GitLab 或 Bitbucket。
2. 在 Render 新建 Blueprint，选择这个仓库。
3. Render 会读取 `render.yaml` 并创建免费 Python Web Service。

## Supabase 配置

注册和 3 次提交限制使用 Supabase 保存。

1. 在 Supabase SQL Editor 执行 `supabase_schema.sql`。
2. 在部署平台设置环境变量：

```text
SUPABASE_URL=你的 Supabase Project URL
SUPABASE_SERVICE_ROLE_KEY=你的 Supabase service_role key
SECRET_KEY=任意一段随机长字符串
```

`SUPABASE_SERVICE_ROLE_KEY` 只能放在服务端环境变量里，不要写进前端页面或公开仓库。

## Hugging Face Spaces 部署

这个项目也可以部署到 Hugging Face Spaces，适合没有信用卡、只想给别人一个公开访问链接的场景。

使用 Docker SDK 时，Spaces 会运行 `Dockerfile` 中的启动命令。
