# 本科论文格式检测 Web 版

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

服务不需要 Supabase。Supabase 适合数据库、登录和文件存储；这个工具只需要临时接收 Word 文件并生成报告。
