---
title: UPC本科论文格式检测工具
emoji: 📄
colorFrom: green
colorTo: gray
sdk: docker
pinned: false
---

# UPC本科论文格式检测工具

面向中国石油大学本科毕业论文的在线格式检测服务。用户上传 `.docx` 或旧版 `.doc` 文档后，系统会运行本项目的 Word 格式检查脚本，生成可下载、可展开技术详情的 HTML 检测报告。

项目不是一个单纯脚本，而是一套完整的小型 Web 服务：包含用户注册登录、邮箱验证码、QQ群注册码、检测次数、管理员后台、报告归档、学院统计看板、Supabase 数据库与私有存储，以及 Google Drive / Google Cloud Storage 可选备份。

## 功能概览

- 论文格式检测：调用 `thesis_format_audit.py` 检查封面、摘要、目录、章节标题、正文、图表、公式编号、参考文献、致谢等格式问题，并尽量贴近维普格式检测的分类和表达。
- DOC 兼容：支持上传 `.docx`，也支持旧版 `.doc`；`.doc` 会先通过 LibreOffice 转换为临时 `.docx` 再检测。
- 在线报告：生成 HTML 报告，支持问题概览、异常详情、修改建议、技术详情、下载报告和 GitHub Star 提示。
- 用户体系：邮箱密码注册、邮箱验证码提示、QQ群注册码、账号状态、默认检测次数、剩余次数展示。
- 管理后台：支持查看用户、增减次数、冻结或注销账号、管理注册码、查看检测记录、学院归类和访问统计。
- 存储归档：检测报告和原文可写入 Supabase Storage，并可迁移到 Google Drive 或 Google Cloud Storage，配套清理脚本控制 Supabase 容量。

## 项目架构

```text
浏览器
  |
  | 上传论文、登录注册、查看报告
  v
Flask Web 服务 app.py
  |
  |-- thesis_format_audit.py
  |     检测 Word 结构与样式，输出 HTML 报告
  |
  |-- Supabase Postgres
  |     用户、次数、注册码、检测记录、访问事件、管理员日志
  |
  |-- Supabase Storage
  |     私有保存检测报告和论文原文归档
  |
  |-- Gmail SMTP
  |     注册验证码和运营提醒邮件
  |
  |-- Google Drive / Google Cloud Storage
        可选二级归档，用于释放 Supabase 存储空间
```

## 目录结构

根目录只保留应用入口、核心检测脚本、依赖和 Dockerfile；部署模板、维护工具、数据库脚本和补充文档都收进子目录。

```text
.
├── app.py                         # Flask Web、API、认证、后台、存储、邮件、页面模板
├── thesis_format_audit.py          # Word 论文格式检测核心脚本和 HTML 报告生成器
├── config.py                       # 环境变量与运行参数集中配置
├── requirements.txt                # Python 依赖
├── uv.toml                         # uv 国内 PyPI 镜像配置
├── Dockerfile                      # Docker / Hugging Face Spaces / Azure 容器镜像
├── deploy/
│   ├── docker-compose.yml          # Mac mini 或本地 Docker 自托管
│   └── render.yaml                 # Render Blueprint 部署配置
├── supabase/
│   └── schema.sql                  # Supabase 表、索引、RPC、策略初始化脚本
├── scripts/
│   ├── maintenance/                # 管理、迁移、清理、回填、邮件等维护脚本
│   └── local/                      # macOS 离线版双击和命令行入口
└── docs/
    ├── env.example                 # 环境变量模板
    ├── MAC_MINI_DEPLOY.md          # Mac mini + Docker 自托管部署说明
    └── README_MAC.md               # 离线 macOS 版使用说明
```

## 本地开发

项目推荐使用 `uv` 创建虚拟环境，根目录的 `uv.toml` 已配置清华 PyPI 镜像。

```bash
uv venv .venv
source .venv/bin/activate
uv pip install -r requirements.txt
python app.py
```

打开：

```text
http://127.0.0.1:8000
```

如果不用 `uv`，也可以使用普通 `pip`：

```bash
python -m venv .venv
source .venv/bin/activate
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt
python app.py
```

检测 `.doc` 文件需要系统里安装 LibreOffice。Docker 镜像会自动安装 `libreoffice-writer` 和中文字体；本机直接运行时需要自己安装。

## 环境变量

可以复制 `docs/env.example` 为 `.env`，再按自己的服务配置填写。

| 分类 | 变量 | 说明 |
| --- | --- | --- |
| 基础 | `SECRET_KEY` | Flask 签名密钥，生产环境必须使用随机长字符串 |
| Supabase | `SUPABASE_URL` | Supabase Project URL |
| Supabase | `SUPABASE_SERVICE_ROLE_KEY` | 服务端专用 service role key，不能暴露到前端或提交到仓库 |
| Supabase | `REPORTS_BUCKET` | 私有报告 bucket，默认 `thesis-audit-reports` |
| 权限 | `SUPER_ADMIN_EMAILS` | 最高管理员邮箱，逗号分隔 |
| 权限 | `ADMIN_EMAILS` | 普通管理员邮箱，兼容旧配置，可选 |
| 用户次数 | `MAX_SUBMISSIONS` | 新用户默认检测次数，当前默认 `100` |
| 上传限制 | `MAX_UPLOAD_MB` | 单个上传文件大小限制 |
| 上传限制 | `MAX_DOCX_ENTRIES` | 防止恶意 docx zip 炸弹的最大条目数 |
| 上传限制 | `MAX_DOCX_UNCOMPRESSED_MB` | docx 解压后最大体积 |
| 超时 | `AUDIT_TIMEOUT_SECONDS` | 单次检测超时时间 |
| 超时 | `DOC_CONVERT_TIMEOUT_SECONDS` | `.doc` 转 `.docx` 超时时间 |
| 邮件 | `GMAIL_SMTP_USER` | Gmail 发件邮箱 |
| 邮件 | `GMAIL_SMTP_APP_PASSWORD` | Gmail App Password，不是账号登录密码 |
| 邮件 | `GMAIL_SMTP_HOST` | 默认 `smtp.gmail.com` |
| 邮件 | `GMAIL_SMTP_PORT` | 默认 `465` |
| 邮件 | `EMAIL_FROM_NAME` | 邮件显示名称 |
| GitHub | `GITHUB_REPO_URL` | 报告弹窗里引导 Star 的仓库地址 |
| Google Drive | `GOOGLE_DRIVE_CREDENTIALS_JSON` | 可选，服务端 Drive 凭据 JSON |
| Google Drive | `GOOGLE_DRIVE_FOLDER_ID` | 可选，Drive 归档目录 ID |
| Google Drive | `GOOGLE_DRIVE_PREFIX` | 可选，Drive 归档路径前缀 |
| GCS | `GCS_BUCKET` | 可选，Google Cloud Storage bucket |
| GCS | `GCS_PREFIX` | 可选，GCS 路径前缀 |
| GCS | `GCS_PROJECT` | 可选，Google Cloud project id |
| GCS | `GOOGLE_APPLICATION_CREDENTIALS_JSON` | 可选，GCS 服务账号 JSON |

## Supabase 初始化

注册登录、检测次数、注册码、检测记录、学院统计、访问事件都依赖 Supabase。

1. 在 Supabase SQL Editor 执行 [supabase/schema.sql](supabase/schema.sql)。
2. 在 Supabase Storage 创建私有 bucket，默认名称为 `thesis-audit-reports`。
3. 在部署平台配置 `SUPABASE_URL`、`SUPABASE_SERVICE_ROLE_KEY`、`SECRET_KEY` 和 `REPORTS_BUCKET`。
4. 确认 service role key 只存在服务端环境变量中，不要写进前端代码、README 示例值或公开仓库。

## 常用维护脚本

所有脚本都建议从项目根目录运行，默认先 dry-run，涉及删除或发信的操作通常需要显式加 `--execute`。

| 命令 | 用途 |
| --- | --- |
| `python scripts/maintenance/check_email_smtp.py 收件邮箱@example.com` | 测试 Gmail SMTP 是否能发验证码 |
| `python scripts/maintenance/grant_low_quota_users.py --help` | 给剩余次数低于阈值的用户补次数，并可邮件提醒 |
| `python scripts/maintenance/backfill_report_colleges.py --help` | 从历史成功检测记录里重新解析学院信息 |
| `python scripts/maintenance/dedupe_duplicate_reports.py --help` | 按账号和文件哈希清理重复检测记录 |
| `python scripts/maintenance/prune_report_archives.py --help` | 控制每个账号最多保留几份检测报告 |
| `python scripts/maintenance/prune_original_archives.py --help` | 控制每个账号原文归档只保留最新记录 |
| `python scripts/maintenance/migrate_archives_to_drive.py --help` | 将 Supabase 中的原文归档迁移到 Google Drive |
| `python scripts/maintenance/maintain_supabase_original_storage.py --help` | Supabase 存储到阈值后，公平清理已迁移的原文文件 |

## 存储与归档策略

默认情况下，系统把检测报告和原文归档放在 Supabase 私有 bucket 中。为了避免免费额度被论文原文占满，可以启用 Google Drive 或 Google Cloud Storage 作为二级归档。

推荐策略：

1. 检测完成后，先写入 Supabase，保证用户能立即下载报告。
2. 定期运行迁移脚本，把原文归档复制到 Google Drive。
3. 当 Supabase Storage 使用量达到阈值时，只删除已经迁移成功的原文，不删除数据库记录和报告记录。
4. 删除采用按账号轮询的公平策略，避免只清理某一个用户的数据。

示例：

```bash
python scripts/maintenance/migrate_archives_to_drive.py --no-reports --execute --client-secret /path/to/client_secret.json
python scripts/maintenance/maintain_supabase_original_storage.py
python scripts/maintenance/maintain_supabase_original_storage.py --execute
```

## 部署

### Docker / Mac mini 自托管

```bash
docker compose -f deploy/docker-compose.yml up -d --build
```

默认监听：

```text
http://127.0.0.1:7860
```

Mac mini 长期自托管和内网穿透参考 [docs/MAC_MINI_DEPLOY.md](docs/MAC_MINI_DEPLOY.md)。

### Azure Container Apps

本项目可以直接构建 Docker 镜像并部署到 Azure Container Apps。部署时需要把 `docs/env.example` 中的生产环境变量配置到 Container App 的环境变量里，尤其是 Supabase、邮件和 `SECRET_KEY`。

### Render

仓库保留了 `deploy/render.yaml`，可以在 Render 创建 Blueprint 后使用：

1. 把仓库推送到 GitHub、GitLab 或 Bitbucket。
2. 在 Render 新建 Blueprint，选择本仓库。
3. Blueprint 文件选择 `deploy/render.yaml`，Render 会启动 Python Web Service。

### Hugging Face Spaces

README 顶部保留了 Hugging Face Spaces 的 YAML front matter。创建 Space 时选择 Docker SDK，平台会使用 `Dockerfile` 中的启动命令。

## 安全说明

- 不要提交 `.env`、Supabase service role key、Google 凭据、Drive token 或任何真实用户论文文件。
- Supabase bucket 应保持私有，下载报告和原文应通过服务端鉴权。
- 管理员接口只应允许已登录管理员访问；最高管理员账号通过 `SUPER_ADMIN_EMAILS` 配置。
- 上传文件可能包含敏感个人信息，清理、迁移和备份前要确认目标存储是私有的。
- `.docx` 本质是 zip 包，项目通过文件大小、条目数、解压体积和检测超时限制降低恶意文件风险。

## 贡献与支持

欢迎提交 Issue 或 Pull Request，一起把检测规则调得更接近真实学校模板和维普报告。
如果这个项目帮到了你，也欢迎给仓库点一个 Star：[landeqiming666/thesis-format-audit](https://github.com/landeqiming666/thesis-format-audit)
