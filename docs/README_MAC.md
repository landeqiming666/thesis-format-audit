# 本科毕业设计论文格式检测工具 macOS 版

这个工具用于检测 Word `.docx` 论文格式，并生成一个可交互的 HTML 看板报告。

## 运行前准备

Mac mini 上需要安装 Python 3。

如果终端输入下面命令能看到版本号，就说明已经安装：

```bash
python3 --version
```

如果没有安装，可以从 Python 官网下载安装：

https://www.python.org/downloads/macos/

## 最简单的使用方法

1. 解压 `thesis-format-audit-mac.zip`。
2. 进入解压后的文件夹。
3. 双击 `scripts/local/run_mac.command`。
4. 按窗口提示，把要检测的 `.docx` 文件拖入窗口。
5. 按回车。
6. 检测完成后会自动打开 HTML 报告。

报告默认生成在桌面：

```text
~/Desktop/thesis_format_audit_reports/
```

## 如果 macOS 提示无法打开

第一次运行 `.command` 文件时，macOS 可能因为安全策略阻止打开。

可以打开“终端”，进入工具文件夹后运行：

```bash
chmod +x scripts/local/run_mac.command scripts/local/run_mac_terminal.sh
./scripts/local/run_mac.command
```

## 终端命令用法

也可以直接指定文件路径运行：

```bash
./scripts/local/run_mac_terminal.sh "/path/to/论文.docx"
```

指定输出报告路径：

```bash
./scripts/local/run_mac_terminal.sh "/path/to/论文.docx" "/path/to/report.html"
```

## 注意事项

- 工具只读取 `.docx`，不会修改原文件。
- 检测报告是 HTML 文件，可以直接发给别人。
- 看板里的勾选记录保存在 HTML 自身的浏览器本地状态中；如果要保留勾选状态，可以使用页面中的导出功能。
- 第一次运行会自动创建 `.venv` 虚拟环境并安装依赖，需要联网。
- 如果 Word 文档是 `.doc`，请先在 Word 中另存为 `.docx`。
