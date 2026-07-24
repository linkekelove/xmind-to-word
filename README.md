# XMind to Word Converter

## 简介

一个将 XMind 思维导图（`.xmind`）转换为 Word 文档（`.docx`）的 Python 脚本。

**功能特点：**
- 保留思维导图原有的父子节点层级关系（对应 Word 标题1~9级）
- 自动嵌入思维导图中的配图
- 按中文排版习惯设置字体格式（标题1~4黑体加粗分级、正文宋体，左对齐、首行缩进2字符、正文段后间距等）
- 自动清理全角空格、不间断空格、零宽字符等不可见排版污染字符
- 标题层级自动编号（二级中文数字，三/四级多级阿拉伯数字；唯一的一级标题自动跳过编号）

## 使用方法

```bash
pip install -r requirements.txt
python xmind_to_word.py 输入文件.xmind 输出文件.docx
```

### 常见问题：提示"找不到文件"（No such file or directory）

如果终端所在目录不是本脚本所在目录，`pip install -r requirements.txt` 和 `python xmind_to_word.py ...` 都会报“找不到文件”，因为它们默认在**当前工作目录**里查找文件。

**命令行 / PyCharm 终端：**
```bash
cd 本仓库所在目录          # 例如 cd xmind-to-word
pip install -r requirements.txt
python xmind_to_word.py 输入文件.xmind 输出文件.docx
```
输入、输出文件路径也要写成相对于该目录的路径（或直接用绝对路径），否则同样会报错。

**PyCharm 运行配置（Run Configuration）：**
如果是通过 PyCharm 的运行按钮而不是终端启动脚本，需要把 Run Configuration 里的 **Working directory** 设置为本脚本所在目录，而不是整个工作区/项目的根目录，否则会出现一样的报错。

## 依赖

- Python 3
- python-docx

更新日志见 [CHANGELOG.md](CHANGELOG.md)。

---

## Introduction

A Python script that converts XMind mind maps (`.xmind`) into Word documents (`.docx`).

**Features:**
- Preserves the original parent-child node hierarchy (mapped to Word Heading 1~9)
- Automatically embeds images from the mind map
- Applies Chinese typesetting conventions (Heading 1~4 in bold 黑体/Heiti with distinct levels, body text in 宋体/SimSun, left alignment, 2-character first-line indent, paragraph spacing after body text, etc.)
- Strips invisible formatting artifacts such as full-width spaces, non-breaking spaces, and zero-width characters
- Automatic heading numbering (Chinese numerals for Heading 2; multi-level Arabic numerals for Heading 3/4; the sole top-level Heading 1 is skipped)

## Usage

```bash
pip install -r requirements.txt
python xmind_to_word.py input.xmind output.docx
```

### Troubleshooting: "No such file or directory"

If your terminal's current directory isn't the directory containing this script, both `pip install -r requirements.txt` and `python xmind_to_word.py ...` will fail with a "file not found" error, since they look for files relative to the **current working directory**.

**Command line / PyCharm terminal:**
```bash
cd path/to/this/repo          # e.g. cd xmind-to-word
pip install -r requirements.txt
python xmind_to_word.py input.xmind output.docx
```
Input/output file paths must also be relative to that directory (or given as absolute paths), or the same error will occur.

**PyCharm Run Configuration:**
If you launch the script via PyCharm's Run button instead of the terminal, set the Run Configuration's **Working directory** to this script's directory, not the overall workspace/project root, to avoid the same error.

## Requirements

- Python 3
- python-docx

See [CHANGELOG.md](CHANGELOG.md) for release history.


