# XMind to Word Converter

## 简介

一个将 XMind 思维导图（`.xmind`）转换为 Word 文档（`.docx`）的 Python 脚本。

**功能特点：**
- 保留思维导图原有的父子节点层级关系（对应 Word 标题1~9级）
- 自动嵌入思维导图中的配图
- 按中文排版习惯设置字体格式（标题黑体、正文宋体，两端对齐、首行缩进2字符等）

## 使用方法

```bash
pip install -r requirements.txt
python xmind_to_word.py 输入文件.xmind 输出文件.docx
```

## 依赖

- Python 3
- python-docx

---

## Introduction

A Python script that converts XMind mind maps (`.xmind`) into Word documents (`.docx`).

**Features:**
- Preserves the original parent-child node hierarchy (mapped to Word Heading 1~9)
- Automatically embeds images from the mind map
- Applies Chinese typesetting conventions (headings in 黑体/Heiti, body text in 宋体/SimSun, justified alignment, 2-character first-line indent, etc.)

## Usage

```bash
pip install -r requirements.txt
python xmind_to_word.py input.xmind output.docx
```

## Requirements

- Python 3
- python-docx


