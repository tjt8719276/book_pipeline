---
name: book_pipeline
description: Process ASR course transcripts into cleaned chapter manuscripts and build printable MD, DOCX, and PDF outputs.
metadata:
  short-description: Turn cleaned ASR transcripts into book-ready chapter files
---

# Book Pipeline

将语音转写（ASR）课程转录稿处理为可打印的清爽书稿（DOCX + PDF + MD）。

## 触发方式

用户说“处理转录稿”、“处理XX章”、“把XX转录稿做成书稿”等类似表达时，按本流程执行。

## 前置条件

- 源目录下已有 `clean/*.md` 文件（ASR 预处理后的口语转录稿）
- Python 环境已安装 `python-docx` 和 `pywin32`
- Word 已安装（用于 DOCX → PDF 转换）

## 完整工作流程

### 第一步：了解输入

向用户确认（如果信息不全）：
1. **源目录**：包含 `clean/*.md` 的章节目录
2. **章节编号和标题**：如 `第七章 细节化客户开发`
3. **输出目录**：默认 `d:\desk\new`

### 第二步：AI 处理

1. 列出所有 `clean` 文件：`ls source_dir/clean/*.md | sort`
2. 批量阅读：每次读 4-6 个文件，理解内容
3. 主题分组：将相关内容合并为 4-8 个逻辑书章
4. 逐章重写：写入 `source_dir/output/clean_chapters/01_xxx.md` 等

重写规则：

- 删除：口播废话、寒暄、重复、“大家好/谢谢大家/下节课再见”、ASR 转写噪音或乱码
- 改写：口语转书面语，长句拆分或重组，段落逻辑清晰
- 保留：核心观点、案例、方法、步骤、模板（含英文邮件/私信范例）
- 格式：`# Chapter Title` / `## Section` / `### Subsection`
- 强调：用 `**粗体**` 标注关键术语和概念
- 不硬控字数，不在正文中加处理说明

英文内容处理规则：

所有英文邮件模板、私信范例、LinkedIn Headline 案例等，保留英文原文，并在下方添加中文翻译：

```text
> This is a sample English email template...

> **中文译文：** 这是一个英文邮件模板示例...
```

如果某章节纯中文、无英文模板，无需强加翻译。

### 第三步：机械构建

`clean_chapters` 全部写完后，运行：

```bash
python d:/desk/book_pipeline/merge_and_build.py "<source_dir>" "<章节标题>" "<输出文件名>" ["<输出目录>"]
```

示例：

```bash
python d:/desk/book_pipeline/merge_and_build.py \
  "E:/01.毅冰业务课2025（新版）/_book_transcripts/5.五_改变思维定势_funasr" \
  "第五章 改变思维定势" \
  "第五章_改变思维定势" \
  "d:/desk/new"
```

脚本自动执行：

1. 合并 `clean_chapters/0*.md` 到 `merged_chapter.md`，并添加主标题 `# 章节标题`
2. 调用 `build_docx.py` 生成 DOCX
3. 通过 Word COM 转为 PDF
4. 复制 MD 到输出目录
5. 打印每个文件的路径和大小

### 第四步：验证

1. 确认输出目录下存在 `.md`、`.docx`、`.pdf` 三个文件
2. DOCX 在 Word 中打开，确认样式正确
3. PDF 打开确认 A4 排版

## DOCX 输出样式

| 样式 | 字号 | 字体 | 颜色 |
|------|------|------|------|
| 总标题 | 18pt | 微软雅黑 | 黑色，居中 |
| H1 章节标题 | 14pt | 微软雅黑 | #1F4E79（深蓝） |
| H2 小节标题 | 11.5pt | 微软雅黑 | #2E75B6（中蓝） |
| H3 子标题 | 10pt | 微软雅黑 | 黑色加粗 |
| 正文 | 10pt | 微软雅黑 | 黑色 |
| 英文模板/代码 | 9pt | Consolas | 黑色，浅灰背景 |
| 中文译文 | 10pt | 微软雅黑 | 黑色，缩进 |

- 行距：1.5
- 页面：A4（21cm × 29.7cm），上下左右 2.5cm 边距

## 处理大型章节

如果 `clean/*.md` 文件超过 30 个，使用后台 Agent 并行处理：

```text
Agent({
  description: "Process Chapter X clean files",
  subagent_type: "general-purpose",
  run_in_background: true,
  prompt: """
  Process Chapter X ...
  Source directory: E:\...\X_funasr\
  Output: E:\...\X_funasr\output\clean_chapters\
  [详细的处理规则，同上述重写规则]
  """
})
```

Agent 完成后，执行第三步合并和构建。

## 处理规则速查

### 删除内容

- “各位朋友大家好”、“欢迎大家来到XX”
- “谢谢大家宝贵时间”、“我们下节课再见”
- “因为时间关系”、“这里就不多说了”
- 口头禅：“对不对”、“是不是”、“就是说”、“那么”、“然后呢”
- 重复表达：同一个意思说两遍以上的
- ASR 转写乱码：无法理解的字符序列
- PPT 提示语：“大家看到的”、“我写在PPT上的”

### 保留内容

- 核心观点和理论框架
- 案例故事和数据
- 操作方法和步骤
- 英文邮件/私信/Headline 模板（必须加中文翻译）
- 对比分析（如三种邮件写法的对比）
- 具体的工具推荐和网址

### 改写原则

- 口语转书面语（“哎呦”删除，“对吧”删除）
- 长句拆分（一句话超过 80 字时拆分）
- 段落有逻辑主题（一段讲一件事）
- 保持毅冰第一人称讲课风格（用“我”而非“毅冰老师”）
- 章节标题简洁有力（4-10 个字）

## 输出文件清单

每章处理后输出 3 个文件：

```text
{output_dir}\
  {输出文件名}.md
  {输出文件名}.docx
  {输出文件名}.pdf
```
