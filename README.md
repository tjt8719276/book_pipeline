# Book Pipeline

将语音转写（ASR）课程转录稿处理为可打印的清爽书稿（DOCX + PDF + MD）。

## 两种工作模式

| | 小项目模式 | 大项目模式 |
|---|---|---|
| **适用** | ≤30 课，<20 万汉字 | >30 课 或 ≥20 万汉字 |
| **处理方式** | 单 Agent 串行：读→写章节 | 多 Agent 并行：提取笔记→落盘→并行组装章节 |
| **并行度** | 无 | 提取 ~10 Agent 并行，组装 10 章并行 |
| **中间产物** | 无 | 每课一个结构化笔记文件（可复用） |

## 快速开始

### 1. 作为 Claude Code Skill

在 Claude Code 中说 `处理转录稿` 即可触发。Skill 会自动判断项目规模并选择合适的流程。

Skill 文件：`SKILL.md`

### 2. Python 脚本独立使用

当 `clean_chapters/` 下已写好章节文件后：

```bash
python merge_and_build.py "source_dir" "书名" "输出文件名" "输出目录"
```

自动完成：合并章节 → 生成 DOCX → 转 PDF → 复制 MD。

### 3. 完整大项目流程（>30 课）

```
阶段 0: 拆分课程文件（如为单一合并文件）
    ↓
阶段 1: 建立课程→章节映射表（lesson_chapter_map.json）
    ↓
阶段 2: 多 Agent 并行提取五维度结构化笔记 → output/lesson_notes/
    ↓
阶段 3: 多 Agent 并行基于笔记组装章节 → output/clean_chapters/
    ↓
阶段 4: 字数验证 + 合并构建 → MD + DOCX + PDF
```

## 真实案例：v8 外贸课程书稿

- **源材料**：176 节 ASR 口语转录课程，~35.7 万汉字
- **处理方式**：大项目多 Agent 并行流程
- **Agent 调用**：88 Agent 并行提取 + 10 Agent 并行组装 = 98 次调用，总计约 15 分钟
- **产出**：98,539 纯中文汉字，27.6% 保留率（目标 20-28%）
- **对比上版**：v7（单 Agent 串行）只做到 25,398 汉字（7.1%）— **提升 3.9 倍**

## 五维度提取格式

每节课提取为结构化笔记：

```markdown
# 第X课笔记：课程标题

## 知识点
- 核心概念、理论框架

## 方法步骤
- 操作流程、步骤拆解

## 英文模板/话术
- 原文 + 中文翻译

## 案例/故事
- 真实经历、客户案例

## 工具/网址
- 软件、插件、网站
```

## 输出样式

| 样式 | 字号 | 字体 |
|------|------|------|
| 总标题 | 22pt | 微软雅黑 |
| H1 章节标题 | 14pt | 微软雅黑（深蓝 #1F4E79） |
| H2 小节标题 | 11.5pt | 微软雅黑（中蓝 #2E75B6） |
| 正文 | 10pt | 微软雅黑 |
| 英文模板 | 9pt | 微软雅黑（浅灰背景） |

- 行距：1.0
- 页面：A4，上下左右 2.5cm 边距

## 前置依赖

- Python 3.x + `python-docx`
- Microsoft Word（DOCX → PDF 转换）

## 文件结构

```
book_pipeline/
  SKILL.md              — Claude Code Skill 定义
  merge_and_build.py    — 章节合并 + 构建流水线
  build_docx.py         — Markdown → DOCX 转换
  run_book_pipeline.py  — 便捷启动脚本
  README.md             — 本文件
```

## License

Private use.
