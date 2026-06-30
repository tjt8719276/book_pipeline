---
name: book_pipeline
description: Process ASR course transcripts into cleaned chapter manuscripts and build printable MD, DOCX, and PDF outputs.
metadata:
  short-description: Turn cleaned ASR transcripts into book-ready chapter files
---

# Book Pipeline

将语音转写（ASR）课程转录稿处理为可打印的清爽书稿（DOCX + PDF + MD）。

## 触发方式

用户说"处理转录稿"、"处理XX章"、"把XX转录稿做成书稿"等类似表达时，按本流程执行。

## 前置条件

- 源目录下已有独立课程文件（`clean/*.md` 或拆分后的单课文件）
- Python 环境已安装 `python-docx` 和 `pywin32`
- Word 已安装（用于 DOCX → PDF 转换）

---

## 项目规模判断（首先执行）

处理前先判断项目规模：

| 条件 | 判定 | 流程 |
|------|------|------|
| 源文件 ≤ 30 且总中文字数 < 20 万 | **小项目** | 标准流程 |
| 源文件 > 30 或总中文字数 ≥ 20 万 | **大项目** | 多 Agent 并行流程 |

判断方法：
```bash
# 统计源文件数量
ls source_dir/clean/*.md | wc -l
# 统计总中文字数
python -c "import re; text=open('merged_file.txt').read(); print(len(re.findall(r'[一-鿿]', text)))"
```

---

## 标准流程（小项目，≤30 课，<20 万汉字）

小项目使用单 Agent 串行处理，简单直接。

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

### 第三步：机械构建

`clean_chapters` 全部写完后，运行：

```bash
python d:/desk/book_pipeline/merge_and_build.py "<source_dir>" "<章节标题>" "<输出文件名>" ["<输出目录>"]
```

### 第四步：验证

1. 确认输出目录下存在 `.md`、`.docx`、`.pdf` 三个文件
2. DOCX 在 Word 中打开，确认样式正确
3. PDF 打开确认 A4 排版

---

## 大项目流程（>30 课 或 ≥20 万汉字）

大项目核心原则：**禁止单 Agent 串行读完所有课再写——必须经过"并行提取→笔记落盘→并行组装"三层架构。** 这是 v8 验证通过的方案（176 课，98,539 汉字，27.6% 保留率）。

### 阶段 0：拆分（如源文件为合并大文件）

如果源文件是单一合并文件（如 `merged/chapter_merged.txt`），先用 Python 拆分为独立课程文件：

```python
import re
from pathlib import Path

text = Path('merged.txt').read_text(encoding='utf-8')
lines = text.split('\n')
lessons = []
current_num, current_lines = None, []
for line in lines:
    m = re.match(r'^(\d+)\s+第\d+课', line)
    if m:
        if current_num is not None and current_lines:
            lessons.append((current_num, '\n'.join(current_lines)))
        current_num = int(m.group(1))
        current_lines = [line]
    else:
        if current_num is not None:
            current_lines.append(line)
if current_num is not None and current_lines:
    lessons.append((current_num, '\n'.join(current_lines)))

out_dir = Path('output/lessons')
out_dir.mkdir(parents=True, exist_ok=True)
for num, content in lessons:
    (out_dir / f'{num:03d}_lesson.txt').write_text(content, encoding='utf-8')
```

如果课程文件已是独立文件（如 `clean/1_第1课...md`），跳过此阶段。

### 阶段 1：课程→章节映射

**目的**：建立每节课归属哪个章节的映射表。

**做法**：
1. 列出全部课程文件的标题（`ls clean/*.md | sed ...`）
2. 对照目标章节结构，逐课归属
3. 产出 `output/lesson_chapter_map.json`：

```json
{
  "project": "项目名",
  "total_lessons": 176,
  "chapters": {
    "01_章节名": {
      "title": "第1章 XXX",
      "description": "本章涵盖的主题概述",
      "lessons": [1, 2, 3, ...]
    },
    ...
  }
}
```

**注意**：一章可引用多课，一课可归属多章（如有需要）。优先单归属。

### 阶段 2：并行提取结构化笔记 ★ 核心

**原则**：每个 Agent 只处理 1-2 节课。上下文小 → 提取质量高。笔记落盘 → 永久复用。

#### 2.1 提取格式（五维度）

每课提取如下格式的结构化笔记：

```markdown
# 第X课笔记：课程标题

## 知识点
- 核心概念、理论框架（去掉口语化表达，提炼为要点列表）

## 方法步骤
- 操作流程、步骤拆解（保持顺序，保留所有实操细节）

## 英文模板/话术
- 原文完整保留 + 中文翻译
- 包括：邮件模板、电话话术、谈判句型、LinkedIn Headline 等
- 如果本节有提到模板但未给出完整文本，注明讨论了哪些模板

## 案例/故事
- 课程中讲述的真实经历、客户案例（保留人名/公司名/数据/结果）

## 工具/网址
- 推荐的软件名称、插件名、网站 URL
```

> 提取笔记**不是摘要**。是**把口语翻译成结构化笔记**，保留所有实质性细节。

#### 2.2 批次规划

以 ~35 课为一批，每批一个 Workflow。每批内每 2 课配一个 Agent：

| 批次 | 课程范围 | Agent 数 |
|------|---------|----------|
| Batch 1 | 1-35 | ~18 |
| Batch 2 | 36-70 | ~18 |
| ... | ... | ... |
| Batch N | 最后一批 | ~N/2 |

#### 2.3 Workflow 脚本模板

```javascript
export const meta = {
  name: 'extract-batch-N',
  description: 'Extract 5-dimension notes from lessons X-Y',
  phases: [{ title: 'Extract', detail: 'Each agent processes 2 lessons' }],
}

const CLEAN = '<项目根目录>/clean'
const NOTES = '<项目根目录>/output/lesson_notes'

const pairs = [
  [1,2], [3,4], [5,6], /* ... */
]

phase('Extract')

const results = await parallel(
  pairs.map(([a, b]) => () =>
    agent(
      `Extract structured notes from ASR Chinese trade course lessons ${a} and ${b}.

## Task
1. Find lesson files in "${CLEAN}/" using Glob pattern "{number}_*.md"
2. Read each lesson
3. Extract into 5 dimensions and write to ${NOTES}/

## Output files
- ${NOTES}/${String(a).padStart(3,'0')}_notes.md
- ${NOTES}/${String(b).padStart(3,'0')}_notes.md

## 5 Dimensions

### 知识点
Core concepts, theories as bullet points. Remove filler (对不对/就是说/大家注意一下).

### 方法步骤
Step-by-step procedures. Preserve ALL sequence and practical details.

### 英文模板/话术
English email templates, phone scripts, negotiation phrases, LinkedIn headlines etc.
Keep original English + add 中文译文 below each.

### 案例/故事
Real customer cases and experiences. Keep names, data, specific results.

### 工具/网址
Software, plugins, website URLs mentioned.

## CRITICAL
- NOT a summary — capture ALL substantive details from the oral transcript
- Remove ONLY: greetings, sign-offs, filler words, repeated phrases, ASR gibberish,
  PPT instructions ("大家看到的", "我写在PPT上的"), screen-clicking descriptions
- If a dimension has no content, write "（本节无此内容）"
- Format each file EXACTLY:

\`\`\`markdown
# 第X课笔记：{lesson title}

## 知识点
- ...

## 方法步骤
- ...

## 英文模板/话术
- ...

## 案例/故事
- ...

## 工具/网址
- ...
\`\`\`

Process both lessons ${a} and ${b}. Write each to its own _notes.md file.`,
      { label: `lessons-${a}-${b}` }
    )
  )
)

const done = results.filter(Boolean).length
log(`${done}/${pairs.length} pairs done`)
return { processed: done, total: pairs.length }
```

#### 2.4 质量检查

每批次完成后，抽样检查 3-5 课笔记：
- 五个维度是否都有内容（无则写"本节无此内容"）
- 英文模板是否保留了原文+翻译
- 工具/网址是否完整列出
- 案例是否保留了具体细节

### 阶段 3：并行组装章节

**原则**：写章节的 Agent **不读 ASR 原文**，只读阶段 2 产出的结构化笔记。笔记已是精华，写作效率极高。

#### 3.1 Workflow 脚本模板（10 章并行）

```javascript
export const meta = {
  name: 'assemble-chapters',
  description: 'Assemble N book chapters from lesson notes in parallel',
  phases: [{ title: 'Assemble', detail: 'One agent per chapter' }],
}

const NOTES = '<项目根目录>/output/lesson_notes'
const CHAPTERS = '<项目根目录>/output/clean_chapters'

const chapterDefs = [
  {
    id: '01', title: '章节名',
    desc: '本章主题概述',
    lessons: [1, 2, 3, ...]
  },
  // ... 全部章节
]

phase('Assemble')

function buildPrompt(ch) {
  const noteFiles = ch.lessons.map(function(n) {
    return NOTES + '/' + String(n).padStart(3, '0') + '_notes.md'
  }).join('\n')

  const outputPath = CHAPTERS + '/' + ch.id + '_' + ch.title + '.md'

  return 'You are writing a chapter for a Chinese reference book.\n\n' +
    '## Chapter: 第' + ch.id + '章 ' + ch.title + '\n' +
    '## Theme: ' + ch.desc + '\n\n' +
    '## Source Material\n' +
    'Read ALL the following lesson notes files:\n\n' +
    noteFiles + '\n\n' +
    '## Task\n' +
    'Read all lesson notes above, then write ONE cohesive chapter synthesizing ALL content.\n\n' +
    '## WRITING RULES (MUST FOLLOW)\n' +
    '1. DELETE: greetings, sign-offs, filler words (对不对/就是说/那么/然后呢), ' +
    'repeated expressions, ASR gibberish, PPT instructions, screen-clicking descriptions\n' +
    '2. PRESERVE: core concepts/theories, methods/steps with details, case stories with data, ' +
    'ALL English templates (add Chinese translation below each), tool names and URLs, risk warnings\n' +
    '3. REWRITE: oral to written Chinese, split sentences over 80 chars, ' +
    'one logical topic per paragraph, section headings concise (4-10 chars)\n' +
    '4. ENGLISH: Keep original English text, add **中文译文：** below each English block\n' +
    '5. STRUCTURE: Use ## for major sections, ### for subsections. ' +
    'Organize by topic (not lesson order). Merge same-topic content from multiple lessons.\n' +
    '6. TARGET: 8,000-12,000 pure Chinese characters (exclude English, markdown syntax, spaces)\n' +
    '7. STYLE: Professional reference book. No "这节课" or "讲师说" — write as cohesive knowledge.\n' +
    '8. OUTPUT PATH: ' + outputPath + '\n\n' +
    '## IMPORTANT\n' +
    '- Read EVERY lesson note file before writing\n' +
    '- Cross-reference across lessons — merge same-topic content into one section\n' +
    '- Preserve ALL English templates, email examples, phone scripts (most valuable content)\n' +
    '- Preserve ALL tool names, URLs, software names\n' +
    '- Include specific data, numbers, steps from notes\n' +
    '- Write directly to the output path using Write tool'
}

const tasks = chapterDefs.map(function(ch) {
  return function() {
    return agent(buildPrompt(ch), { label: 'ch-' + ch.id + '-' + ch.title })
  }
})

const results = await parallel(tasks)
const done = results.filter(Boolean).length
log(done + '/' + chapterDefs.length + ' chapters written')
return { written: done, total: chapterDefs.length }
```

#### 3.2 章节写作约束

- **删除**：口语问候、口头禅、重复表达、ASR 乱码、PPT 提示语、纯操作演示描述
- **保留**：核心观点与方法、案例数据、英文模板（加中文翻译）、工具网址、风险提示
- **改写**：口语转书面语、长句拆分（≤80 字）、段落有逻辑主题
- **英文模板**：保留原文 + 下方加 `**中文译文：**`
- **目标密度**：每章 8,000-12,000 纯中文汉字

### 阶段 4：统稿与构建

#### 4.1 字数和覆盖验证

```bash
# 统计纯中文汉字数
python -c "
import re, os
d = 'output/clean_chapters'
total = 0
for f in sorted(os.listdir(d)):
    if f.endswith('.md'):
        text = open(os.path.join(d, f), encoding='utf-8').read()
        cn = len(re.findall(r'[一-鿿]', text))
        total += cn
        print(f'{f}: {cn:,} 汉字')
print(f'总计: {total:,} 汉字')
print(f'保留率: {total/原始总汉字数*100:.1f}%')
"
```

- 确认总字数 ≥ 80,000（对应 ~22% 保留率，目标 20-28%）
- 确认 `clean_chapters` 目录为最新版本
- 对照 `lesson_chapter_map.json` 确认覆盖

#### 4.2 构建

```bash
python d:/desk/book_pipeline/merge_and_build.py "<source_dir>" "<书名>" "<输出文件名>" ["<输出目录>"]
```

#### 4.3 对比验证

- 新旧文件大小对比，确认不是读了旧目录
- DOCX 在 Word 中打开确认中文字体（微软雅黑）正确
- PDF 打开确认 A4 排版

---

## 通用重写规则

### 删除内容

- "各位朋友大家好"、"欢迎大家来到XX"
- "谢谢大家宝贵时间"、"我们下节课再见"
- "因为时间关系"、"这里就不多说了"
- 口头禅："对不对"、"是不是"、"就是说"、"那么"、"然后呢"
- 重复表达：同一个意思说两遍以上的
- ASR 转写乱码：无法理解的字符序列
- PPT 提示语："大家看到的"、"我写在PPT上的"
- 纯操作演示描述："我现在点这里...然后再点这里..."

### 保留内容

- 核心观点和理论框架
- 案例故事和数据
- 操作方法和步骤
- 英文邮件/私信/Headline 模板（**必须加中文翻译**）
- 对比分析（如三种邮件写法的对比）
- 具体的工具推荐和网址
- 风险提示和注意事项

### 改写原则

- 口语转书面语（"哎呦"删除，"对吧"删除）
- 长句拆分（一句话超过 80 字时拆分）
- 段落有逻辑主题（一段讲一件事）
- 章节标题简洁有力（4-10 个字）
- 不在正文中加处理说明

### 英文内容处理规则

所有英文邮件模板、私信范例、LinkedIn Headline 案例等，保留英文原文，并在下方添加中文翻译：

```text
> This is a sample English email template...

> **中文译文：** 这是一个英文邮件模板示例...
```

如果某章节纯中文、无英文模板，无需强加翻译。

---

## DOCX 输出样式

| 样式 | 字号 | 字体 | 颜色 |
|------|------|------|------|
| 总标题 | 22pt | 微软雅黑 | 黑色，居中 |
| 副标题 | 14pt | 微软雅黑 | #1F4E79（深蓝），居中 |
| H1 章节标题 | 14pt | 微软雅黑 | #1F4E79（深蓝） |
| H2 小节标题 | 11.5pt | 微软雅黑 | #2E75B6（中蓝） |
| H3 子标题 | 10pt | 微软雅黑 | 黑色加粗 |
| 正文 | 10pt | 微软雅黑 | 黑色 |
| 英文模板/代码 | 9pt | 微软雅黑 | 黑色，浅灰背景 |
| 中文译文 | 10pt | 微软雅黑 | 黑色，缩进 |

- 行距：1.0
- 页面：A4（21cm × 29.7cm），上下左右 2.5cm 边距

---

## 大项目 vs 小项目对比

| | 小项目流程 | 大项目流程 |
|---|---|---|
| 适用 | ≤30 课，<20 万汉字 | >30 课 或 ≥20 万汉字 |
| 提取方式 | Agent 直接读→写章节 | 多 Agent 并行提取笔记→落盘→并行组装 |
| 中间产物 | 无 | 每课一个 `_notes.md` 文件 |
| 并行度 | 串行 | 提取阶段 ~10 Agent 并行，组装阶段全章并行 |
| 失败恢复 | 对话丢了重来 | 只重跑失败单课 |
| 迭代能力 | 改一章需重读相关原文 | 直接基于笔记文件修改 |

---

## 输出文件清单

大项目处理完成后产出：

```text
output/
  lesson_chapter_map.json      — 课程→章节映射表
  lesson_notes/                 — 结构化笔记（每课一个文件）
    001_notes.md
    002_notes.md
    ...
  clean_chapters/               — 正式章节（10 章）
    01_章节名.md
    02_章节名.md
    ...
  merged_chapter.md             — 合并后的完整书稿

{输出目录}/
  {书名}.md                     — 最终 MD
  {书名}.docx                   — 最终 DOCX
  {书名}.pdf                    — 最终 PDF
```

---

## 同步提醒

本 skill 有两份副本，修改时必须同时更新：
- `~/.claude/skills/book_pipeline/`（Claude Code 加载位置）
- `d:/desk/book_pipeline/`（git 仓库位置）
