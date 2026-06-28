#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
End-to-end book pipeline:
raw/*.txt -> normalized text -> themed volumes -> clean_chapters -> MD/DOCX/PDF

Usage:
  python run_book_pipeline.py <source_dir> [output_dir] [series_title] [--limit N] [--total-limit N] [--volume XX]
"""

import json
import re
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
MERGE_AND_BUILD = SCRIPT_DIR / "merge_and_build.py"

SERIES_TITLE_DEFAULT = "外贸课程大合集"

VOLUME_RULES = [
    {
        "id": "01",
        "title": "外贸入门与客户开发",
        "keywords": [
            "外贸新人入门",
            "找国外客户",
            "搜索引擎",
            "谷歌",
            "脸书",
            "linkedin",
            "领英",
            "instagram",
            "youtube",
            "gmail",
            "海关数据",
        ],
    },
    {
        "id": "02",
        "title": "询盘回复、报价与邮件沟通",
        "keywords": [
            "客户常问",
            "开发信",
            "询盘",
            "报价单",
            "外贸邮件",
            "电话话术",
            "跟客户打电话",
            "采购邮箱",
            "联系提升回复率",
        ],
    },
    {
        "id": "03",
        "title": "谈判、逼单、成交与客户维护",
        "keywords": [
            "逼单",
            "背景调查",
            "卖点提炼",
            "价格高",
            "谈判",
            "客户说考虑一下",
            "涨价",
            "付款方式",
            "样品谈判",
            "交期谈判",
            "催收尾款",
            "重大投诉",
            "信任",
            "老客户",
            "目标价",
            "接待客户看厂",
        ],
    },
    {
        "id": "04",
        "title": "商务英语与贸易沟通实战",
        "keywords": [
            "基础实用商务英语",
            "基础商务英语",
            "贸易往来",
            "商务谈判",
            "签订合同",
            "合作细则",
        ],
    },
    {
        "id": "05",
        "title": "外贸SOHO实战",
        "keywords": [
            "外贸soho",
            "soho",
            "收款",
            "选品",
            "生存法则",
            "搜索管理",
            "货代流程",
            "快递流程",
            "展会",
            "运输基本",
        ],
    },
    {
        "id": "06",
        "title": "独立站实战",
        "keywords": [
            "独立站",
            "店铺",
            "商城装修",
            "上传产品",
            "选品渠道",
            "商业模式",
            "建站流程",
            "货源和发货",
            "引流渠道",
            "新增页面",
        ],
    },
    {
        "id": "07",
        "title": "TikTok获客与运营",
        "keywords": [
            "tiktok",
            "tk",
            "capcut",
            "创作者基金",
            "paypal",
            "whatsapp",
            "安装环境",
            "短链接",
            "翻译工具",
            "热门",
            "多个链接",
            "直播",
            "搬运",
        ],
    },
]


def slugify(text):
    text = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", text.strip(), flags=re.UNICODE)
    return re.sub(r"_+", "_", text).strip("_")


def source_title(source_dir):
    title = source_dir.name.replace("_funasr", "")
    return title or SERIES_TITLE_DEFAULT


def read_text_file(path):
    for encoding in ("utf-8", "utf-8-sig", "gb18030", "gbk"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="ignore")


def flatten_json_text(path):
    try:
        data = json.loads(read_text_file(path))
    except json.JSONDecodeError:
        return ""
    items = data.get("metadata", {}).get("raw_result", [])
    parts = []
    for item in items:
        if isinstance(item, dict):
            text = item.get("text", "").strip()
            if text:
                parts.append(text)
    return "\n".join(parts)


def flatten_srt_text(path):
    lines = []
    for line in read_text_file(path).splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.isdigit():
            continue
        if "-->" in stripped:
            continue
        lines.append(stripped)
    return "\n".join(lines)


def normalize_text(text):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"(大家好|谢谢大家|下节课再见|拜拜)$", "", text, flags=re.MULTILINE)
    text = re.sub(r"([。！？!?])([^\n])", r"\1\n\2", text)
    text = re.sub(r"(，|；|：)([^ \n])", r"\1\2", text)
    return text.strip()


def build_normalized_inputs(source_dir):
    raw_dir = source_dir / "raw"
    normalized_dir = source_dir / "output" / "normalized_text"
    normalized_dir.mkdir(parents=True, exist_ok=True)

    entries = []
    base_names = sorted({path.stem for path in raw_dir.iterdir() if path.is_file()})
    for base_name in base_names:
        txt_path = raw_dir / f"{base_name}.txt"
        json_path = raw_dir / f"{base_name}.json"
        srt_path = raw_dir / f"{base_name}.srt"

        if txt_path.exists():
            text = read_text_file(txt_path)
            source_kind = "txt"
        elif json_path.exists():
            text = flatten_json_text(json_path)
            source_kind = "json"
        elif srt_path.exists():
            text = flatten_srt_text(srt_path)
            source_kind = "srt"
        else:
            continue

        text = normalize_text(text)
        if not text:
            continue

        normalized_path = normalized_dir / f"{base_name}.txt"
        normalized_path.write_text(text, encoding="utf-8")
        entries.append(
            {
                "base_name": base_name,
                "title": derive_title(base_name),
                "source_kind": source_kind,
                "normalized_path": str(normalized_path),
            }
        )
    return entries


def derive_title(base_name):
    title = re.sub(r"^\d+_", "", base_name)
    title = re.sub(r"_[0-9a-f]{8,}$", "", title, flags=re.IGNORECASE)
    title = re.sub(r"_(720p|1080p).*?$", "", title, flags=re.IGNORECASE)
    title = title.replace("_", " ")
    return title.strip()


def volume_for_title(title):
    lowered = title.lower()
    for rule in VOLUME_RULES:
        for keyword in rule["keywords"]:
            if keyword.lower() in lowered:
                return rule["id"]
    return "01"


def assign_volumes(entries):
    grouped = {rule["id"]: [] for rule in VOLUME_RULES}
    for entry in entries:
        grouped[volume_for_title(entry["title"])].append(entry)
    return grouped


def chunk_sentences(text, size=5):
    raw_sentences = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        raw_sentences.extend([part.strip() for part in re.split(r"(?<=[。！？!?])\s*", line) if part.strip()])

    chunks = []
    current = []
    for sentence in raw_sentences:
        current.append(sentence)
        if len(current) >= size:
            chunks.append("".join(current))
            current = []
    if current:
        chunks.append("".join(current))
    return chunks


def write_clean_chapters(volume_dir, volume_title, entries):
    clean_dir = volume_dir / "clean_chapters"
    clean_dir.mkdir(parents=True, exist_ok=True)

    grouped = {}
    for entry in entries:
        normalized_text = Path(entry["normalized_path"]).read_text(encoding="utf-8")
        chapter_key = chapter_key_for_title(entry["title"])
        grouped.setdefault(chapter_key, []).append((entry["title"], normalized_text))

    items = sorted(grouped.items())
    for idx, (chapter_title, chapter_entries) in enumerate(items, start=1):
        parts = [f"# {chapter_title}", ""]
        for lesson_title, lesson_text in chapter_entries:
            parts.append(f"## {lesson_title}")
            parts.append("")
            paragraphs = chunk_sentences(lesson_text)
            for p_idx, paragraph in enumerate(paragraphs, start=1):
                if p_idx == 1:
                    parts.append("### 核心内容")
                parts.append(paragraph)
                parts.append("")
            parts.append("---")
            parts.append("")
        output_path = clean_dir / f"{idx:02d}_{slugify(chapter_title)}.md"
        output_path.write_text("\n".join(parts).strip() + "\n", encoding="utf-8")


def chapter_key_for_title(title):
    rules = [
        ("入门", ["入门", "全流程", "概述", "基本概念"]),
        ("客户开发", ["客户", "找客户", "开发", "海关数据", "谷歌", "脸书", "领英", "instagram", "youtube"]),
        ("沟通与回复", ["询盘", "报价", "邮件", "电话", "回复"]),
        ("谈判与成交", ["谈判", "逼单", "价格", "付款", "交期", "投诉", "信任", "订单"]),
        ("运营与实操", ["独立站", "tiktok", "tk", "soho", "展会", "货代", "运输", "装修", "引流"]),
    ]
    lowered = title.lower()
    for chapter_title, keywords in rules:
        if any(keyword.lower() in lowered for keyword in keywords):
            return chapter_title
    return "主题整理"


def build_manifest(source_dir, series_title, grouped):
    manifest_items = []
    for rule in VOLUME_RULES:
        volume_entries = grouped.get(rule["id"], [])
        if not volume_entries:
            continue
        output_name = f"{rule['id']}_{slugify(rule['title'])}"
        manifest_items.append(
            {
                "volume_id": rule["id"],
                "series_title": series_title,
                "volume_title": rule["title"],
                "subtitle": rule["title"],
                "source_files": [entry["base_name"] for entry in volume_entries],
                "target_output_name": output_name,
            }
        )
    manifest_path = source_dir / "output" / "volume_manifest.json"
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest_items, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest_items, manifest_path


def run_volume_build(volume_dir, volume_title, output_name, output_dir, series_title):
    result = subprocess.run(
        [
            sys.executable,
            str(MERGE_AND_BUILD),
            str(volume_dir),
            volume_title,
            output_name,
            str(output_dir),
            series_title,
        ],
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"Volume build failed: {volume_title}")


def parse_args(argv):
    positional = []
    limit = None
    total_limit = None
    volume_filter = None
    i = 0
    while i < len(argv):
        arg = argv[i]
        if arg == "--limit":
            limit = int(argv[i + 1])
            i += 2
            continue
        if arg == "--total-limit":
            total_limit = int(argv[i + 1])
            i += 2
            continue
        if arg == "--volume":
            volume_filter = argv[i + 1]
            i += 2
            continue
        positional.append(arg)
        i += 1
    return positional, limit, total_limit, volume_filter


def main():
    positional, limit, total_limit, volume_filter = parse_args(sys.argv[1:])
    if len(positional) < 1:
        print("Usage: python run_book_pipeline.py <source_dir> [output_dir] [series_title] [--limit N] [--total-limit N] [--volume XX]")
        sys.exit(1)

    source_dir = Path(positional[0]).resolve()
    output_dir = Path(positional[1]).resolve() if len(positional) >= 2 else Path(r"d:\desk\new")
    series_title = positional[2] if len(positional) >= 3 else source_title(source_dir)

    normalized_entries = build_normalized_inputs(source_dir)
    if total_limit is not None:
        normalized_entries = normalized_entries[:total_limit]
    grouped = assign_volumes(normalized_entries)
    if volume_filter:
        grouped = {key: value for key, value in grouped.items() if key == volume_filter}
    if limit is not None:
        grouped = {key: value[:limit] for key, value in grouped.items()}
    manifest_items, manifest_path = build_manifest(source_dir, series_title, grouped)

    print(f"Normalized {len(normalized_entries)} files.")
    print(f"Manifest: {manifest_path}")

    for item in manifest_items:
        volume_entries = grouped[item["volume_id"]]
        volume_dir = source_dir / "output" / "volumes" / f"{item['volume_id']}_{slugify(item['volume_title'])}"
        write_clean_chapters(volume_dir, item["volume_title"], volume_entries)
        run_volume_build(volume_dir, item["volume_title"], item["target_output_name"], output_dir, series_title)

    print("Pipeline complete.")


if __name__ == "__main__":
    main()
