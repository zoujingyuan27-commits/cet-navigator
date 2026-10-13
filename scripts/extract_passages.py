"""
extract_passages.py
从 doc/ 目录下的 .docx 文件中提取文章原文，生成 data/passages.json

文件名格式：英语六级{year}年{month}月第{set}套阅读理解（四选一）第{passage_index}篇真题.docx
每个文件包含：1篇文章 + 对应的5道题目（Q46-50 或 Q51-55）

输出：data/passages.json，以 passage_id 为键
passage_id 格式：CET6-RC-{year}{month:02d}-S{set}-P{passage_index}
  例如：CET6-RC-202306-S3-P2
"""

import json
import re
import sys
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("错误：缺少 python-docx 库，请先运行：pip install python-docx")
    sys.exit(1)


# ── 路径配置 ──────────────────────────────────────────────
ROOT_DIR = Path(__file__).parent.parent
DOC_DIR  = ROOT_DIR / "doc"
OUT_FILE = ROOT_DIR / "data" / "passages.json"


# ── 文件名解析 ────────────────────────────────────────────
FILENAME_PATTERN = re.compile(
    r"英语六级"
    r"(\d{4})年"   # year
    r"(\d{1,2})月"  # month
    r"第(\d+)套"   # set
    r"阅读理解"
    r".*?"
    r"第(\d+)篇"   # passage_index (1 or 2)
    r"真题"
)


def parse_filename(filename: str) -> dict | None:
    """从文件名提取 year/month/set/passage_index，失败返回 None"""
    m = FILENAME_PATTERN.search(filename)
    if not m:
        return None
    year, month, set_num, p_idx = (int(x) for x in m.groups())
    passage_id = f"CET6-RC-{year}{month:02d}-S{set_num}-P{p_idx}"
    return {
        "passage_id":    passage_id,
        "year":          year,
        "month":         month,
        "set":           set_num,
        "passage_index": p_idx,
    }


# ── docx 文章提取 ─────────────────────────────────────────
# 文章结束标志：遇到题号行（如 "46." 或 "46、" 开头）
QUESTION_NUM_RE = re.compile(r"^\s*\d{2}[\.、．]\s+\S")

# 文章开始标志：Questions XX to XX are based on ...
PASSAGE_START_RE = re.compile(
    r"Questions?\s+\d+\s+to\s+\d+\s+are\s+based\s+on", re.IGNORECASE
)

# 需要跳过的行（非文章正文）
SKIP_PATTERNS = [
    re.compile(r"^\s*真题卷第"),                        # 标题行
    re.compile(r"^\s*Part\s+III", re.IGNORECASE),      # Part III Section C
    re.compile(r"^\s*Directions\s*:", re.IGNORECASE),  # Directions 开头
    re.compile(r"^\s*Passage\s+(One|Two)\s*$", re.IGNORECASE),  # "Passage One"
    re.compile(r"^\s*Questions?\s+\d+\s+to\s+\d+", re.IGNORECASE),  # Questions 46 to 50...
    re.compile(r"^\s*[A-D][)）]\s+\S"),                # 选项行 A) B) C) D)
    re.compile(r"^\s*$"),                              # 空行
]


def should_skip(line: str) -> bool:
    return any(p.search(line) for p in SKIP_PATTERNS)


def extract_passage_text(docx_path: Path) -> str:
    """
    从 docx 中提取文章正文。

    策略：
    1. 找到 "Questions XX to XX are based on..." 这一行
    2. 该行之后开始收集段落
    3. 遇到题号行（46. / 51.）停止
    4. 跳过 Directions、Passage One/Two、选项行等
    """
    doc = Document(docx_path)
    paragraphs = [p.text for p in doc.paragraphs]

    collecting = False
    passage_lines = []

    for text in paragraphs:
        # 找到开始标志
        if not collecting:
            if PASSAGE_START_RE.search(text):
                collecting = True
            continue  # 开始标志这一行本身不收录

        # 遇到题号，停止
        if QUESTION_NUM_RE.match(text):
            break

        # 跳过非正文行
        if should_skip(text):
            continue

        passage_lines.append(text.strip())

    return "\n".join(passage_lines).strip()


# ── 主流程 ────────────────────────────────────────────────
def main():
    if not DOC_DIR.exists():
        print(f"错误：找不到 doc/ 目录：{DOC_DIR}")
        sys.exit(1)

    docx_files = sorted(DOC_DIR.glob("*.docx"))
    if not docx_files:
        print(f"错误：doc/ 目录下没有找到 .docx 文件")
        sys.exit(1)

    # 读取已有的 passages.json（增量更新）
    if OUT_FILE.exists():
        with open(OUT_FILE, encoding="utf-8") as f:
            passages = json.load(f)
        print(f"已有 {len(passages)} 条记录，将增量更新")
    else:
        passages = {}

    success, skipped, failed = 0, 0, 0

    for docx_path in docx_files:
        meta = parse_filename(docx_path.name)
        if meta is None:
            print(f"  [跳过] 文件名无法解析：{docx_path.name}")
            skipped += 1
            continue

        pid = meta["passage_id"]
        print(f"  处理：{pid}  ←  {docx_path.name}")

        try:
            text = extract_passage_text(docx_path)
        except Exception as e:
            print(f"    ✗ 读取失败：{e}")
            failed += 1
            continue

        if not text:
            print(f"    ⚠ 未提取到文章正文，请检查该文件格式")
            failed += 1
            continue

        passages[pid] = {
            **{k: v for k, v in meta.items() if k != "passage_id"},
            "source_file": docx_path.name,
            "text": text,
            "char_count": len(text),
        }
        success += 1

    # 按 passage_id 排序后写入
    passages_sorted = dict(sorted(passages.items()))
    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(passages_sorted, f, ensure_ascii=False, indent=2)

    print(f"\n完成！成功 {success} / 跳过 {skipped} / 失败 {failed}")
    print(f"输出文件：{OUT_FILE}")

    # 抽查：打印第一条记录的前200字
    if passages_sorted:
        first_id, first = next(iter(passages_sorted.items()))
        print(f"\n── 抽查 {first_id} ──")
        print(first["text"][:200], "...")


if __name__ == "__main__":
    main()
