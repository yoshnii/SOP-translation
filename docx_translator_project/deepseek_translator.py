#!/usr/bin/env python3
"""
DeepSeek 章节级翻译引擎 —— 保格式 docx 翻译,靠上下文消歧(不依赖术语表)。

相比豆包翻译模型的根本改进:
  - 章节级:把一批段落(带上下文)一次性给 DeepSeek,[P1][P2] 标记进出,按标记拆回。
  - 上下文消歧:模型看到整章语境,自行判断"建库=Library Preparation"(无需术语表补丁)。
  - 指令遵循:NGS 背景 system + SOP 规范(祈使句/图题名词短语/英文标点)。
  - 轻后处理:删图题多余的 illustration、统一小瑕疵。

复用 translate_docx 的保格式管线(collect_segments / write_back / Segment),只换翻译层。

用法:
  export ARK_API_KEY=...
  python deepseek_translator.py 输入.docx -o 输出.docx
  python deepseek_translator.py 输入.docx --model deepseek-v4-flash-260425   # 换引擎
"""
from __future__ import annotations
import os, re, sys, json, argparse, time
from openai import OpenAI

import translate_docx as core   # 复用保格式管线
from docx import Document

BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
DEFAULT_MODEL = "deepseek-v4-pro-260425"

# 每个翻译块的最大段数(章节级:大到有上下文,小到不超输出上限)
CHUNK_SIZE = 40
MAX_WORKERS = 3   # DeepSeek 较慢,并发块数

SYSTEM = """You are a professional English translator specializing in NGS sequencing and molecular diagnostics. You are translating a laboratory SOP (Standard Operating Procedure) for an automated sequencing library-preparation / pathogen-detection platform.

Translate the Chinese paragraphs (marked [P1][P2]...) into professional, native English.

Rules:
1. Keep the [P1][P2]... markers before each translated paragraph; SAME count and order, no additions or omissions. Even if a paragraph is only a number/code, output its marker with the unchanged content.
2. This is a SEQUENCING library-prep workflow. Use correct, consistent NGS terminology throughout based on context. Key disambiguations: "建库"=Library Preparation (NOT database); "建库仓"=Library Preparation chamber; "上机"=loading onto the sequencer / sequencing (NOT boarding/computer); "中台"=Console (the control software); "洗液"=Wash Buffer; "标曲"=standard curve.
3. Keep the SAME translation for the same term across all paragraphs (consistency is critical for SOP safety).
4. Do not alter numbers, codes, units, Barcode IDs, Rack numbers, IP addresses, reagent codes (R1/R2 etc.).
5. Use imperative mood for operation steps; use concise NOUN PHRASES for figure/table captions (e.g. "Figure 12 System interface", NOT "Figure 12 shows the system interface" and NOT adding "illustration/diagram").
6. Use English half-width punctuation; never leave Chinese full-width punctuation.

Output ONLY the marked English translation, no explanations."""


def collect_all_segments(doc):
    """全覆盖收集:在 core.collect_segments(正文+目录) 基础上,补齐页眉/页脚里
    所有含中文的文本节点(标题、文件编号、页脚保密声明等,常在表格里,
    iter_block_paragraphs 抓不全)。确保除图片外的可编辑文字一个不漏。"""
    from docx.oxml.ns import qn
    segments = core.collect_segments(doc)
    seen = {id(s.toc_node) for s in segments if s.toc_node is not None}
    # 已被普通段落覆盖的页眉文字(按文本去重,避免重复翻)
    seen_texts = set()

    # 遍历所有 section 的所有页眉页脚的 XML,抽含中文的 <w:t>
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.first_page_header, section.first_page_footer,
                   section.even_page_header, section.even_page_footer):
            try:
                el = hf._element
            except Exception:
                continue
            for t in el.iter(qn("w:t")):
                if id(t) in seen:
                    continue
                s = t.text or ""
                if s.strip() and re.search(r"[一-鿿぀-ヿ가-힯]", s):
                    seen.add(id(t))
                    segments.append(core.Segment(text=s, toc_node=t))
    return segments


def split_chunks(segments):
    """把需翻译的段落切成章节级块。尽量在'章节标题'(如 6.8 / 6.9)前断开,
    让每块是一个自然小节(上下文完整)。"""
    idxs = [i for i, s in enumerate(segments) if not core.should_skip(s.text)]
    if not idxs:
        return []
    chunks = []
    cur = []
    for i in idxs:
        # 遇到一级小节标题(如 "6.9 xxx" / "6.10 xxx")且当前块已较大,就断开
        is_heading = bool(re.match(r"^\d+\.\d+\s|\d+\.\d+\.\d+\s", segments[i].text.strip()))
        if cur and (len(cur) >= CHUNK_SIZE or (is_heading and len(cur) >= CHUNK_SIZE // 2)):
            chunks.append(cur)
            cur = []
        cur.append(i)
    if cur:
        chunks.append(cur)
    return chunks


def load_glossary_dict(path):
    """读 glossary.csv,返回 {中文: 英文}(只取前两列,忽略错译补丁)。
    用于跨文档术语统一 —— 把术语注入 DeepSeek 的 prompt,让模型主动遵守。"""
    import csv as _csv
    g = {}
    if not path or not os.path.exists(path):
        return g
    for row in _csv.reader(open(path, newline="", encoding="utf-8")):
        if not row or row[0].strip().startswith("#") or len(row) < 2:
            continue
        zh, en = row[0].strip(), row[1].strip()
        if zh and en and zh != "中文":
            g[zh] = en
    return g


def _relevant_glossary(glossary, chunk_text):
    """只挑出该块文本里实际出现的术语,减少 prompt 体积、跨文档仍统一。
    按中文长度降序(长术语优先,避免子串误判)。"""
    hits = [(zh, en) for zh, en in glossary.items() if zh in chunk_text]
    hits.sort(key=lambda p: len(p[0]), reverse=True)
    return hits


def translate_chunk(client, model, segments, idxs, glossary=None):
    """翻译一个块,返回 {段index: 译文}。glossary 命中的术语注入 prompt 强制统一。"""
    marked = "\n".join(f"[P{k+1}] {segments[idxs[k]].text}" for k in range(len(idxs)))
    # 注入该块出现的术语(跨文档一致性)
    system = SYSTEM
    if glossary:
        chunk_text = "\n".join(segments[i].text for i in idxs)
        rel = _relevant_glossary(glossary, chunk_text)
        if rel:
            lines = "\n".join(f"  {zh} => {en}" for zh, en in rel)
            system = SYSTEM + (
                "\n\nMANDATORY GLOSSARY — when the source contains the term on the left, "
                "you MUST use exactly the English on the right (for cross-document consistency):\n"
                + lines
            )
    resp = client.chat.completions.create(
        model=model, temperature=0.2, max_tokens=8000,
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": marked}],
    )
    out = resp.choices[0].message.content.strip()
    # 按 [Pn] 拆回
    parts = re.split(r"\[P(\d+)\]", out)
    local = {}
    for j in range(1, len(parts), 2):
        local[int(parts[j]) - 1] = parts[j + 1].strip()
    # 映射回全局段index
    result = {}
    for k in range(len(idxs)):
        result[idxs[k]] = local.get(k, segments[idxs[k]].text)  # 缺失保留原文
    return result


def post_process(en: str) -> str:
    """轻后处理:删图题多余词、清中文标点。"""
    s = en
    # 删图/表标题里多余的 "illustration/diagram of"
    s = re.sub(r"\b(illustration|diagram|schematic)\s+of\s+", "", s, flags=re.I)
    s = re.sub(r"\s+(illustration|diagram)\b", "", s, flags=re.I)
    # 中文标点 -> 英文
    for zh, eng in {"。": ".", "，": ",", "；": ";", "：": ":", "！": "!", "？": "?",
                    "（": " (", "）": ") ", "、": ", "}.items():
        s = s.replace(zh, eng)
    s = re.sub(r"\s+([,.;:!?])", r"\1", s)
    s = re.sub(r"  +", " ", s).strip()
    return s


def translate_doc(in_path, out_path, model, progress_cb=None, glossary_path="glossary.csv"):
    key = os.environ.get("ARK_API_KEY")
    if not key:
        sys.exit("未设置 ARK_API_KEY")
    client = OpenAI(api_key=key, base_url=BASE_URL)

    glossary = load_glossary_dict(glossary_path)  # 跨文档术语统一(注入prompt)
    if glossary:
        print(f"术语表:加载 {len(glossary)} 条(注入prompt,保证跨文档一致)")

    doc = Document(in_path)
    segments = collect_all_segments(doc)   # 全覆盖:正文+目录+页眉页脚
    chunks = split_chunks(segments)
    print(f"段落总数 {len(segments)} | 切成 {len(chunks)} 个章节块 | 模型 {model}")

    from concurrent.futures import ThreadPoolExecutor, as_completed
    cache = {}
    done = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futs = {ex.submit(translate_chunk, client, model, segments, c, glossary): c for c in chunks}
        for fut in as_completed(futs):
            try:
                cache.update(fut.result())
            except Exception as e:
                idxs = futs[fut]
                for i in idxs:
                    cache[i] = segments[i].text
                print(f"  ⚠️ 块翻译失败({len(idxs)}段),保留原文: {type(e).__name__}")
            done += 1
            print(f"  进度 {done}/{len(chunks)} 块")
            if progress_cb:
                progress_cb(done, len(chunks))

    # 残留中文补翻:章节级翻译偶有个别段没翻全(跨中英边界的run、块边界等),
    # 扫描译文里还含中日韩字符的,单独逐段补翻一遍。
    residual = []
    for i, seg in enumerate(segments):
        if core.should_skip(seg.text):
            continue
        en = cache.get(i, seg.text)
        if re.search(r"[一-鿿぀-ヿ가-힯]", en):
            residual.append(i)
    if residual:
        print(f"  残留中文 {len(residual)} 段,逐段补翻...")
        for i in residual:
            try:
                r = translate_chunk(client, model, segments, [i], glossary)
                fixed = r.get(i, segments[i].text)
                if not re.search(r"[一-鿿]", fixed):  # 补翻成功才用
                    cache[i] = fixed
            except Exception:
                pass

    # 写回(后处理 + 保格式)
    for i, seg in enumerate(segments):
        if core.should_skip(seg.text):
            continue
        en = post_process(cache.get(i, seg.text))
        core.write_back(seg, en)

    doc.save(out_path)
    # 最终残留统计
    n_left = sum(1 for i, s in enumerate(segments)
                 if not core.should_skip(s.text) and re.search(r"[一-鿿]", post_process(cache.get(i, s.text))))
    print(f"完成 ✅ -> {out_path}  (残留中文 {n_left} 段)")
    return out_path


def main():
    ap = argparse.ArgumentParser(description="DeepSeek 章节级保格式翻译")
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--model", default=DEFAULT_MODEL)
    args = ap.parse_args()
    out = args.output or args.input.rsplit(".", 1)[0] + "_DS.docx"
    translate_doc(args.input, out, args.model)


if __name__ == "__main__":
    main()
