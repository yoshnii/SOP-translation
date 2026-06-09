#!/usr/bin/env python3
"""
translate_docx.py — 就地翻译 .docx,保留全部格式、图片、logo、表格、页眉页脚。

核心思路:代码只负责结构,AI 只负责文字。
  1. 遍历文档所有"段落"(正文 / 表格单元格 / 页眉 / 页脚),抽出纯文本。
  2. 去重 + 跳过纯编号/数字/日期,只把真正需要翻译的文本发给豆包翻译模型。
  3. 译后用术语表做一道查找替换,强制术语统一。
  4. 把译文写回每个段落的"第一个 run",清空其余 run。
     —— 段落级格式(对齐、字体、颜色、图片锚点、表格结构)全部不动,自动保留。

翻译引擎:豆包翻译专用模型 doubao-seed-translation(火山引擎方舟 /responses 接口)。
  - 该模型不走标准 chat/completions,而是 /responses,且 translation_options 只认源/目标语言。
  - 它一次只翻一段,所以这里用线程池并发逐段翻译。
  - 术语一致性靠"译后替换"实现(见 glossary.csv)。

用法:
  export ARK_API_KEY=...          # 火山引擎方舟控制台拿到的 API Key
  pip install python-docx requests
  python translate_docx.py 输入.docx                       # 默认中->英,输出 输入_en.docx
  python translate_docx.py 输入.docx -o 译文.docx           # 指定输出名
  python translate_docx.py 输入.docx --to zh                # 改成 英->中
  python translate_docx.py 输入.docx --glossary glossary.csv  # 指定术语表
  python translate_docx.py 输入.docx --dry-run             # 只抽取不调 API,看统计

已覆盖:正文段落、表格(含嵌套表格)、页眉、页脚。
暂未覆盖:文本框(<w:txbxContent>)、SmartArt、图表内文字 —— 见文末 TODO。
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass

try:
    from docx import Document
    from docx.document import Document as _DocType
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph
except ImportError:
    sys.exit("缺少依赖,请先运行:  pip install python-docx requests")

try:
    import requests
except ImportError:
    sys.exit("缺少依赖,请先运行:  pip install requests")

# ----------------------------------------------------------------------------
# 翻译方向。键是 --to 用的代码,值是豆包翻译模型的 language 代码。
# 豆包翻译模型语言代码:zh 中文 / en 英文 / ja 日文 / ko 韩文 等。
# ----------------------------------------------------------------------------
LANG_NAMES = {
    "en": "English", "zh": "Simplified Chinese", "ja": "Japanese",
    "ko": "Korean", "fr": "French", "de": "German", "es": "Spanish",
}

# ----------------------------------------------------------------------------
# 豆包(火山引擎方舟)配置
# ----------------------------------------------------------------------------
ARK_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"

# 翻译专用模型。它走 /responses 接口,不走 chat/completions。
DEFAULT_MODEL = os.environ.get("ARK_MODEL", "doubao-seed-translation-250915")

# 并发线程数:翻译模型逐段调用,并发能大幅提速。火山有 QPS 限制,8~16 较稳。
MAX_WORKERS = 8

# 源语言:中->英时源是 zh。若做英->中,把这里和 --to 一起改。
SOURCE_LANG = "zh"

# 翻译模型对"被全角括号包裹"或"以全角标点结尾"的短文本会原样不翻
# (如 注意事项: / 【注意】: / (20倍内标))。送翻译前剥掉这些外壳,
# 翻完再拼回。WRAP=成对包裹符;TRAIL=尾部可剥的标点。
WRAP_PAIRS = [("【", "】"), ("（", "）"), ("(", ")"),
              ("[", "]"), ("「", "」"), ("《", "》")]
TRAIL_PUNCT = "：:。.，,、;；!！?？ "


def split_shell(text: str) -> "tuple[str, str, str]":
    """把首尾的包裹符/标点剥离,返回 (前缀, 核心, 后缀)。核心是真正要翻的部分。"""
    s = text
    prefix = suffix = ""
    for L, R in WRAP_PAIRS:
        if s.startswith(L) and s.endswith(R) and len(s) > len(L) + len(R):
            prefix, suffix, s = L, R, s[len(L):-len(R)]
            break
    m = re.search(r"[%s]+$" % re.escape(TRAIL_PUNCT), s)
    if m:
        suffix = m.group() + suffix
        s = s[:m.start()]
    return prefix, s, suffix


def should_skip(text: str) -> bool:
    """判断一段文本是否无需翻译(纯编号/数字/英文符号,不含中日韩字符)。"""
    t = text.strip()
    if not t:
        return True
    if not re.search(r"[一-鿿぀-ヿ가-힯]", t):
        return True
    return False


# ----------------------------------------------------------------------------
# 术语表:中文 -> 固定译法。译后在英文译文里做替换,强制术语统一。
# ----------------------------------------------------------------------------
def load_glossary(path: str | None) -> "list[tuple[str, str, list]]":
    """读 CSV。列:中文,英文,备注[,错译]。返回 (中文, 英文, [错译...]) 列表,
    按中文长度降序(先替换长术语,避免'末端修复酶'被'末端修复'抢先匹配)。

    第4列"错译"可选,分号(;)分隔多个豆包常见错误英文译法;译后会被强制
    替换成标准英文(解决"豆包把术语翻成错误英文、术语表抓不到"的问题)。
    """
    if not path or not os.path.exists(path):
        if path:
            print(f"⚠️  术语表 {path} 不存在,跳过。")
        return []
    triples: list[tuple[str, str, list]] = []
    for row in csv.reader(open(path, newline="", encoding="utf-8")):
        if not row or row[0].strip().startswith("#") or len(row) < 2:
            continue
        zh, en = row[0].strip(), row[1].strip()
        if not (zh and en and zh != "中文"):
            continue
        wrongs = []
        if len(row) >= 4 and row[3].strip():
            wrongs = [w.strip() for w in row[3].split(";") if w.strip()]
        triples.append((zh, en, wrongs))
    triples.sort(key=lambda p: len(p[0]), reverse=True)
    nwrong = sum(1 for _, _, w in triples if w)
    print(f"术语表:加载 {len(triples)} 条固定译法(其中 {nwrong} 条带错译纠正)。")
    return triples


def load_overrides(path: str) -> "dict[str, str]":
    """读整句覆盖表 overrides.csv(原文,译文)。原文精确命中时直接用译文,
    优先级高于豆包翻译。用于固化 QC 人工修复的整句标准译文。"""
    overrides: dict[str, str] = {}
    if not os.path.exists(path):
        return overrides
    for row in csv.reader(open(path, newline="", encoding="utf-8")):
        if not row or row[0].strip().startswith("#") or len(row) < 2:
            continue
        zh, en = row[0].strip(), row[1].strip()
        if zh and en and zh != "原文":
            overrides[zh] = en
    if overrides:
        print(f"整句覆盖表:加载 {len(overrides)} 条。")
    return overrides


def apply_glossary(src_zh: str, translated_en: str, glossary) -> str:
    """对一段译文做术语替换,三重纠正:

    1. 大小写归一:译文里标准译法的大小写变体 → 统一成标准写法。
    2. 错译纠正:若原文含某术语,且译文出现该术语的【已知错误英文译法】
       (术语表第4列),强制替换成标准英文。解决"豆包把术语翻成错误英文、
       靠中文匹配抓不到"的问题(如 建库→database construction→Library Preparation)。
    3. 残留中文兜底:译文仍残留该中文术语(模型没翻)→ 直接替换成标准英文。
    """
    if not glossary:
        return translated_en
    out = translated_en
    for zh, en, wrongs in glossary:
        in_src = zh in src_zh
        if in_src:
            # 1. 大小写归一
            out = re.compile(re.escape(en), re.IGNORECASE).sub(en, out)
            # 2. 错译纠正:把已知错误英文换成标准英文(整词、忽略大小写、
            #    兼容词尾复数 s,这样错译列写单数即可同时匹配单/复数)。
            for w in wrongs:
                out = re.compile(r"\b" + re.escape(w) + r"s?\b", re.IGNORECASE).sub(en, out)
        # 3. 残留中文兜底
        if zh in out:
            out = out.replace(zh, en)
    return _normalize_punct(out)


# 中文全角标点 → 英文半角(翻译模型常在英文译文里残留 。：；,!? 等)
_PUNCT_MAP = str.maketrans({
    "。": ".", "，": ",", "；": ";", "：": ":", "！": "!", "？": "?",
    "（": " (", "）": ") ", "、": ", ", "～": "~",
    "“": '"', "”": '"', "‘": "'", "’": "'",
})


def _normalize_punct(en: str) -> str:
    """把英文译文里残留的中文全角标点换成英文半角,并清理多余空格。"""
    s = en.translate(_PUNCT_MAP)
    s = re.sub(r"\s+([,.;:!?])", r"\1", s)   # 标点前的空格去掉
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r"\s+\)", ")", s)
    s = re.sub(r"  +", " ", s).strip()
    return s


# ----------------------------------------------------------------------------
# 第一步:遍历文档,收集所有"可翻译段落"
# ----------------------------------------------------------------------------
@dataclass
class Segment:
    text: str
    # 普通段落:写回 paragraph.runs。二选一。
    paragraph: "Paragraph | None" = None
    # 目录(TOC)文本节点:写回这个 <w:t> 节点的 .text。
    toc_node: "object | None" = None


def iter_block_paragraphs(parent) -> "list[Paragraph]":
    """递归遍历 body / 单元格里的所有段落,包括嵌套表格,按 XML 顺序。"""
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

    paragraphs: list[Paragraph] = []
    if isinstance(parent, _DocType):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        parent_elm = parent

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            paragraphs.append(Paragraph(child, parent))
        elif isinstance(child, CT_Tbl):
            table = Table(child, parent)
            for row in table.rows:
                for cell in row.cells:
                    paragraphs.extend(iter_block_paragraphs(cell))
    return paragraphs


def collect_segments(doc) -> "list[Segment]":
    """收集正文 + 页眉页脚段落,以及目录(TOC)里的中文文本节点。"""
    from docx.oxml.ns import qn
    segments: list[Segment] = []

    def add_from(container):
        for para in iter_block_paragraphs(container):
            # 目录段落:文字 run 嵌在 <w:hyperlink> 里,para.runs 为空。
            # 单独抽里面含中文的 <w:t> 节点(跳过编号/页码/域指令)。
            if para.style is not None and para.style.name.lower().startswith("toc"):
                for t in para._p.iter(qn("w:t")):
                    s = t.text or ""
                    if s.strip() and re.search(r"[一-鿿぀-ヿ가-힯]", s):
                        segments.append(Segment(text=s, toc_node=t))
                continue
            if para.text.strip():
                segments.append(Segment(text=para.text, paragraph=para))

    add_from(doc)
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.first_page_header, section.first_page_footer,
                   section.even_page_header, section.even_page_footer):
            try:
                add_from(hf)
            except Exception:
                pass

    # 目录(TOC)有两种承载,iter_block_paragraphs 都进不去,单独抽:
    #   1) 普通 toc 样式段落里 <w:hyperlink> 中的文字(上面 add_from 已处理)
    #   2) 包在 <w:sdt> 内容控件里的目录(python-docx 的 paragraphs 看不到)
    # 这里扫描全文所有 sdt,抽其中含中日韩字符的 <w:t> 节点(跳过编号/页码/域)。
    seen_nodes = {id(s.toc_node) for s in segments if s.toc_node is not None}
    for sdt in doc.element.body.iter(qn("w:sdt")):
        for t in sdt.iter(qn("w:t")):
            if id(t) in seen_nodes:
                continue
            s = t.text or ""
            if s.strip() and re.search(r"[一-鿿぀-ヿ가-힯]", s):
                segments.append(Segment(text=s, toc_node=t))
    return segments


# ----------------------------------------------------------------------------
# 第二步:调用豆包翻译模型(/responses 接口),逐段翻译
# ----------------------------------------------------------------------------
def _call_api(session, api_key, model, text, target_lang) -> str:
    """真正发一次翻译请求。"""
    body = {
        "model": model,
        "input": [{
            "role": "user",
            "content": [{
                "type": "input_text",
                "text": text,
                "translation_options": {
                    "source_language": SOURCE_LANG,
                    "target_language": target_lang,
                },
            }],
        }],
    }
    r = session.post(
        f"{ARK_BASE_URL}/responses",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json=body, timeout=60,
    )
    r.raise_for_status()
    out = r.json()["output"][0]["content"][0]["text"]
    return out.strip(" \t\r\n")  # 模型偶尔在首尾加换行,清掉


def _looks_bad(core_zh: str, en: str) -> bool:
    """判断一条译文是否可疑,需要重试。

    短文本(尤其孤立词)翻译模型会:① 原样不翻(译文仍含中文);
    ② 幻觉补一整句(译文比原文长数倍)。两种都判为 bad。
    """
    if re.search(r"[一-鿿]", en):           # 还残留中文 = 没翻
        return True
    # 幻觉:中文核心很短(<=8字)但英文暴涨(>核心字数的6倍且>40字符)
    if len(core_zh) <= 8 and len(en) > max(40, len(core_zh) * 6):
        return True
    return False


def translate_one(session, api_key, model, text, target_lang,
                  gloss_exact=None, overrides=None, retries=5) -> str:
    """翻一段。失败时抛异常由上层兜底。

    优先级:① overrides 整句覆盖(QC人工修复的标准译文,最高优先)→ ② 术语表
    短路 → ③ 豆包翻译。
    预处理:压空格 + 剥外壳(全角括号/标点包裹的短文本模型不翻)。
    其余文本走模型;短文本可能不翻或幻觉,_looks_bad 命中则重试。
    """
    clean = re.sub(r"\s+", " ", text).strip()
    # ① 整句覆盖:原文(原始或压空格后)精确命中 → 直接用人工标准译文
    if overrides:
        if text.strip() in overrides:
            return overrides[text.strip()]
        if clean in overrides:
            return overrides[clean]
    prefix, core, suffix = split_shell(clean)
    if not re.search(r"[一-鿿぀-ヿ가-힯]", core):
        return clean
    # ② 术语表精确命中 → 直接固定译法,跳过 API
    if gloss_exact and core in gloss_exact:
        return prefix + gloss_exact[core] + suffix
    en = ""
    for _ in range(retries):
        en = _call_api(session, api_key, model, core, target_lang)
        if not _looks_bad(core, en):
            return prefix + en + suffix
    # 多次仍不理想:返回最后一次结果(总比丢内容好)
    return prefix + en + suffix


def translate_all(api_key, model, texts, target_lang, glossary, overrides=None) -> "dict[str, str]":
    """并发翻译所有唯一文本,返回 原文->译文 缓存。术语替换在此应用。"""
    session = requests.Session()
    cache: dict[str, str] = {}
    done = 0
    total = len(texts)
    overrides = overrides or {}
    # 精确匹配表:整段(剥壳后)等于某词条时直接用固定译法,不调 API。
    gloss_exact = {zh: en for zh, en, _ in glossary}

    def work(t):
        en = translate_one(session, api_key, model, t, target_lang, gloss_exact, overrides)
        return t, apply_glossary(t, en, glossary)

    failed: list[str] = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(work, t): t for t in texts}
        for fut in as_completed(futures):
            src = futures[fut]
            try:
                t, en = fut.result()
                cache[t] = en
            except Exception:
                failed.append(src)  # 先记下,稍后串行重翻
            done += 1
            if done % 25 == 0 or done == total:
                print(f"  进度 {done}/{total}")

    # 并发期失败的段(多为限流/抖动),串行重翻一遍,降低永久漏翻。
    if failed:
        print(f"  并发期 {len(failed)} 段失败,串行补翻...")
        for src in failed:
            try:
                en = translate_one(session, api_key, model, src, target_lang, gloss_exact, overrides)
                cache[src] = apply_glossary(src, en, glossary)
            except Exception as e:
                cache[src] = src  # 仍失败才保留原文
                print(f"    ⚠️ 仍失败,保留原文:{src[:30]}... ({type(e).__name__})")
    return cache


# ----------------------------------------------------------------------------
# 第三步:把译文写回段落(只动 run.text,不碰任何格式)
# ----------------------------------------------------------------------------
def write_back(segment: Segment, translated: str) -> None:
    # 目录文本节点:直接改 <w:t> 的文字,不碰域结构。
    if segment.toc_node is not None:
        segment.toc_node.text = translated
        return
    runs = segment.paragraph.runs
    if not runs:
        return
    runs[0].text = translated
    for r in runs[1:]:
        r.text = ""


def mark_fields_dirty(doc) -> None:
    """让 Word 打开文档时更新所有域(目录 TOC 等)。

    目录是自动生成的域,我们翻译了正文标题但目录缓存的还是旧中文。
    设 settings 里 <w:updateFields w:val="true"/>,Word 打开时会提示
    "是否更新域",点"是"目录即按英文标题重建。
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        settings.insert(0, uf)


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="就地翻译 .docx,保留格式/图片/排版(豆包翻译模型)")
    ap.add_argument("input", help="输入 .docx 路径")
    ap.add_argument("-o", "--output", help="输出 .docx 路径(默认在原名后加 _<lang>)")
    ap.add_argument("--to", default="en", help=f"目标语言,默认 en。可选:{', '.join(LANG_NAMES)}")
    ap.add_argument("--glossary", default="glossary.csv", help="术语表 CSV(默认 glossary.csv)")
    ap.add_argument("--overrides", default="overrides.csv", help="整句覆盖表 CSV(默认 overrides.csv,原文精确命中直接用指定译文)")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"模型,默认 {DEFAULT_MODEL}")
    ap.add_argument("--dry-run", action="store_true", help="只统计不调 API")
    args = ap.parse_args()

    if not os.path.exists(args.input):
        sys.exit(f"找不到文件:{args.input}")
    out_path = args.output or _default_out(args.input, args.to)

    print(f"读取:{args.input}")
    doc = Document(args.input)
    segments = collect_segments(doc)
    to_translate = [s for s in segments if not should_skip(s.text)]
    skipped = len(segments) - len(to_translate)
    unique_texts = list(dict.fromkeys(s.text for s in to_translate))
    print(f"段落总数 {len(segments)} | 跳过纯编号 {skipped} | 需翻译 {len(to_translate)} "
          f"| 去重后唯一文本 {len(unique_texts)}")

    glossary = load_glossary(args.glossary)
    overrides = load_overrides(args.overrides)

    if args.dry_run:
        print("\n--- dry-run:去重后将翻译的前 20 条 ---")
        for i, t in enumerate(unique_texts[:20]):
            print(f"  [{i}] {t[:60]}")
        if len(unique_texts) > 20:
            print(f"  ... 还有 {len(unique_texts) - 20} 条")
        return

    api_key = os.environ.get("ARK_API_KEY")
    if not api_key:
        sys.exit("未设置 ARK_API_KEY 环境变量(火山引擎方舟控制台获取)。")

    print(f"引擎:豆包翻译模型 {args.model} | 并发 {MAX_WORKERS} | 共 {len(unique_texts)} 段")
    cache = translate_all(api_key, args.model, unique_texts, args.to, glossary, overrides)

    for seg in segments:
        if should_skip(seg.text):
            continue
        write_back(seg, cache.get(seg.text, seg.text))

    # 注意:不调用 mark_fields_dirty。我们已直接翻译目录文字,若再让 Word
    # 更新域,会用正文里(样式混乱、未必抓得到)的标题重算,反而可能覆盖回中文。
    doc.save(out_path)
    print(f"完成 ✅  输出:{out_path}")


def _default_out(input_path: str, lang: str) -> str:
    base, ext = os.path.splitext(input_path)
    return f"{base}_{lang}{ext}"


if __name__ == "__main__":
    main()

# ----------------------------------------------------------------------------
# TODO(第二阶段):
#   - 文本框 <w:txbxContent> 里的文字:需直接遍历 XML 的 w:txbxContent//w:p。
#   - 段内逐字格式保留:若必须保留段中加粗/超链接,改按 run 分组翻译。
#   - 术语库:若火山控制台支持配置术语库,可换成模型侧术语干预,效果更准。
# ----------------------------------------------------------------------------
