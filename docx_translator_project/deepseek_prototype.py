#!/usr/bin/env python3
"""
DeepSeek 章节级翻译原型 —— 验证"上下文消歧 + 术语硬约束 + 不确定标记"。

对比豆包翻译模型(单段、无上下文、无指令)的关键改进:
  1. 章节级:整章一次性给模型,带 [P1][P2]... 段落标记,按标记拆回(保格式)。
  2. 上下文消歧:模型能看到整章语境,自行判断"建库=Library Preparation"还是 database。
  3. 术语硬约束:术语表 + 触发条件写进 prompt,命中必须用指定译法。
  4. 不确定标记:判据不足时输出 [?:候选A/候选B],留人工 review,不静默错译。

用法:
  export ARK_API_KEY=...
  python deepseek_prototype.py            # 翻译 _deepseek_test_chunk.json 并打印结果
"""
import os, json, re, sys
from openai import OpenAI

MODEL = "deepseek-v3-2-251201"
BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"

# 带触发条件的术语表(消歧判据)——这是关键改进:不只给译法,给判据。
GLOSSARY_RULES = """
术语对照(命中时必须用右侧译法;带"当…"的按上下文判断):
- 建库 => Library Preparation （NGS测序文库制备;当上下文涉及仪器/试剂/cycle/DNB/测序/Rack/分装时。本SOP是测序流程,几乎总是此义,绝不是database）
- 建库仓 => Library Preparation Module （"仓"=物理模块,强信号,绝不是database）
- 建库板号 => Library Preparation plate number
- 建库浓度 => Library Preparation concentration
- 建库孔位 => Library Preparation well position
- 前处理 => Pre-processing
- 去宿主 => Host depletion
- 上机 / 上机前准备 => loading onto the instrument / Pre-loading Preparation （上测序仪,不是登机/电脑）
- 洗液 => Wash Buffer （不是 Lotion）
- 标曲 / 标准曲线 => standard curve
- 中台 => Task Center / middleware （全文统一一种）
- 冻存管 => Cryovial
- 试剂槽 => Reagent Trough
- 磁珠 => Magnetic Beads
- 分装 => Aliquoting
- 通量 => throughput
- DNB => DNB (DNA Nanoball)
- 全流程 => Full Process
"""

SYSTEM = f"""你是一位资深的 NGS 测序/分子诊断领域的专业英文翻译。你在翻译一份实验室 SOP(标准作业指导书),关于 GenSIRO-16 自动化测序文库制备/病原检测平台。

任务:把下面带 [P1][P2]... 标记的中文段落逐段翻译成专业、地道(native)的英文。

严格规则:
1. 保持段落标记:每段译文前保留对应的 [P1][P2]... 标记,顺序和数量与原文完全一致,不增不减。
2. 术语硬约束:命中下方术语表的词,必须用指定英文译法,不得自由发挥。
{GLOSSARY_RULES}
3. 上下文消歧:利用整章语境判断多义词。本文档是测序文库制备流程,"建库"=Library Preparation(绝不是 database/数据库)。
4. 只翻译,不改写:不增删内容,保留数字、代号、单位、Barcode、Rack编号、试剂代号(R1/R2等)不变。
5. 不确定就标记:若某术语判据不足、你拿不准,输出 [?:候选A/候选B] 而不是硬翻,留给人工审校。
6. SOP 语气:操作步骤用祈使句;图表标题用名词短语(不要译成完整句子)。
7. 标点用英文半角,不要残留中文全角标点。

只输出带标记的英文译文,不要解释。"""


def main():
    key = os.environ.get("ARK_API_KEY")
    if not key:
        sys.exit("未设置 ARK_API_KEY")
    chunk = json.load(open("_deepseek_test_chunk.json", encoding="utf-8"))

    # 组装带标记的输入
    marked = "\n".join(f"[P{i+1}] {t}" for i, t in enumerate(chunk))

    client = OpenAI(api_key=key, base_url=BASE_URL)
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0.2,
        max_tokens=8000,
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": marked},
        ],
    )
    out = resp.choices[0].message.content.strip()

    # 按标记拆回
    parts = re.split(r"\[P(\d+)\]", out)
    result = {}
    for i in range(1, len(parts), 2):
        idx = int(parts[i]) - 1
        result[idx] = parts[i + 1].strip()

    print(f"=== DeepSeek 章节级翻译结果({len(result)}/{len(chunk)} 段)===\n")
    for i, zh in enumerate(chunk):
        en = result.get(i, "⚠️未返回")
        print(f"[{i+1}] 原: {zh[:45]}")
        print(f"    译: {en[:70]}")
        if "[?" in en:
            print(f"    ⚠️ 含不确定标记(待人工审校)")
        print()

    # 统计:建库消歧
    full = " ".join(result.values())
    print("=== 消歧检查 ===")
    print(f"  database 错译: {len(re.findall(r'database', full, re.I))} 处")
    print(f"  Library Preparation: {len(re.findall(r'Library Preparation', full, re.I))} 处")
    print(f"  不确定标记 [?]: {full.count('[?')} 处")
    print(f"  中文标点残留: {'有' if re.search(r'[。，；：！？]', full) else '无'}")


if __name__ == "__main__":
    main()
