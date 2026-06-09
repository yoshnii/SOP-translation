#!/usr/bin/env python3
"""一次性验证脚本:不调 API,用假译文替换文字,验证格式/图片/排版是否保留。

把每段文字替换成 "[EN] 原文",生成 source_FAKE.docx。
打开它应该看到:图片、logo、表格、页眉全在,只有文字前面多了 [EN]。
这证明"格式锁死"成立。
"""
import sys
from translate_docx import Document, collect_segments, write_back

doc = Document("source.docx")
segs = collect_segments(doc)
print(f"段落数:{len(segs)}")
for s in segs:
    write_back(s, "[EN] " + s.text)
doc.save("source_FAKE.docx")
print("已生成 source_FAKE.docx —— 打开它检查图片/表格/排版是否原样保留。")
