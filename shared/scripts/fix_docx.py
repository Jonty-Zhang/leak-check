#!/usr/bin/env python3
"""按 leak-check 的扫描结果改 .docx，保留原有格式。

    fix_docx.py 论文.docx edits.json --inplace          # 原地改，先备份成 论文.leakfix-backup.docx
    fix_docx.py 论文.docx edits.json -o 论文_改.docx
    fix_docx.py 论文.docx --accept-revisions --remove-comments --clear-highlight --inplace

edits.json 是一个列表，每项是下面一种：
    {"para": 12, "find": "根据用户的要求，", "replace": ""}     第 12 段里替换（段号同 scan_leaks.py）
    {"find": "（改进版）", "replace": ""}                      不给 para：全文（含页眉页脚）都替换
    {"find": "图\\s*5", "replace": "图 4", "regex": true}      正则
    {"para": 30, "delete": true, "find": "希望对你有帮助"}       删掉第 30 段；给了 find 就先核对这段确实含这句
可选 "count": 只替换前 N 处。

替换时文字落在第一个被命中的 run 里，沿用它的字体字号；跨 run 的命中也能改。
任何一条没命中就整份不写入（退出码 1），避免改了一半。
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from scan_leaks import _q, iter_docx_paragraphs, p_text  # noqa: E402

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def replace_in_paragraph(p, pattern, repl, count=0, is_regex=False):
    """在一个 w:p 里替换，返回替换次数。pattern 是编译好的正则；is_regex 时 repl 里可以用 \\1 引用分组。"""
    ts = list(p.iter(_q("t")))
    if not ts:
        return 0
    full = "".join(t.text or "" for t in ts)
    matches = list(pattern.finditer(full))
    if count:
        matches = matches[:count]
    if not matches:
        return 0
    # 每个字符属于哪个 w:t
    owner, offset = [], []
    for ti, t in enumerate(ts):
        for oi in range(len(t.text or "")):
            owner.append(ti)
            offset.append(oi)
    texts = [t.text or "" for t in ts]
    for m in reversed(matches):  # 从后往前改，前面的下标不受影响
        if m.start() == m.end():
            continue
        new = m.expand(repl) if is_regex else repl
        si, so = owner[m.start()], offset[m.start()]
        ei, eo = owner[m.end() - 1], offset[m.end() - 1] + 1
        if si == ei:
            texts[si] = texts[si][:so] + new + texts[si][eo:]
        else:
            texts[si] = texts[si][:so] + new
            for k in range(si + 1, ei):
                texts[k] = ""
            texts[ei] = texts[ei][eo:]
    for t, s in zip(ts, texts):
        t.text = s
        if s != s.strip() or "  " in s:
            t.set(XML_SPACE, "preserve")
        run = t.getparent()
        if not s and run is not None and run.tag == _q("r") and all(c.tag in (_q("rPr"), _q("t")) for c in run):
            run.getparent().remove(run)  # 只剩空文字的 run 删掉
    return len(matches)


def delete_paragraph(p):
    parent = p.getparent()
    siblings = [c for c in parent if c.tag == _q("p")]
    if parent.tag == _q("tc") and len(siblings) == 1:
        for t in p.iter(_q("t")):
            t.text = ""  # 单元格里至少要留一个段落
    else:
        parent.remove(p)


def accept_revisions(root):
    n = 0
    for tag in ("del", "moveFrom"):
        for el in list(root.iter(_q(tag))):
            parent = el.getparent()
            if parent is not None and parent.tag != _q("rPr"):
                parent.remove(el)
                n += 1
    for tag in ("ins", "moveTo"):
        for el in list(root.iter(_q(tag))):
            parent = el.getparent()
            if parent is None:
                continue
            if parent.tag == _q("rPr"):  # 段落标记上的修订记号
                parent.remove(el)
                continue
            idx = parent.index(el)
            for child in list(el):
                parent.insert(idx, child)
                idx += 1
            parent.remove(el)
            n += 1
    for tag in ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange"):
        for el in list(root.iter(_q(tag))):
            el.getparent().remove(el)
            n += 1
    for rpr in root.iter(_q("rPr")):
        for tag in ("del", "ins"):
            for el in rpr.findall(_q(tag)):
                rpr.remove(el)
    return n


def remove_comment_marks(root):
    n = 0
    for tag in ("commentRangeStart", "commentRangeEnd"):
        for el in list(root.iter(_q(tag))):
            el.getparent().remove(el)
    for ref in list(root.iter(_q("commentReference"))):
        run = ref.getparent()
        if run is not None and run.tag == _q("r") and run.getparent() is not None:
            run.getparent().remove(run)
            n += 1
    return n


def clear_highlight(root):
    n = 0
    for el in list(root.iter(_q("highlight"))):
        el.getparent().remove(el)
        n += 1
    return n


def strip_comments_part(path):
    """把 comments.xml 里的批注清空（保留空的根节点，文档关系不用动）。"""
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    changed = False
    out = []
    for info, data in items:
        if info.filename in ("word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsExtensible.xml"):
            text = data.decode("utf-8")
            new = re.sub(r"<w(15|16cid|16cex)?:comment(Ex|Id|Extensible)?\b[^>]*?(/>|>.*?</w(15|16cid|16cex)?:comment(Ex|Id|Extensible)?>)", "", text, flags=re.S)
            if new != text:
                changed = True
                data = new.encode("utf-8")
        out.append((info, data))
    if changed:
        tmp = Path(tempfile.mkstemp(suffix=".docx")[1])
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for info, data in out:
                z.writestr(info, data)
        shutil.move(tmp, path)


def main(argv=None):
    ap = argparse.ArgumentParser(description="按扫描结果修改 docx，保留格式")
    ap.add_argument("docx")
    ap.add_argument("edits", nargs="?", help="edits.json")
    ap.add_argument("-o", "--output")
    ap.add_argument("--inplace", action="store_true", help="原地修改（先备份成 .leakfix-backup.docx）")
    ap.add_argument("--accept-revisions", action="store_true", help="接受所有修订")
    ap.add_argument("--remove-comments", action="store_true", help="删除所有批注")
    ap.add_argument("--clear-highlight", action="store_true", help="去掉荧光高亮")
    ap.add_argument("--dry-run", action="store_true", help="只报告会改什么，不写文件")
    args = ap.parse_args(argv)
    if not (args.output or args.inplace or args.dry_run):
        ap.error("请给 -o 输出路径，或 --inplace，或 --dry-run")

    from docx import Document
    src = Path(args.docx)
    document = Document(str(src))
    paras = {n: p for n, p, _ in iter_docx_paragraphs(document)}
    extra = []
    for section in document.sections:
        for part in (section.header, section.footer, section.first_page_header, section.first_page_footer,
                     section.even_page_header, section.even_page_footer):
            try:
                extra.extend(p._p for p in part.paragraphs)
            except Exception:
                pass

    edits = json.loads(Path(args.edits).read_text(encoding="utf-8")) if args.edits else []
    failures, log, to_delete = [], [], []
    for i, e in enumerate(edits, 1):
        label = f"#{i}"
        if "para" in e and e["para"] not in paras:
            failures.append(f"{label} 第 {e['para']} 段不存在（全文共 {len(paras)} 段）")
            continue
        if e.get("delete"):
            p = paras.get(e.get("para"))
            if p is None:
                failures.append(f"{label} 删除段落必须给 para")
                continue
            if e.get("find") and e["find"] not in p_text(p):
                failures.append(f"{label} 第 {e['para']} 段里没有“{e['find']}”，没删（段落内容：{p_text(p)[:40]}）")
                continue
            to_delete.append(p)
            log.append(f"{label} 删除第 {e['para']} 段：{p_text(p)[:50]}")
            continue
        find = e.get("find")
        if not find:
            failures.append(f"{label} 缺少 find")
            continue
        pattern = re.compile(find if e.get("regex") else re.escape(find))
        repl = e.get("replace", "")
        scope = [paras[e["para"]]] if "para" in e else list(paras.values()) + extra
        total = 0
        for p in scope:
            left = e.get("count", 0) - total if e.get("count") else 0
            if e.get("count") and left <= 0:
                break
            before = p_text(p)
            k = replace_in_paragraph(p, pattern, repl, left, bool(e.get("regex")))
            if k:
                total += k
                log.append(f"{label} {'第 %d 段' % e['para'] if 'para' in e else '全文'}：“{before[:40]}” → “{p_text(p)[:40]}”")
        if not total:
            where = f"第 {e['para']} 段" if "para" in e else "全文"
            failures.append(f"{label} {where}没找到“{find}”")

    root = document.element.body
    if args.accept_revisions:
        log.append(f"接受修订：{accept_revisions(root)} 处")
    if args.remove_comments:
        log.append(f"删除批注标记：{remove_comment_marks(root)} 处")
    if args.clear_highlight:
        log.append(f"去掉高亮：{clear_highlight(root)} 处")
    for p in to_delete:
        delete_paragraph(p)

    for line in log:
        print(line)
    if failures:
        print("\n以下修改没执行，整份文件未写入：", file=sys.stderr)
        for f in failures:
            print("  " + f, file=sys.stderr)
        return 1
    if args.dry_run:
        print("（dry-run，没有写文件）")
        return 0
    if args.inplace:
        backup = src.with_name(src.stem + ".leakfix-backup" + src.suffix)
        if not backup.exists():
            shutil.copy2(src, backup)
            print(f"原文件备份：{backup}")
        out = src
    else:
        out = Path(args.output)
    document.save(str(out))
    if args.remove_comments:
        strip_comments_part(str(out))
    print(f"已写入：{out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
