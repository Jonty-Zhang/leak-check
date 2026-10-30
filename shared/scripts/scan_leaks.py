#!/usr/bin/env python3
"""扫一遍论文 / 课程作业 / 报告，找出交上去会露馅或一眼不对的句子。

    scan_leaks.py 论文.docx [更多文件 ...] [--json] [--ignore CODE ...] [--format md] [--no-style]
    scan_leaks.py 论文.docx --dump      # 按同一套段号输出全文，供通读和 fix_docx.py 定位
    cat 草稿.txt | scan_leaks.py -

支持 .docx .md .txt .tex .pdf（pdf 需要 pdftotext）。'-' 表示从标准输入读纯文本。

两类问题：
  雷点  元话语（根据用户的要求 / 用户说 / 老师要求）、草稿残留（之后要替换 / 待补充 / TODO）、
        版本标记（改进版 / 最终版）、对话残留（以下是 / 希望对你有帮助）、AI 自曝、
        Markdown / LaTeX 残留、Word 批注与未接受的修订
  硬伤  图表编号断裂或重复、正文引用了不存在的图表、引用编号超出参考文献表、
        “错误！未找到引用源”、标题编号跳号、空章节、重复段落、括号不配对、前后数量不一致
  AI 味 套话词汇、“综上所述”泛滥、三段式排比、节末升华句、让步式自贬、破折号连用、
        空洞的乐观结尾。只报 WARN，按全文出现次数和篇幅给阈值，同类合成一条；--no-style 关掉

ERROR = 基本可以确定要改；WARN = 要结合上下文判断。
退出码：有 ERROR 为 1，否则 0；文件读不了为 2。

脚本只查得出能靠字符串和结构判断的部分，语义上的矛盾要配合
references/semantic-checklist.md 人工（或 AI）通读。
"""
from __future__ import annotations

import argparse
import datetime
import json
import re
import shutil
import subprocess
import sys
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

ERROR, WARN = "ERROR", "WARN"

CATEGORY = {
    "META": "元话语",
    "DRAFT": "草稿残留",
    "VERSION": "版本标记",
    "CHAT": "对话残留",
    "AI": "AI 自曝",
    "MD": "格式残留",
    "BROKEN": "坏掉的引用",
    "NUM": "编号问题",
    "CITE": "参考文献",
    "STRUCT": "结构问题",
    "DOCX": "Word 残留",
    "STYLE": "AI 味",
}

SUGGEST = {
    "META": "删去这半句，直接陈述做了什么。例：“根据用户的要求，本文采用问卷法”→“本文采用问卷法”",
    "META_TEACHER": "读者就是老师，不用转述要求。改成陈述任务本身：“本实验测量……”",
    "META_SELFCERT": "删掉。字数、格式是否达标由读者判断，不写进正文",
    "DRAFT": "把这里的内容补齐；补不了就删掉这一处，并告诉作者缺什么",
    "DRAFT_OMIT": "补全内容，或者写明省略的理由和出处（如“推导见文献[3]”）",
    "DRAFT_FAKE": "示例或虚构数据不能出现在成品里：换成真实数据，或者删掉并告诉作者",
    "VERSION": "删掉版本标记，只留正文",
    "VERSION_DIFF": "这是改稿时的对照说明，删掉，只留改后的正文",
    "CHAT": "聊天口吻，整句删掉",
    "AI": "整句删掉；需要的话改写成作者自己的话",
    "MD": "转成正式格式：标题用标题样式，加粗用字体加粗，列表用编号，表格用原生表格",
    "LATEX": "改成 Word 公式（或 pandoc 转换），不要留 LaTeX 源码",
    "BROKEN_FIELD": "交叉引用断了：在 Word 里重新插入交叉引用，或改成正确的编号后更新域（全选按 F9）",
    "BROKEN_MARK": "补上真实的引用或编号；找不到出处就删掉这句话",
    "STYLE_WORD": "删掉套话，或换成它背后的具体内容：“具有重要意义”→说出意义是什么；“深入探讨”→“分析”；“值得注意的是”→直接说那件事",
    "STYLE_SUMUP": "全文留一处（放在结论开头），其余删掉这几个字，直接写结论",
    "STYLE_TRIAD": "不必凑满三项：改成实际的两项，或者给每一项说出具体内容、数值",
    "STYLE_ELEVATE": "删掉节末的升华句，一节讲完就停；意义留到结论里具体地写",
    "STYLE_CONCEDE": "正面写选择和理由：“更好的做法是 X，但为了简便采用 Y”→“采用 Y，原因是……”；“水平有限，难免疏漏”整句删",
    "STYLE_DASH": "大多数破折号改成逗号、冒号或句号，全文留两三处",
    "STYLE_ENDING": "把空泛的展望换成下一步具体要做的事；没有就删掉这句",
}


@dataclass
class Unit:
    n: int                 # 第几段 / 第几行（1 起）
    loc: str               # 给人看的位置
    text: str
    kind: str = "para"     # para / heading / caption / table / code / meta
    level: int = 0         # 标题级别
    in_refs: bool = False  # 位于参考文献表内


@dataclass
class Finding:
    severity: str
    code: str
    category: str
    loc: str
    excerpt: str
    problem: str
    suggestion: str = ""
    n: int = 0


@dataclass
class Doc:
    path: str
    fmt: str
    units: list = field(default_factory=list)
    findings: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    docx_zip: object = None


# ---------------------------------------------------------------- 读取

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(tag):
    return f"{{{W_NS}}}{tag}"


def iter_docx_paragraphs(document):
    """按文档顺序给出正文段落（含表格里的段落）：(段号, w:p 元素, 是否在表格内)。

    leak-fix 的 fix_docx.py 用同一套编号，scan 报“第 N 段”，fix 就能按第 N 段改。
    """
    n = 0

    def walk(el, in_table):
        nonlocal n
        for child in el.iterchildren():
            if child.tag == _q("p"):
                n += 1
                yield n, child, in_table
            elif child.tag == _q("tbl"):
                for tr in child.iter(_q("tr")):
                    for tc in tr.iterchildren(_q("tc")):
                        yield from walk(tc, True)
            elif child.tag == _q("sdt"):
                content = child.find(_q("sdtContent"))
                if content is not None:
                    yield from walk(content, in_table)

    yield from walk(document.element.body, False)


M_T = "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"


def p_text(p):
    """段落文字，含 Word 公式（OMML）里的文字。"""
    return "".join(el.text or "" for el in p.iter(_q("t"), M_T))


def _docx_style_info(document):
    """styleId → (名称, 标题级别, 是否题注, 是否代码)"""
    info = {}
    for st in document.styles:
        try:
            sid, name = st.style_id, (st.name or "")
        except Exception:
            continue
        low = name.lower()
        level = 0
        m = re.match(r"(heading|标题)\s*(\d)", low)
        if m:
            level = int(m.group(2))
        elif low in ("title", "标题"):
            level = 1
        caption = low in ("caption", "题注") or "caption" in low
        code = any(k in low for k in ("code", "代码", "source", "verbatim"))
        info[sid] = (name, level, caption, code)
    return info


MONO = ("consolas", "courier", "menlo", "monaco", "source code", "jetbrains", "fira code", "sf mono", "cascadia")


def read_docx(path, doc):
    try:
        from docx import Document
    except ImportError:
        sys.exit("需要 python-docx：pip install python-docx")
    document = Document(path)
    styles = _docx_style_info(document)
    for n, p, in_table in iter_docx_paragraphs(document):
        text = p_text(p)
        ppr = p.find(_q("pPr"))
        sid = None
        level = 0
        if ppr is not None:
            ps = ppr.find(_q("pStyle"))
            if ps is not None:
                sid = ps.get(_q("val"))
            ol = ppr.find(_q("outlineLvl"))
            if ol is not None and ol.get(_q("val"), "").isdigit():
                level = int(ol.get(_q("val"))) + 1
        name, slevel, caption, code = styles.get(sid, ("", 0, False, False))
        level = slevel or level
        fonts = [f.lower() for f in (rf.get(_q("ascii")) or "" for rf in p.iter(_q("rFonts"))) if f]
        if fonts and all(any(m in f for m in MONO) for f in fonts):
            code = True
        kind = "table" if in_table else "heading" if level else "caption" if caption else "code" if code else "para"
        doc.units.append(Unit(n, f"第 {n} 段" + ("（表格内）" if in_table else ""), text, kind, level))
    # 页眉页脚
    for si, section in enumerate(document.sections, 1):
        for label, part in (("页眉", section.header), ("页脚", section.footer)):
            try:
                if part.is_linked_to_previous and si > 1:
                    continue
                for p in part.paragraphs:
                    if p.text.strip():
                        doc.units.append(Unit(0, f"{label}（第 {si} 节）", p.text, "meta"))
            except Exception:
                pass
    doc.docx_zip = zipfile.ZipFile(path)


def _lines_to_units(text, doc, fmt):
    in_code = False
    in_verbatim = False
    for i, raw in enumerate(text.splitlines(), 1):
        line = raw.rstrip("\n")
        kind, level = "para", 0
        if fmt == "md":
            if re.match(r"^\s*(```|~~~)", line):
                in_code = not in_code
                kind = "code"
            elif in_code or line.startswith("    ") and line.strip():
                kind = "code"
            else:
                m = re.match(r"^(#{1,6})\s+(.*)$", line)
                if m:
                    kind, level = "heading", len(m.group(1))
                elif re.match(r"^\s*\|.*\|\s*$", line):
                    kind = "table"
        elif fmt == "tex":
            if re.search(r"\\begin\{(verbatim|lstlisting|minted|Verbatim)\}", line):
                in_verbatim = True
            if in_verbatim or re.match(r"^\s*%", line):
                kind = "code"
            if re.search(r"\\end\{(verbatim|lstlisting|minted|Verbatim)\}", line):
                in_verbatim = False
            m = re.match(r"^\s*\\(chapter|section|subsection|subsubsection)\*?\{", line)
            if m and kind != "code":
                kind, level = "heading", ("chapter", "section", "subsection", "subsubsection").index(m.group(1)) + 1
        doc.units.append(Unit(i, f"第 {i} 行", line, kind, level))


def read_pdf(path, doc):
    exe = shutil.which("pdftotext")
    if not exe:
        sys.exit("读 PDF 需要 pdftotext（brew install poppler）")
    out = subprocess.run([exe, "-enc", "UTF-8", path, "-"], capture_output=True, text=True, check=True).stdout
    n = 0
    for pi, page in enumerate(out.split("\f"), 1):
        for li, line in enumerate(page.splitlines(), 1):
            n += 1
            doc.units.append(Unit(n, f"第 {pi} 页第 {li} 行", line))


def load(path, fmt=None):
    if path == "-":
        doc = Doc("<stdin>", fmt or "txt")
        _lines_to_units(sys.stdin.read(), doc, doc.fmt)
        return doc
    p = Path(path)
    fmt = fmt or p.suffix.lower().lstrip(".")
    fmt = {"markdown": "md", "text": "txt", "latex": "tex"}.get(fmt, fmt)
    doc = Doc(str(p), fmt)
    if fmt == "docx":
        read_docx(str(p), doc)
    elif fmt == "pdf":
        read_pdf(str(p), doc)
    elif fmt in ("md", "txt", "tex"):
        _lines_to_units(p.read_text(encoding="utf-8", errors="replace"), doc, fmt)
    else:
        sys.exit(f"不支持的格式：{p.suffix}（支持 .docx .md .txt .tex .pdf）")
    return doc


# ---------------------------------------------------------------- 短语规则

VER = r"(改进版|最终版|修改版|修订版|优化版|完善版|润色版|精简版|扩写版|改写版|升级版|新版本|定稿|终稿|初稿|草稿)"

# (代码, 类别, 级别, 正则, 问题说明, 建议键, 选项)
# 选项：user = 文档把“用户”当领域词时降为 WARN；i = 忽略大小写
RULES = [
    # ---- 元话语
    ("META_USER", "META", ERROR,
     r"(根据|按照|按|依照|遵照|遵循|应|依据|满足)(了)?用户(的|所)?(要求|需求|指示|指令|意见|反馈|说明|描述|提供|偏好|想法)",
     "转述“用户”的要求，暴露了这是按指令写的", "META", "user"),
    ("META_USER", "META", ERROR,
     r"用户(刚才|之前|上面|前面|已经|曾)?(说|提到|指出|表示|希望|想要|要求|强调|补充|反馈|给出|提供|上传|发来|询问|问到|问|让我|希望我|让)(了|的|过|道)?",
     "正文里出现“用户说/用户希望……”，读起来是在复述聊天记录", "META", "user"),
    ("META_USER", "META", ERROR,
     r"(由于|因为)?用户(未|没有|没)(提供|给出|说明|上传)",
     "把“用户没给材料”写进了正文", "META", "user"),
    ("META_YOU", "META", ERROR,
     r"(根据|按照|按|依照|遵照|如)(你|您)(的|所)?(要求|需求|指示|说的|提供|描述|意见|反馈|想法)|(你|您)(提供|上传|给出|发来|给)的(资料|材料|数据|文件|内容|信息|图片|文本|代码)",
     "对“你/您”说话，是对话残留", "META", ""),
    ("META_TEACHER", "META", ERROR,
     r"(根据|按照|按|依照|遵照|满足|应)?(老师|教师|任课老师|助教|导师)(的|所)?(要求|规定|布置的要求)",
     "转述老师的要求；读者就是老师", "META_TEACHER", ""),
    ("META_TEACHER", "META", WARN,
     r"(老师|教师|任课老师|助教)(说|提到|强调|希望|让我们|叫我们|布置)",
     "转述老师的话；心得、反思类文章里可能是正常叙述，其他文体一般要删", "META_TEACHER", ""),
    ("META_TASK", "META", ERROR,
     r"(根据|按照|按|依照|满足|符合)?(课堂|课程|作业|题目|本次作业|本题)(的)?(要求|规定)(?!分析)|题目要求(我们|本文|学生)|作业要求(我们|本文|学生)",
     "转述作业/题目要求，读起来像在交差", "META_TEACHER", ""),
    ("META_TASK", "META", WARN,
     r"(根据|按照|按|依照)(实验|任务|项目)(的)?(要求|指导书要求)",
     "转述实验/任务要求；实验报告里偶尔可以，最好直接写做了什么", "META_TEACHER", ""),
    ("META_PROMPT", "META", ERROR,
     r"上一轮(对话|回复|回答)|(之前|上面|刚才)的对话|聊天记录(中|里)|(根据|针对|回答|关于|按照)(你|您)的(提问|问题|消息|指令)",
     "提到了对话过程", "META", ""),
    ("META_PROMPT", "META", WARN,
     r"提示词|(?<![A-Za-z])prompt(?![A-Za-z])",
     "出现“提示词/prompt”；写大模型相关论文时正常，否则多半是残留", "META", "i"),
    ("META_SELFCERT", "META", ERROR,
     r"字数(已|已经)?(达到|满足|超过|符合|够)|(全文|本文|正文|以上)(共计?|约|总计|合计)\s*\d{3,6}\s*字|(已|均已|已经)(满足|符合|覆盖|完成)(了)?(全部|所有|上述|作业|题目|格式|评分)?(的)?(要求|格式要求|评分标准)",
     "在正文里自证字数/要求达标", "META_SELFCERT", ""),
    ("META_SELFCERT", "META", WARN,
     r"不少于\s*\d{3,6}\s*字|\d{3,6}\s*字(以上|左右)的(要求|篇幅)",
     "正文里提到字数要求", "META_SELFCERT", ""),

    # ---- 草稿残留
    ("DRAFT_TODO", "DRAFT", ERROR,
     r"(?<![A-Za-z])(TODO|TBD|FIXME)(?![A-Za-z])|lorem ipsum|placeholder",
     "占位标记没删", "DRAFT", "i"),
    ("DRAFT_TODO", "DRAFT", ERROR,
     r"待(补充|填写|填入|完善|确认|核实|修改|插入|更新|补(?!偿)|替换|添加|整理)|[（(【\[]\s*待定\s*[）)】\]]",
     "“待补充/待确认”一类的草稿标记", "DRAFT", ""),
    ("DRAFT_REPLACE", "DRAFT", ERROR,
     r"(之后|稍后|后续|后面|以后|回头|最后|到时)(再)?(要|需要|会|可以|记得|请)?(再)?(把)?(这里|此处|这部分|这段|该处|本段|这些|上面|下面|它)?(的)?(内容|数据|数值|图片|截图|文字)?(替换|换成|换掉|改掉|填上|填入|补上)",
     "“之后要替换……”是写作过程的备注", "DRAFT", ""),
    ("DRAFT_REPLACE", "DRAFT", ERROR,
     r"(替换|换成|换掉|填入|填上|补充|补上|改成|修改|核实|确认|删掉|删除)(掉)?(这里|此处|这部分|该处|本段|这段)(的)?",
     "让人去改“这里”的内容，是草稿备注", "DRAFT", ""),
    ("DRAFT_REPLACE", "DRAFT", ERROR,
     r"(这里|此处|该处|本段|这段)(的)?(内容|数据|图片|截图|文字|数值|公式|图|表|代码)?(需要|要|请|应|待|可以|可|稍后|之后|后续|先)?(再)?(替换|填写|填入|插入|放置|粘贴|更新|核实|删除|补上|换成)",
     "“此处插入图片/这里先用……”一类的草稿备注", "DRAFT", ""),
    ("DRAFT_REPLACE", "DRAFT", ERROR,
     r"请(替换|填写|补充|插入|粘贴|根据实际|自行|按需|在此)|替换为(你|您|实际|真实|自己)|(根据|按照|结合)(实际|自己的|你的|您的)情况(修改|调整|填写|替换|补充)|在此(插入|填写|放置|粘贴|添加)",
     "写给作者看的填空说明留在了正文里", "DRAFT", ""),
    ("DRAFT_REPLACE", "DRAFT", WARN,
     r"这里先(用|写|放|假设)|(暂时|暂且|先)(用|写|放)(一下|着)",
     "像是临时写法，确认后面没有要替换的", "DRAFT", ""),
    ("DRAFT_PLACEHOLDER", "DRAFT", ERROR,
     r"【[^】]{0,30}(填写|插入|补充|替换|此处|这里|待|图片|截图|姓名|学号|日期)[^】]{0,30}】|[\[［](此处|这里|在此|插入|填写|图片|截图|待)[^\]］]{0,30}[\]］]|[（(](此处|这里|在此)(插入|放|填|补|为|是|应)[^）)]{0,30}[）)]",
     "方括号占位符", "DRAFT", ""),
    ("DRAFT_PLACEHOLDER", "DRAFT", ERROR,
     r"(?<![A-Za-z])X{3,}(?![A-Za-z])|×{2,}|(学号|姓名|班级|日期|学院|专业)[:：]\s*(X+|x+|_+|某某|待填|\s*)$",
     "XXX / ×× 一类的占位符", "DRAFT", ""),
    ("DRAFT_PLACEHOLDER", "DRAFT", WARN,
     r"某某|(?<![A-Za-z])XX(?![A-Za-z])",
     "“某某/XX”：举例时可以，若是没填的空就要补", "DRAFT", ""),
    ("DRAFT_OMIT", "DRAFT", ERROR,
     r"[（(]\s*略\s*[）)]|内容同上|此处略$|(其余|剩下|其他)(的)?(部分|内容|代码|步骤)(同理|类似|省略|略)",
     "内容被省略了", "DRAFT_OMIT", ""),
    ("DRAFT_OMIT", "DRAFT", WARN,
     r"(此处省略|(限于|由于)篇幅.{0,6}(省略|略去|不再))(?!.{0,20}(详见|参见|见文献|见附录|见参考))",
     "省略内容；论文里可以，但要写明出处", "DRAFT_OMIT", ""),
    ("DRAFT_FAKE", "DRAFT", ERROR,
     r"(示例|虚构|假设|假的|编造|随机生成|模拟的?示例)(的)?数据(?!集成)|(以下|上述|本表|表中|图中)(数据|数值)(为|是|均为)(示例|虚构|假设|示意)|数据(仅|仅为|只是)(示例|示意|参考)",
     "声明了数据是示例/虚构的", "DRAFT_FAKE", ""),

    # ---- 版本标记
    ("VERSION_MARK", "VERSION", ERROR,
     rf"[（(【\[]\s*{VER}\s*[）)】\]]",
     "括号里的版本标记", "VERSION", ""),
    ("VERSION_MARK", "VERSION", ERROR,
     rf"(以下是|下面是|这是|如下为|下为|附上|给出|这里是|这版是|这一版是?)(.{{0,10}}){VER}",
     "“以下是改进版……”", "VERSION", ""),
    ("VERSION_MARK", "VERSION", ERROR,
     rf"^\W*{VER}\s*[:：]|{VER}(论文|报告|作业|正文|全文|文章|内容|稿件)?\s*[:：]\s*$",
     "版本标签", "VERSION", ""),
    ("VERSION_DIFF", "VERSION", ERROR,
     r"^\W*(修改前|修改后|原文|原句|原稿|改写后|润色后|优化后|改后|改前|修订说明|修改说明|改动说明|更新说明|修改要点|主要修改)\s*[:：]",
     "改稿对照残留", "VERSION_DIFF", ""),
    ("VERSION_DIFF", "VERSION", ERROR,
     r"[（(【\[](新增|已新增|修改|已修改|已更新|已润色|已删除|已补充|已调整|改动|重写|修改处|此处有改动|新增内容|修改内容)[）)】\]]",
     "改动标注残留", "VERSION_DIFF", ""),

    # ---- 对话残留
    ("CHAT_OPENER", "CHAT", ERROR,
     r"^\W*(好的|当然可以|当然|没问题|可以的|明白了|收到|好嘞)[，,！!。~]",
     "聊天的开场白", "CHAT", ""),
    ("CHAT_INTRO", "CHAT", ERROR,
     r"(以下是|下面是|下面为|如下是|这是)(.{0,20})(修改后|改进后|润色后|优化后|完整的|重写|扩写|为您|为你|你要的|您要的|按照|根据)",
     "“以下是为你修改后的……”", "CHAT", ""),
    ("CHAT_INTRO", "CHAT", WARN,
     r"^\W*(以下是|下面是|下面给出|接下来给出|如下是)",
     "“以下是……”开头，口气像回答问题；论文里直接写内容", "CHAT", ""),
    ("CHAT_CLOSER", "CHAT", ERROR,
     r"希望(这|以上|本文|这些|这份|这篇)?.{0,8}(对你|对您)(有|有所)?(帮助|用|启发)|(如果|如|若)(你|您)(还)?(需要|有任何|想要|希望|觉得|对)|(需要|要)我(再|帮你|帮您|为你|为您|继续|进一步)|(你|您)可以(根据|按照|结合)(需要|实际|自己|具体)|请(告诉|让)我|(随时|欢迎)(告诉|联系|提问|反馈)|有(任何)?问题(可以)?(随时)?(再)?(问我|找我|告诉我)",
     "聊天的收尾或客套话", "CHAT", ""),
    ("CHAT_EN", "CHAT", ERROR,
     r"\b(I hope this helps|hope this helps|let me know if|feel free to|here is (the|a|your) (revised|improved|updated|final|polished|rewritten)|here's (the|a|your) (revised|improved|updated|final)|certainly!|sure!|as you requested|per your request|as requested by the user|the user (wants|wanted|asked|requested|said|mentioned|provided)|according to the user)\b",
     "英文聊天残留", "CHAT", "i"),

    # ---- AI 自曝
    ("AI_SELF", "AI", ERROR,
     r"作为(一个|一名)?(AI|人工智能|语言模型|大模型|大语言模型|AI ?助手|智能助手)|(我是|我只是|我作为)(一个)?(AI|人工智能|语言模型|大语言模型)|我的(训练数据|知识截止)|知识截止(日期|时间)|截至我的(知识|训练)",
     "AI 自报身份", "AI", "i"),
    ("AI_SELF", "AI", ERROR,
     r"(抱歉|对不起|很遗憾)[，,]?\s*(我)?(无法|不能|没法|不可以)|我(无法|不能)(访问|查看|浏览|打开|获取|读取|联网)",
     "AI 的拒绝或道歉", "AI", ""),
    ("AI_SELF", "AI", WARN,
     r"(ChatGPT|GPT-?\d|Claude|文心一言|通义千问|Kimi|豆包|DeepSeek|AI)(帮我|为我|帮忙)?(生成|撰写|写的|输出|润色)",
     "提到用 AI 生成；研究大模型的论文里正常，否则是自曝", "AI", "i"),
    ("AI_EN", "AI", ERROR,
     r"\b(as an AI|as a language model|as an AI language model|my knowledge cutoff|my training data|I (cannot|can't|am unable to) (browse|access))\b",
     "英文 AI 自曝", "AI", "i"),

    # ---- 坏掉的引用
    ("BROKEN_FIELD", "BROKEN", ERROR,
     r"错误[！!]\s*(未找到引用源|书签未定义|文档中没有指定样式的文字|未定义书签)|Error!\s*(Reference source not found|Bookmark not defined|No text of specified style)",
     "Word 交叉引用断了，打印出来就是这行字", "BROKEN_FIELD", "i"),
    ("BROKEN_MARK", "BROKEN", ERROR,
     r"[\[［](citation needed|需要引用|引用|来源|文献|ref|cite|待引)[\]］]|(?<![A-Za-z0-9_\]\)])[\[［]\s*\?\s*[\]］]|(?<![A-Za-z0-9_\]\)\s])[\[［]\s*[\]］](?!\s*[=;,)])",
     "空的或待填的引用标记", "BROKEN_MARK", "i"),
]

TEX_BROKEN = ("BROKEN_MARK", "BROKEN", ERROR,
              r"(图|表|式|公式|节|章|Figure|Fig\.|Table|Eq\.?|Equation|Section)\s*~?\?\?|\?\?(?=\s*(所示|中|节|章))",
              "LaTeX 交叉引用没解析（显示成 ??）", "BROKEN_MARK", "i")

MD_RULES = [
    (r"^\s*#{1,6}\s+[\u4e00-\u9fa5A-Za-z0-9][^。！？；!?;，,<>=*+()]{0,40}$", "行首 # 号（Markdown 标题）"),
    (r"\*\*[^*\n]+\*\*", "**加粗** 标记"),
    (r"```|~~~", "代码围栏"),
    (r"`[^`\n]+`", "反引号"),
    (r"^\s*\|.*\|\s*$", "Markdown 表格"),
    (r"^\s*-{3,}\s*$|^\s*\*{3,}\s*$", "--- 分隔线"),
    (r"!?\[[^\]\n]+\]\((https?://|\.{0,2}/|[\w\-]+\.(png|jpg|jpeg|gif|svg))", "Markdown 链接或图片语法"),
    (r"^\s*>\s+\S", "> 引用符号"),
]
MD_LIST_RULE = (r"^\s*[-*]\s+[\u4e00-\u9fa5A-Za-z]", "- 列表符号")
LATEX_RULES = [
    (r"\$\$|\\\[|\\\]|\\\(|\\\)", "LaTeX 公式定界符"),
    (r"\$[^$\n]{1,80}\$", "$…$ 行内公式源码"),
    (r"\\(frac|alpha|beta|gamma|delta|theta|lambda|sigma|omega|sum|int|sqrt|mathrm|mathbf|text|times|cdot|leq|geq|infty|partial|begin|end)\b", "LaTeX 命令"),
]

COMPILED = []
for code, cat, sev, pattern, problem, skey, opts in RULES:
    flags = re.MULTILINE | (re.IGNORECASE if "i" in opts else 0)
    COMPILED.append((code, cat, sev, re.compile(pattern, flags), problem, skey, opts))


def excerpt(text, start, end, width=70):
    """取命中位置所在的句子，太长就以命中处为中心截一段。"""
    left = max(text.rfind(c, 0, start) for c in "。！？!?；;\n") + 1
    rights = [i for i in (text.find(c, end) for c in "。！？!?；;\n") if i != -1]
    right = min(rights) + 1 if rights else len(text)
    s = text[left:right].strip()
    if len(s) <= width:
        return s
    mid = start - left
    a = max(0, mid - width // 3)
    return ("…" if a else "") + s[a:a + width] + "…"


def add(doc, sev, code, cat, unit_or_loc, text, problem, skey="", n=0):
    if isinstance(unit_or_loc, Unit):
        loc, n = unit_or_loc.loc, unit_or_loc.n
    else:
        loc = unit_or_loc
    doc.findings.append(Finding(sev, code, CATEGORY[cat], loc, text, problem, SUGGEST.get(skey, ""), n))


# ---------------------------------------------------------------- 参考文献区

REF_HEAD = re.compile(r"^\W*(参考文献|参考资料|引用文献|References|Bibliography|Works Cited)\W*$", re.I)
AFTER_REFS = re.compile(r"^\W*(附录|致谢|致 谢|Appendix|Acknowledg)", re.I)
REF_ENTRY = re.compile(r"^\s*[\[［]?\s*(\d{1,3})\s*[\]］.、．]\s*\S")


def mark_refs(doc):
    start = None
    for i, u in enumerate(doc.units):
        if REF_HEAD.match(u.text.strip()) and len(u.text.strip()) <= 30:
            start = i
    if start is None:
        return None
    for u in doc.units[start + 1:]:
        if AFTER_REFS.match(u.text.strip()) and (u.kind == "heading" or len(u.text.strip()) <= 20):
            break
        u.in_refs = True
    return start


# ---------------------------------------------------------------- 短语检查

def user_domain_count(doc):
    """“用户”作为领域词（需求分析、产品设计里）出现的段数；越多越要谨慎。"""
    user_rules = [c[3] for c in COMPILED if "user" in c[6]]
    cnt = 0
    for u in doc.units:
        if "用户" in u.text and u.kind != "code" and not any(r.search(u.text) for r in user_rules):
            cnt += 1
    return cnt


SELF_REF = re.compile(r"本文|本报告|报告|论文|这份|文中|作业|本研究|本节|本章")


def check_phrases(doc, ignore):
    domain = user_domain_count(doc)
    if domain >= 3:
        doc.notes.append(f"全文有 {domain} 段把“用户”当领域词用（如需求分析），“用户……”类命中已降为 WARN，要逐条看上下文")
    rules = list(COMPILED)
    if doc.fmt in ("pdf", "docx"):
        c = TEX_BROKEN
        rules.append((c[0], c[1], c[2], re.compile(c[3], re.I), c[4], c[5], ""))
    for u in doc.units:
        text = u.text
        if not text.strip():
            continue
        for code, cat, sev, rx, problem, skey, opts in rules:
            if code in ignore:
                continue
            if u.in_refs and cat in ("VERSION", "DRAFT") and code != "DRAFT_TODO":
                continue  # 参考文献里的“第 2 版”“修订版”是书名信息
            for m in rx.finditer(text):
                s = sev
                prob = problem
                ex = excerpt(text, m.start(), m.end())
                if "user" in opts and domain >= 3:
                    if SELF_REF.search(ex):
                        prob = problem + "（同一句还提到了报告/论文本身，基本可以确定是在转述对话）"
                    else:
                        s, prob = WARN, problem + "（本文把“用户”当领域词用，要看上下文）"
                if u.kind == "code":
                    s, prob = WARN, problem + "（在代码块里，确认是不是代码本身的内容）"
                add(doc, s, code, cat, u, ex, prob, skey)
                break  # 同一条规则在一段里只报一次


def check_short_version_lines(doc, ignore):
    """标题或短行就是版本名，如“改进版”“最终版论文”。"""
    if "VERSION_MARK" in ignore:
        return
    rx = re.compile(VER)
    for u in doc.units:
        t = re.sub(r"[#*_`>\s（）()【】\[\]：:，,。.!！-]", "", u.text)
        if u.in_refs or u.kind == "code" or not t or len(t) > 12:
            continue
        if rx.search(t) and not any(f.n == u.n and f.code == "VERSION_MARK" and f.loc == u.loc for f in doc.findings):
            add(doc, ERROR, "VERSION_MARK", "VERSION", u, u.text.strip(), "单独一行的版本名", "VERSION")


def check_format_residue(doc, ignore):
    if doc.fmt == "md" or "MD_RESIDUE" in ignore:
        return
    md_rules = [(re.compile(p, re.M), d) for p, d in MD_RULES]
    if doc.fmt == "docx":
        md_rules.append((re.compile(MD_LIST_RULE[0], re.M), MD_LIST_RULE[1]))
    if doc.fmt == "tex":
        md_rules = [r for r in md_rules if r[1] in ("行首 # 号（Markdown 标题）", "**加粗** 标记", "代码围栏", "Markdown 表格")]
    latex_rules = [(re.compile(p), d) for p, d in LATEX_RULES] if doc.fmt in ("docx", "pdf", "txt") else []
    for u in doc.units:
        if u.kind == "code" or not u.text.strip():
            continue
        for rx, desc in md_rules:
            if u.kind == "table" and (desc.startswith("行首 #") or desc == MD_LIST_RULE[1]):
                continue  # 表格单元格里不会有标题
            m = rx.search(u.text)
            if m:
                sev = WARN if doc.fmt == "txt" or desc == MD_LIST_RULE[1] else ERROR
                add(doc, sev, "MD_RESIDUE", "MD", u, excerpt(u.text, m.start(), m.end()), f"{desc}没转换", "MD")
                break
        for rx, desc in latex_rules:
            m = rx.search(u.text)
            if m:
                add(doc, ERROR, "LATEX_RESIDUE", "MD", u, excerpt(u.text, m.start(), m.end()), f"{desc}没转换成公式", "LATEX")
                break


# ---------------------------------------------------------------- 图表编号

NUM = r"(\d+(?:\s*[.\-－–]\s*\d+)?)"
CAPTION_RX = re.compile(rf"^\s*(?:!\[)?[*_]*\s*(图|表|Figure|Fig\.|Table)\s*{NUM}(?![.\-－–]?\d)", re.I)
FIG_REF_RX = re.compile(rf"(?<![地意试企构蓝草插视版路略拼宏])(图|Figure|Fig\.)\s*{NUM}|(?<![代列发外报手年课时价数目图量])(表)\s*{NUM}", re.I)


def _norm_num(s):
    parts = re.split(r"\s*[.\-－–]\s*", s.strip())
    return tuple(int(x) for x in parts)


def _kind_of(word):
    return "图" if word.lower() in ("图", "figure", "fig.") else "表"


def check_numbering(doc, ignore):
    if doc.fmt == "tex":
        return  # LaTeX 自动编号
    caps = defaultdict(list)  # 类型 → [(编号, unit)]
    cap_units = set()
    for i, u in enumerate(doc.units):
        if u.in_refs or u.kind in ("code", "heading"):
            continue
        text = u.text
        m = re.match(r"^\s*!\[([^\]]*)\]", text)
        if m:
            text = m.group(1)
        m = CAPTION_RX.match(text)
        if not m:
            continue
        rest = text[m.end():]
        looks_caption = u.kind == "caption" or len(text.strip()) <= 60 or re.match(r"^\s*[\s:：.、　]", rest)
        if not looks_caption or re.match(r"^\s*(所示|中|可见|表明|显示|给出|列出|是|为|的)", rest):
            continue
        caps[_kind_of(m.group(1))].append((_norm_num(m.group(2)), u))
        cap_units.add(id(u))

    for kind, items in caps.items():
        if "NUM_CAPTION" in ignore:
            break
        seen = {}
        for num, u in items:
            if num in seen:
                add(doc, ERROR, "NUM_CAPTION", "NUM", u, u.text.strip()[:60],
                    f"{kind} {'-'.join(map(str, num))} 重复了（{seen[num].loc} 已有一个）")
            seen.setdefault(num, u)
        prev = None
        for num, u in items:
            if prev is not None:
                ok = False
                if len(num) == 1 and len(prev) == 1:
                    ok = num[0] == prev[0] + 1
                elif len(num) == 2 and len(prev) == 2:
                    ok = (num[0] == prev[0] and num[1] == prev[1] + 1) or (num[0] > prev[0] and num[1] == 1)
                else:
                    ok = num != prev
                if not ok and num != prev:
                    sev = ERROR if num > prev else WARN
                    add(doc, sev, "NUM_CAPTION", "NUM", u, u.text.strip()[:60],
                        f"{kind}号不连续：上一个是 {kind} {'-'.join(map(str, prev))}，这里是 {kind} {'-'.join(map(str, num))}")
            elif (len(num) == 1 and num[0] != 1) or (len(num) == 2 and num[1] != 1):
                add(doc, WARN, "NUM_CAPTION", "NUM", u, u.text.strip()[:60], f"第一个{kind}不是从 1 开始")
            prev = num

    if "NUM_FIGREF" in ignore:
        return
    refs = defaultdict(list)
    for u in doc.units:
        if id(u) in cap_units or u.in_refs or u.kind == "code":
            continue
        for m in FIG_REF_RX.finditer(u.text):
            word, num = (m.group(1), m.group(2)) if m.group(1) else (m.group(3), m.group(4))
            refs[_kind_of(word)].append((_norm_num(num), u, m))
    for kind, items in refs.items():
        have = {n for n, _ in caps.get(kind, [])}
        if not have:
            continue  # 没识别出题注（可能是图片里的文字），不做比对
        reported = set()
        for num, u, m in items:
            if num not in have and num not in reported:
                reported.add(num)
                add(doc, ERROR, "NUM_FIGREF", "NUM", u, excerpt(u.text, m.start(), m.end()),
                    f"正文提到了{kind} {'-'.join(map(str, num))}，但全文没有这个{kind}的题注")
        cited = {n for n, _, _ in items}
        uncited = [(num, u) for num, u in caps.get(kind, []) if num not in cited]
        if uncited:
            names = "、".join(f"{kind} {'-'.join(map(str, num))}" for num, _ in uncited[:12])
            more = f" 等 {len(uncited)} 个" if len(uncited) > 12 else ""
            add(doc, WARN, "NUM_FIGREF", "NUM", uncited[0][1], uncited[0][1].text.strip()[:60],
                f"{names}{more}在正文里没被提到过（很多老师要求写“如{kind} X 所示”）")


# ---------------------------------------------------------------- 标题编号与空章节

CN = "零一二三四五六七八九十"


def cn2int(s):
    if s.isdigit():
        return int(s)
    if s == "十":
        return 10
    if len(s) == 1:
        return CN.index(s)
    if s.startswith("十"):
        return 10 + CN.index(s[1])
    if s.endswith("十"):
        return CN.index(s[0]) * 10
    if len(s) == 3 and s[1] == "十":
        return CN.index(s[0]) * 10 + CN.index(s[2])
    return -1


HEAD_SCHEMES = [
    ("章", re.compile(r"^\s*第\s*([一二三四五六七八九十]+|\d+)\s*(章|部分)")),
    ("节", re.compile(r"^\s*第\s*([一二三四五六七八九十]+|\d+)\s*节")),
    ("一、", re.compile(r"^\s*([一二三四五六七八九十]+)\s*[、.．]")),
    ("（一）", re.compile(r"^\s*[（(]\s*([一二三四五六七八九十]+)\s*[)）]")),
    ("dot", re.compile(r"^\s*(\d+(?:\.\d+)*)(?:\s+|[、．]|\.(?!\d)|(?=[\u4e00-\u9fa5A-Za-z]))")),
]


def check_headings(doc, ignore):
    heads = [u for u in doc.units if u.kind == "heading" and u.text.strip()]
    if not heads:
        return
    if "NUM_HEADING" not in ignore and doc.fmt != "tex":
        last = {}  # (标题级别, 编号方案) → 上一个编号；上级标题出现时，下级重新计数
        for u in heads:
            t = re.sub(r"^#+\s*", "", u.text).strip()
            for key in [k for k in last if k[0] > u.level]:
                del last[key]
            for name, rx in HEAD_SCHEMES:
                m = rx.match(t)
                if not m:
                    continue
                key = (u.level, name)
                prev = last.get(key)
                if name == "dot":
                    num = tuple(int(x) for x in m.group(1).split("."))
                    if prev is not None and len(num) == len(prev):
                        ok = num[:-1] == prev[:-1] and num[-1] == prev[-1] + 1 or num[:-1] > prev[:-1] and num[-1] == 1
                        if not ok and num != prev:
                            add(doc, ERROR if num > prev else WARN, "NUM_HEADING", "NUM", u, t[:50],
                                f"标题编号接不上：同级上一个是 {'.'.join(map(str, prev))}，这里是 {'.'.join(map(str, num))}")
                        elif num == prev:
                            add(doc, ERROR, "NUM_HEADING", "NUM", u, t[:50], f"标题编号 {'.'.join(map(str, num))} 重复了")
                    last[key] = num
                else:
                    v = cn2int(m.group(1))
                    if prev is not None and v != prev + 1:
                        add(doc, ERROR if v > prev else WARN, "NUM_HEADING", "NUM", u, t[:50],
                            f"标题编号接不上：同级上一个“{name}”编号是 {prev}，这里是 {v}")
                    last[key] = v
                break
    if "EMPTY_SECTION" in ignore:
        return
    idx = {id(u): i for i, u in enumerate(doc.units)}
    for k, u in enumerate(heads):
        i = idx[id(u)]
        if REF_HEAD.match(u.text.strip("# ")):
            continue
        nxt_head = heads[k + 1] if k + 1 < len(heads) else None
        end = idx[id(nxt_head)] if nxt_head else len(doc.units)
        body = [x for x in doc.units[i + 1:end] if x.text.strip() and x.kind != "meta"]
        if body:
            continue
        if nxt_head is None or nxt_head.level <= u.level:
            add(doc, WARN, "EMPTY_SECTION", "STRUCT", u, re.sub(r"^#+\s*", "", u.text).strip()[:50],
                "这个标题下面没有内容（空章节）")


# ---------------------------------------------------------------- 参考文献

CITE_RX = re.compile(r"(?<![A-Za-z0-9_\]\)])[\[［]\s*(\d{1,3}(?:\s*[-–~～,，、]\s*\d{1,3})*)\s*[\]］](?!\()")


def _expand(s):
    out = []
    for part in re.split(r"\s*[,，、]\s*", s):
        if re.search(r"[-–~～]", part):
            a, b = (int(x) for x in re.split(r"\s*[-–~～]\s*", part))
            out.extend(range(a, b + 1) if 0 < b - a < 50 else [a, b])
        elif part.strip():
            out.append(int(part))
    return out


def check_citations(doc, ignore):
    start = mark_refs(doc)
    body = doc.units if start is None else doc.units[:start]
    cites = []
    for u in body:
        if u.kind == "code":
            continue
        for m in CITE_RX.finditer(u.text):
            for n in _expand(m.group(1)):
                cites.append((n, u, m))
    if start is None:
        if len(cites) >= 2 and "CITE_LIST" not in ignore:
            u = cites[0][1]
            add(doc, WARN, "CITE_LIST", "CITE", u, excerpt(u.text, cites[0][2].start(), cites[0][2].end()),
                f"正文有 {len(cites)} 处 [n] 式标记，但没找到“参考文献”标题和文献表；如果是文献引用，要补上文献表")
        return
    entries = []
    for u in doc.units[start + 1:]:
        if not u.in_refs:
            break
        m = REF_ENTRY.match(u.text)
        if m:
            entries.append((int(m.group(1)), u))
    if not entries:
        doc.notes.append("找到了“参考文献”标题，但条目不是 [1] / 1. 编号格式，没做引用比对")
        return
    nums = [n for n, _ in entries]
    if "CITE_LIST" not in ignore:
        prev = 0
        for n, u in entries:
            if n != prev + 1:
                add(doc, ERROR, "CITE_LIST", "CITE", u, u.text.strip()[:60],
                    f"文献编号接不上：上一条是 [{prev}]，这里是 [{n}]")
            prev = n
    if "CITE_RANGE" not in ignore:
        bad = set()
        for n, u, m in cites:
            if n not in nums and n not in bad:
                bad.add(n)
                add(doc, ERROR, "CITE_RANGE", "CITE", u, excerpt(u.text, m.start(), m.end()),
                    f"引用了 [{n}]，但参考文献表里只有 {len(nums)} 条（没有第 {n} 条）")
    if "CITE_UNUSED" not in ignore and cites:
        cited = {n for n, _, _ in cites}
        unused = [n for n in nums if n not in cited]
        if unused:
            u = dict(entries)[unused[0]]
            add(doc, WARN, "CITE_UNUSED", "CITE", u, u.text.strip()[:60],
                f"这些文献在正文里没被引用过：{', '.join(f'[{n}]' for n in unused[:15])}")
    if "CITE_ORDER" not in ignore and cites:
        order, seen = [], set()
        for n, u, m in cites:
            if n not in seen:
                seen.add(n)
                order.append((n, u, m))
        expect = 1
        for n, u, m in order:
            if n > expect:
                add(doc, WARN, "CITE_ORDER", "CITE", u, excerpt(u.text, m.start(), m.end()),
                    f"按顺序编码制，这里应首次出现 [{expect}]，实际先出现了 [{n}]（学校要求 GB/T 7714 顺序编码时要调）")
                break
            expect = max(expect, n + 1)
    if "REF_YEAR" not in ignore:
        this_year = datetime.date.today().year
        for n, u in entries:
            years = [int(y) for y in re.findall(r"(?<!\d)(1[89]\d\d|20\d\d)(?!\d)", u.text)]
            if not years:
                add(doc, WARN, "REF_YEAR", "CITE", u, u.text.strip()[:60], "这条文献没有年份，核实它是否真实存在、信息是否完整")
            elif min(years) > this_year:
                add(doc, ERROR, "REF_YEAR", "CITE", u, u.text.strip()[:60], f"文献年份 {min(years)} 比今年（{this_year}）还晚")


# ---------------------------------------------------------------- 其他结构

PAIRS = [("（", "）"), ("《", "》"), ("“", "”"), ("【", "】")]


def check_misc(doc, ignore):
    if "DUP_PARA" not in ignore:
        seen = {}
        for u in doc.units:
            if u.kind in ("code", "table", "heading", "meta") or u.in_refs:
                continue
            key = re.sub(r"\s+", "", u.text)
            if len(key) < 25 or len(re.findall(r"[\u4e00-\u9fa5]", key)) < 10:
                continue
            if key in seen:
                add(doc, WARN, "DUP_PARA", "STRUCT", u, u.text.strip()[:60], f"和{seen[key].loc}一字不差，像是重复粘贴")
            else:
                seen[key] = u
    if "UNBALANCED" not in ignore and doc.fmt != "pdf":
        reversed_quotes = []
        item_paren = re.compile(r"(?:^|[\s。；;，,？?！!：:、])\s*(\d{1,2}|[a-zA-Z]|[一二三四五六七八九十])）")
        for u in doc.units:
            if u.kind == "code" or not u.text.strip():
                continue
            for a, b in PAIRS:
                ca, cb = u.text.count(a), u.text.count(b)
                if a == "（" and ca < cb:
                    cb -= len(item_paren.findall(u.text))  # “1）2）3）”式编号只有右括号
                if ca != cb:
                    if a == "“" and ca == 0 and cb % 2 == 0:
                        reversed_quotes.append(u)
                    else:
                        add(doc, WARN, "UNBALANCED", "STRUCT", u, u.text.strip()[:60], f"“{a}”{ca} 个、“{b}”{cb} 个，没配对")
                    break
        if reversed_quotes:
            u = reversed_quotes[0]
            m = re.search("”[^”]*”", u.text)
            where = "、".join(x.loc for x in reversed_quotes[:10]) + (f" 等 {len(reversed_quotes)} 处" if len(reversed_quotes) > 10 else "")
            add(doc, WARN, "UNBALANCED", "STRUCT", u, excerpt(u.text, m.start(), m.end()),
                f"引号方向写反了（两边都是右引号 ”，前一个应为 “）：{where}")
    if "COUNT_MISMATCH" not in ignore:
        rx = re.compile(r"(共|共计|总共|总计|合计|一共)(收集|回收|有|获得|采集|选取|招募|调查|发放|纳入|涉及|包含|包括|访谈)?(了)?"
                        r"([\u4e00-\u9fa5]{0,4}?)\s*(\d+(?:\.\d+)?)\s*"
                        r"(名|个|条|份|组|例|人|张|篇|次|项|所|家|位|道|题|种|类|批|台|件|只|株)([\u4e00-\u9fa5]{0,3})")
        groups = defaultdict(list)
        for u in doc.units:
            if u.kind == "code" or u.in_refs:
                continue
            for m in rx.finditer(u.text):
                noun = (m.group(4) or m.group(7))[:2]
                key = (m.group(2) or "", m.group(6), noun)
                groups[key].append((m.group(5), u, m))
        for (verb, unit, noun), items in groups.items():
            vals = {v for v, _, _ in items}
            if len(vals) > 1 and (verb or noun):
                where = "；".join(f"{x.loc}“{excerpt(x.text, m.start(), m.end(), 30)}”" for _, x, m in items[:4])
                add(doc, WARN, "COUNT_MISMATCH", "STRUCT", items[1][1], excerpt(items[1][1].text, items[1][2].start(), items[1][2].end()),
                    f"同一个总数前后写得不一样（{' / '.join(sorted(vals))}）：{where}")


def check_docx_parts(doc, ignore):
    z = doc.docx_zip
    if z is None:
        return
    names = set(z.namelist())
    if "word/comments.xml" in names and "DOCX_COMMENT" not in ignore:
        xml = z.read("word/comments.xml").decode("utf-8", "replace")
        n = len(re.findall(r"<w:comment\b", xml))
        if n:
            first = re.sub(r"<[^>]+>", "", re.search(r"<w:comment\b.*?</w:comment>", xml, re.S).group(0))[:40]
            add(doc, ERROR, "DOCX_COMMENT", "DOCX", "整个文档", first,
                f"还有 {n} 条批注没删（审阅 → 删除 → 删除文档中的所有批注）")
    body = z.read("word/document.xml").decode("utf-8", "replace")
    if "DOCX_REVISION" not in ignore:
        ins = len(re.findall(r"<w:ins\b", body))
        dels = len(re.findall(r"<w:del\b", body))
        if ins or dels:
            m = re.search(r"<w:(ins|del)\b.*?</w:\1>", body, re.S)
            sample = re.sub(r"<[^>]+>", "", m.group(0))[:40] if m else ""
            add(doc, ERROR, "DOCX_REVISION", "DOCX", "整个文档", sample,
                f"还有未处理的修订（插入 {ins} 处、删除 {dels} 处），打开就是红字（审阅 → 接受所有修订）")
    if "DOCX_HIGHLIGHT" not in ignore:
        texts = []
        for r in re.findall(r"<w:r\b(?:(?!</w:r>).)*?<w:highlight\b(?:(?!</w:r>).)*?</w:r>", body, re.S):
            texts.append(re.sub(r"<[^>]+>", "", r))
        if texts:
            sample = "".join(texts)[:40]
            add(doc, WARN, "DOCX_HIGHLIGHT", "DOCX", "整个文档", sample,
                f"有 {len(texts)} 段文字带荧光高亮，常是改稿时标的“待改”，确认后去掉")
    if "DOCX_HIDDEN" not in ignore:
        n = len(re.findall(r"<w:vanish/>", body))
        if n:
            add(doc, WARN, "DOCX_HIDDEN", "DOCX", "整个文档", "", f"有 {n} 处隐藏文字，打印或别人打开“显示隐藏文字”时会露出来")


# ---------------------------------------------------------------- AI 味

# 单独一处都不算问题，“满篇都是”才是。所以全部只报 WARN，按全文次数过阈值才报，
# 同一类合成一条并列出所有位置。阈值 = 基数 + 每万字增加的量，见 style_limit。
# 清单和每一类“什么时候没问题”见 references/ai-style.md。

STYLE_WORDS = [
    # (显示名, 正则)
    ("至关重要/举足轻重/不可或缺", r"至关重要|举足轻重|不可或缺|不可忽视"),
    ("深入探讨/深入分析", r"深入(探讨|研究|分析|剖析|探究|挖掘)|深度(剖析|挖掘|解析)"),
    ("赋能/助力/抓手", r"赋能|助力|抓手"),
    ("彰显/凸显/标志着", r"彰显|凸显|标志着"),
    ("具有重要意义", r"具有(十分|非常|极其|极为)?(重大|重要|深远|深刻|积极)(的)?(现实|理论|实践|指导)?(意义|价值)|意义(重大|深远)"),
    ("奠定基础/有力支撑", r"奠定了?.{0,8}基础|(提供|给予)了?.{0,4}(有力|坚实|重要|可靠)(的)?(支撑|保障|支持)"),
    ("发挥重要作用", r"(扮演|发挥)(着|了)?.{0,4}(至关重要|举足轻重|关键|重要|核心|不可替代)(的)?(角色|作用)"),
    ("值得注意的是", r"值得(注意|一提|关注|强调)的是|需要(指出|强调|说明)的是|不难(发现|看出)"),
    ("众所周知/毋庸置疑", r"众所周知|毋庸置疑|不言而喻|显而易见"),
    ("全方位/多维度/深度融合", r"全方位|多维度|多层次|深度融合|有机结合|有机统一"),
    ("新篇章/新活力", r"新篇章|新征程|新高度|新活力|注入(新的?)?(活力|动力)|画卷|织就|宏伟蓝图"),
    ("不仅是……更是……", r"不(仅|只|单)(仅)?是.{1,24}?[，,]?\s*(更|而|也)是"),
    ("段首“此外/与此同时”", r"^\s*(此外|另外|与此同时)[，,]"),
]

SUMUP_RX = re.compile(r"综上所述|总而言之|总的来说|总之|总体而言|综合来看|一言以蔽之")
# 解题里的“综上所述，答案为 B”不算；题目解析类文档整篇不查这一项
SUMUP_ANSWER = re.compile(r"本题|答案|选[A-DＡ-Ｄ]|选项|[A-D]项|取值范围|[=＝<>≤≥∈]")
SOLUTION_MARK = re.compile(r"【(答案|解析|详解|分析|点睛|解答)】|故选[A-DＡ-Ｄ]|故答案为|^\s*解[:：]")

TRIAD_ADJ = (r"(高效|准确|精准|精确|稳定|可靠|便捷|快速|智能|安全|灵活|全面|系统|科学|规范|合理|完善|简洁|清晰|直观|"
             r"实用|新颖|丰富|深入|严谨|有效|先进|绿色|低碳|环保|开放|共享|协同|创新|多元|包容|友好|舒适|健康|和谐|"
             r"可持续|高质量|个性化|智能化|数字化|信息化|规范化|标准化|系统化|精细化|多样化|便捷化)")
TRIAD_PREFIX = (r"(更加|更为|更|进一步|不断|全面|持续|切实|提升了?|提高了?|增强了?|促进了?|推动了?|实现了?|培养了?|"
                r"锻炼了?|加深了?|拓宽了?|拓展了?|深化了?|丰富了?|激发了?|有助于|让|使)")
TRIAD_VERB = r"(锻炼|培养|加深|提升|提高|增强|拓宽|拓展|丰富|激发|深化|巩固|掌握|学会|收获|认识到|体会到|理解)了"
TRIAD_RULES = [
    re.compile(rf"{TRIAD_ADJ}[、，,]\s*(更加|更|且)?{TRIAD_ADJ}[、，,]?\s*(和|与|及|以及|且|又|而)?(更加|更)?{TRIAD_ADJ}"),
    re.compile(rf"{TRIAD_PREFIX}[\u4e00-\u9fa5]{{1,8}}[、，,；;]\s*\1[\u4e00-\u9fa5]{{1,8}}[、，,；;]?\s*(和|与|及|以及|并|还|更)?\1"),
    re.compile(rf"{TRIAD_VERB}[\u4e00-\u9fa5]{{1,10}}[、，,；;]\s*{TRIAD_VERB}[\u4e00-\u9fa5]{{1,10}}[、，,；;]\s*(也|还|并|更)?{TRIAD_VERB}"),
]

ELEVATE_RX = re.compile(
    r"(具有|有着|富有)(十分|非常|极其|极为)?(重大|重要|深远|深刻|积极|现实)?(的)?(现实|理论|实践|指导)?(意义|价值)|意义(重大|深远)"
    r"|(奠定|夯实)了?.{0,8}基础|(提供|带来|给予)了?.{0,10}(新的|有力|重要|坚实|宝贵|有益)?(的)?(思路|视角|启示|借鉴|支撑|保障|参考|方向)"
    r"|(彰显|凸显|体现|展现)(出|了)?.{0,14}(重要性|价值|意义|魅力|力量|精神|担当|优势|潜力)"
    r"|(推动|促进|助力|赋能|引领).{0,14}(发展|进步|转型|升级|未来|变革)|(必将|势必|终将|定将)"
    r"|未来可期|前景(广阔|光明|可观|无限)|不(仅|只)(仅)?是.{1,24}?(更|而)是|这(正|也)是.{0,16}(缩影|体现|写照|意义所在)")
ELEVATE_SKIP_HEAD = re.compile(r"结论|总结|结语|小结|展望|致谢|摘要|心得|体会|感想|收获|反思|Abstract|Conclusion|Summary", re.I)

CONCEDE_RULES = [
    (re.compile(r"(更好|更优|更理想|更合理|更准确|更严谨)的(做法|方法|方案|选择|处理|办法)(是|应该是|应当是|应是).{0,40}?(但|不过|然而|只是)"),
     "“更好的做法是……但……”"),
    (re.compile(r"理想(情况下|状态下|的做法|的方案)?.{0,30}?(应当|应该|应|需要|最好).{0,30}?(但|不过|然而|限于|由于|受限于)"),
     "“理想情况下应……但限于……”"),
    (re.compile(r"(为了|为|出于)(简便|简单|方便|省事|简化)(起见)?.{0,6}(本文|这里|此处|我们|笔者|本实验|本报告)?.{0,4}(只|仅|直接|暂且|选择|采用|使用)"
                r"(?=.{0,30}(而不是|而非|没有采用|未采用|没有使用|未使用))"),
     "“为了简便只……而不是……”"),
    (re.compile(r"严格(来说|地说|意义上)?.{0,20}?(应该|应当|需要|应).{0,30}?(这里|此处|本文|本实验)(只|仅|直接|近似|简化|粗略)"),
     "“严格来说应……这里只……”"),
    (re.compile(r"(虽然|尽管)(这种|该|本文的?|此|这个)?(方法|做法|方案|模型|处理)(比较|较为|很|相对|过于|有些)?(简单|粗糙|简陋|粗略|初级)"),
     "“这种方法虽然简单，但……”"),
    (re.compile(r"(本人|笔者|作者|本文|自己)的?(水平|能力|学识|知识)(有限|不足)|才疏学浅|(时间|时间较为|时间比较)仓促|难免(存在|有)?(疏漏|不足|错误|纰漏|不妥)|恳请.{0,6}(批评指正|指正|斧正)"),
     "“水平有限，难免疏漏”"),
]

ENDING_RX = re.compile(
    r"未来可期|前景(十分|非常)?(广阔|光明|无限|可观)|广阔的?(应用|发展)?前景|大有可为|值得期待|必将|定将"
    r"|相信(在|随着|未来|通过|经过|只要)|让我们(一起|携手|共同)|携手|迈向|开创.{0,8}新|新篇章|新征程|美好(的)?(未来|明天|生活)"
    r"|无限(可能|潜力)|潜力巨大|充满(信心|希望|期待)|任重道远|砥砺前行|行稳致远|不断(前进|进步|前行|攀登)|继续(努力|前行|奋斗)"
    r"|更加(美好|辉煌|光明)|贡献(自己的?|一份)?(力量|智慧)")

# 至少一侧是中文才算：不算“1990—2000 年”“3—5 页”这类范围，也不算方程式里当箭头用的“——”
DASH_RX = re.compile(r"(?<=[\u4e00-\u9fa5，。、：”）》])\s?——|——\s?(?=[\u4e00-\u9fa5“（《])"
                     r"|(?<=[\u4e00-\u9fa5，。、”）])\s?—\s?(?=[\u4e00-\u9fa5“（])|(?<=[\u4e00-\u9fa5])\s?--\s?(?=[\u4e00-\u9fa5])")


def style_units(doc):
    """参与 AI 味统计的段落：正文段落（不含标题、题注、表格、代码、页眉页脚、参考文献）。"""
    return [u for u in doc.units if u.kind == "para" and not u.in_refs and u.text.strip()]


def style_limit(chars, base, per10k):
    return base + int(chars * per10k / 10000)


def _where(units, cap=20):
    seen, out = set(), []
    for u in units:
        if u.loc not in seen:
            seen.add(u.loc)
            out.append(u)
    names = "、".join(re.sub(r"^第 (\d+) (段|行)$", r"\1", u.loc) for u in out[:cap])
    unit_word = "行" if out and out[0].loc.endswith("行") else "段"
    tail = f" 等 {len(out)} 处" if len(out) > cap else ""
    if all(re.match(r"^第 \d+ (段|行)$", u.loc) for u in out[:cap]):
        return f"第 {names} {unit_word}{tail}"
    return "、".join(u.loc for u in out[:cap]) + tail


def _style_add(doc, code, hits, problem):
    """hits: [(unit, match)]，合成一条 WARN，原句取第一处。"""
    u, m = hits[0]
    add(doc, WARN, code, "STYLE", u, excerpt(u.text, m.start(), m.end()),
        f"{problem}。位置：{_where([h[0] for h in hits])}", code)


def check_style(doc, ignore):
    units = style_units(doc)
    chars = sum(len(re.findall(r"[\u4e00-\u9fa5]", u.text)) for u in units)
    if chars < 200:
        return  # 太短，没有“泛滥”可言

    def on(code):
        return code not in ignore

    if on("STYLE_WORD"):
        hits, counts = [], defaultdict(int)
        for name, pattern in STYLE_WORDS:
            rx = re.compile(pattern)
            for u in units:
                for m in rx.finditer(u.text):
                    hits.append((u, m))
                    counts[name] += 1
        limit = style_limit(chars, 4, 3)
        if len(hits) >= limit:
            hits.sort(key=lambda h: (h[0].n, h[1].start()))
            words = "、".join(f"{k}×{v}" for k, v in sorted(counts.items(), key=lambda kv: -kv[1]))
            _style_add(doc, "STYLE_WORD", hits,
                       f"套话词全文 {len(hits)} 处（{words}），按本文篇幅（约 {chars} 字）超过 {limit} 处就算泛滥")

    if on("STYLE_SUMUP") and sum(bool(SOLUTION_MARK.search(u.text)) for u in units) < 3:
        hits = [(u, m) for u in units for m in SUMUP_RX.finditer(u.text)
                if not SUMUP_ANSWER.search(excerpt(u.text, m.start(), m.end()))]
        limit = style_limit(chars, 2, 1)
        if len(hits) >= limit:
            _style_add(doc, "STYLE_SUMUP", hits,
                       f"“综上所述/总而言之/总之”一类总结词全文 {len(hits)} 处，一篇一般只在结论开头用一次")

    if on("STYLE_TRIAD"):
        hits = []
        for u in units:
            for rx in TRIAD_RULES:
                m = rx.search(u.text)
                if m:
                    hits.append((u, m))
                    break
        limit = style_limit(chars, 3, 2)
        if len(hits) >= limit:
            _style_add(doc, "STYLE_TRIAD", hits,
                       f"三段式排比（“高效、准确、稳定”“提升了……提升了……提升了……”）全文 {len(hits)} 处")

    if on("STYLE_ELEVATE"):
        heads = [i for i, u in enumerate(doc.units) if u.kind == "heading" and u.text.strip()]
        sections, hits = 0, []
        for k, i in enumerate(heads):
            end = heads[k + 1] if k + 1 < len(heads) else len(doc.units)
            if ELEVATE_SKIP_HEAD.search(doc.units[i].text) or REF_HEAD.match(doc.units[i].text.strip("# ")):
                continue
            body = [u for u in doc.units[i + 1:end] if u.kind == "para" and not u.in_refs and u.text.strip()]
            if not body:
                continue
            sections += 1
            last = body[-1].text.strip()
            sents = [s for s in re.split(r"(?<=[。！？!?])", last) if s.strip()]
            tail = sents[-1] if sents else last
            m = ELEVATE_RX.search(tail)
            if m:
                hits.append((body[-1], _Shift(m, last.rfind(tail) + len(body[-1].text) - len(body[-1].text.lstrip()))))
        if len(hits) >= 2 and len(hits) >= sections * 0.25:
            _style_add(doc, "STYLE_ELEVATE", hits,
                       f"{sections} 个小节里有 {len(hits)} 个以升华句收尾（“具有重要意义”“为……奠定了基础”“彰显了……”）")

    if on("STYLE_CONCEDE"):
        hits, kinds = [], []
        for u in units:
            for rx, name in CONCEDE_RULES:
                m = rx.search(u.text)
                if m:
                    hits.append((u, m))
                    if name not in kinds:
                        kinds.append(name)
                    break
        if hits:
            _style_add(doc, "STYLE_CONCEDE", hits,
                       f"让步式自贬 {len(hits)} 处（{'、'.join(kinds)}）：先说别的更好再说自己选了差的，或者自谦套话，读起来像在找借口")

    if on("STYLE_DASH"):
        hits = [(u, m) for u in units for m in DASH_RX.finditer(u.text)]
        limit = style_limit(chars, 4, 2)
        if len(hits) >= limit:
            _style_add(doc, "STYLE_DASH", hits,
                       f"破折号全文 {len(hits)} 处，一篇里两三处为宜；连着用是 AI 制造节奏的典型写法")

    if on("STYLE_ENDING"):
        stop = len(doc.units)
        for i, u in enumerate(doc.units):
            t = u.text.strip("# ").strip()
            if u.in_refs or (REF_HEAD.match(t) or AFTER_REFS.match(t)) and len(t) <= 20:
                stop = i
                break
        tail_units = [u for u in doc.units[:stop] if u.kind == "para" and u.text.strip()][-3:]
        hits = [(u, m) for u in tail_units for m in [ENDING_RX.search(u.text)] if m]
        if hits:
            _style_add(doc, "STYLE_ENDING", hits,
                       "全文以空泛的乐观展望收尾（“未来可期”“前景广阔”“相信……必将……”），没有信息量")


class _Shift:
    """把在子串上的匹配位置平移回整段，供 excerpt 定位。"""

    def __init__(self, m, off):
        self._s, self._e = m.start() + off, m.end() + off

    def start(self):
        return self._s

    def end(self):
        return self._e


# ---------------------------------------------------------------- 汇总输出

def scan(path, fmt=None, ignore=(), style=True):
    doc = load(path, fmt)
    ignore = set(ignore)
    mark_refs(doc)
    check_phrases(doc, ignore)
    check_short_version_lines(doc, ignore)
    check_format_residue(doc, ignore)
    check_numbering(doc, ignore)
    check_headings(doc, ignore)
    check_citations(doc, ignore)
    check_misc(doc, ignore)
    check_docx_parts(doc, ignore)
    if style:
        check_style(doc, ignore)
    # 同一句话命中多条规则时合成一条：级别取最严重的，原因都列出来（AI 味是全文汇总，不和别的合并）
    merged = {}
    for f in doc.findings:
        k = (f.loc, f.excerpt, f.code if f.code.startswith("STYLE_") else "")
        if k not in merged:
            merged[k] = f
            continue
        g = merged[k]
        if f.code not in g.code.split("+"):
            g.code += "+" + f.code
        if f.problem not in g.problem:
            g.problem += "；" + f.problem
        if f.severity == ERROR and g.severity == WARN:
            g.severity, g.category, g.suggestion = ERROR, f.category, f.suggestion or g.suggestion
    doc.findings = sorted(merged.values(), key=lambda f: (f.severity != ERROR, f.code.startswith("STYLE_"), f.n == 0, f.n))
    return doc


def render(doc):
    e = sum(f.severity == ERROR for f in doc.findings)
    w = len(doc.findings) - e
    out = [f"leak-check 扫描：{doc.path}（{doc.fmt}，{len(doc.units)} 段/行）", f"  ERROR {e} · WARN {w}"]
    for note in doc.notes:
        out.append(f"  注：{note}")
    for f in doc.findings:
        out.append("")
        out.append(f"[{f.severity}] {f.category} · {f.loc} · {f.code}")
        if f.excerpt:
            out.append(f"  原句：{f.excerpt}")
        out.append(f"  问题：{f.problem}")
        if f.suggestion:
            out.append(f"  建议：{f.suggestion}")
    if not doc.findings:
        out.append("  没有扫到问题。脚本只覆盖字符串和结构能判断的部分，语义问题还要通读。")
    return "\n".join(out)


def main(argv=None):
    ap = argparse.ArgumentParser(description="找出论文/作业里露馅或一眼不对的句子")
    ap.add_argument("files", nargs="+", help="文件路径，'-' 表示标准输入")
    ap.add_argument("--format", help="强制按某种格式读（md/txt/tex/docx/pdf）")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--ignore", nargs="*", default=[], help="忽略的规则代码，如 CITE_ORDER DUP_PARA")
    ap.add_argument("--no-style", action="store_true", help="不查 AI 味（STYLE_* 规则）")
    ap.add_argument("--dump", action="store_true", help="不检查，按扫描用的段号/行号输出全文，供通读和定位")
    args = ap.parse_args(argv)
    if args.dump:
        for path in args.files:
            doc = load(path, args.format)
            mark_refs(doc)
            print(f"===== {doc.path}")
            for u in doc.units:
                if u.text.strip():
                    tag = u.kind + (f"{u.level}" if u.kind == "heading" else "") + ("·参考文献" if u.in_refs else "")
                    print(f"[{u.loc}|{tag}] {u.text}")
        return 0
    results, any_error = [], False
    for path in args.files:
        if path != "-" and not Path(path).exists():
            print(f"找不到文件：{path}", file=sys.stderr)
            return 2
        doc = scan(path, args.format, args.ignore, style=not args.no_style)
        any_error |= any(f.severity == ERROR for f in doc.findings)
        results.append(doc)
    if args.json:
        print(json.dumps([{
            "file": d.path, "format": d.fmt, "notes": d.notes,
            "errors": sum(f.severity == ERROR for f in d.findings),
            "warnings": sum(f.severity == WARN for f in d.findings),
            "findings": [f.__dict__ for f in d.findings],
        } for d in results], ensure_ascii=False, indent=2))
    else:
        print("\n\n".join(render(d) for d in results))
    return 1 if any_error else 0


if __name__ == "__main__":
    sys.exit(main())
