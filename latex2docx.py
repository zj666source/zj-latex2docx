#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
latex2docx.py  —— 把 Word 文档里以纯文本形式存在的 LaTeX 公式，
一键编译成 Word 原生可编辑公式对象（OMML），无需打开 Word。

流程:  每段文本 -> latex2mathml(LaTeX->MathML) -> 微软官方 MML2OMML.XSL(MathML->OMML)
依赖:  latex2mathml, lxml, MML2OMML.XSL(随 exe 打包)
"""
import os
import re
import sys
import io
import zipfile

import latex2mathml.converter as l2m
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MATHML_NS = "http://www.w3.org/1998/Math/MathML"

# 程序署名（窗口标题、文件框标题、完成提示均使用此署名）
BRAND = "zj公式转化工具"

# 让 lxml 输出时用 m: 前缀（Word 按命名空间 URI 识别，前缀名本身不影响解析）
etree.register_namespace("m", M)

# ---------- 资源路径（兼容 PyInstaller 单文件） ----------
def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

XSL_PATH = resource_path("MML2OMML.XSL")
# 用字节读取，避免 lxml 对含中文路径的 file: URL 解析失败
with open(XSL_PATH, "rb") as _f:
    _XSLT = etree.XSLT(etree.parse(io.BytesIO(_f.read())))

# ---------- LaTeX -> OMML 元素 ----------
def latex_to_omml(latex):
    """返回 <m:oMath> 的 lxml 元素（独立，可插入文档树）"""
    mm = l2m.convert(latex)
    root = etree.fromstring(mm.encode("utf-8"))
    res = _XSLT(root)
    omml_root = res.getroot()  # <m:oMath ...>
    # 重新序列化再解析，脱离 XSLT 结果树
    return etree.fromstring(etree.tostring(omml_root))

# ---------- 文本分段：识别公式区间 ----------
_CJK = re.compile(r"[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]")
# 中文标点（作为“非公式”边界）。注意：不要包含空格，空格是公式内部的合法字符，
# 否则会把 "A = B" 按空格切成多个独立公式。
_CN_PUNCT = set("，。．、；：！？（）《》「」『』“”‘’—…·")
# 触发“公式”判定的强信号：LaTeX 反斜杠、上下标、典型数学运算符
# （不含普通比较符 = < > 与 + -，避免把“≤26℃”“HEPA+”“+ A”误转成公式）
_SIGNAL = re.compile(r"[\\^_=+×÷±≤≥≈∼∂√∑∫∞→∝<>]")
_STRONG = re.compile(r"[\\^_×·÷±≤≥≈∼∂√∑∫∞→∝]")

def _classify(ch, depth):
    # 括号深度 > 0 时，即便遇到中文也视为公式内部（如 E_{夏} 的下标）
    if depth > 0:
        return "mathchar"
    if _CJK.match(ch) or ch in _CN_PUNCT:
        return "other" if not _CJK.match(ch) else "other"
    return "mathchar"

def _is_math(seg):
    s = seg.strip()
    if not s:
        return False
    # 含 LaTeX 反斜杠：必为公式
    if "\\" in s:
        return True
    # 含上下标：必为公式
    if "_" in s or "^" in s:
        return True
    # 含强数学运算符（× · ÷ ± ≤ ≥ ≈ √ ∑ 等）
    if _STRONG.search(s):
        return True
    return False

def _is_bare_operator(seg):
    """孤立的单个运算符（如单元格里单独的 '+'）不转成公式，保留为文本。"""
    s = seg.strip()
    if len(s) <= 1:
        return True
    return False

def split_text(text):
    """括号深度感知地分段：括号(含下标)内的中文视为公式一部分，括号外的纯中文作边界。
    返回 [(kind, seg), ...]，kind ∈ {'text','math'}，相邻同类合并。"""
    depth = 0
    tokens = []
    cur_kind = None
    cur = []
    def flush():
        nonlocal cur_kind, cur
        if cur_kind is not None and cur:
            tokens.append((cur_kind, "".join(cur)))
        cur_kind = None
        cur = []
    for ch in text:
        if ch == "{":
            depth += 1
            k = "mathchar"
        elif ch == "}":
            depth -= 1
            k = "mathchar"
        elif depth > 0:
            k = "mathchar"
        elif _CJK.match(ch) or ch in _CN_PUNCT:
            k = "other"
        else:
            k = "mathchar"
        if k != cur_kind:
            flush()
            cur_kind = k
        cur.append(ch)
    flush()
    # mathchar 段再判定是否为公式
    result = []
    for kind, seg in tokens:
        if kind == "mathchar":
            result.append(("math" if _is_math(seg) else "text", seg))
        else:
            result.append(("text", seg))
    # 合并相邻同类
    merged = []
    for kind, seg in result:
        if merged and merged[-1][0] == kind:
            merged[-1] = (kind, merged[-1][1] + seg)
        else:
            merged.append((kind, seg))
    return merged

# ---------- 重建一个 run ----------
def rebuild_run(wr, text, stats):
    """把含有 LaTeX 的 <w:r> 拆成若干 <w:r>(文本) / <m:oMath>(公式) 元素列表。"""
    rpr = wr.find(f"{{{W}}}rPr")
    tokens = split_text(text)
    out = []
    for kind, seg in tokens:
        if kind == "math":
            latex = seg.strip()
            if not latex:
                continue
            if _is_bare_operator(latex):
                # 孤立运算符（如表格单元格里单独的 '+'），保留为文本
                stats.setdefault("skip", [])
                r = etree.Element(f"{{{W}}}r")
                if rpr is not None:
                    r.append(rpr_copy(rpr))
                t = etree.SubElement(r, f"{{{W}}}t")
                t.text = seg
                out.append(("run", r))
                continue
            try:
                omath = latex_to_omml(latex)
            except Exception as e:
                # 解析失败：当普通文本保留
                stats["skip"].append((latex, str(e)[:60]))
                r = etree.Element(f"{{{W}}}r")
                if rpr is not None:
                    r.append(rpr_copy(rpr))
                t = etree.SubElement(r, f"{{{W}}}t")
                t.text = seg
                out.append(("run", r))
                continue
            stats["converted"].append(latex)
            out.append(("omath", omath))
        else:
            seg2 = seg
            if not seg2:
                continue
            r = etree.Element(f"{{{W}}}r")
            if rpr is not None:
                r.append(rpr_copy(rpr))
            t = etree.SubElement(r, f"{{{W}}}t")
            t.text = seg2
            out.append(("run", r))
    return out

def rpr_copy(rpr):
    return etree.fromstring(etree.tostring(rpr))

# ---------- 处理整个文档树 ----------
def process_tree(root_el, stats):
    body = root_el
    # 遍历所有段落（含表格内）
    for p in root_el.iter(f"{{{W}}}p"):
        children = list(p)
        for child in children:
            if child.tag != f"{{{W}}}r":
                continue
            t_el = child.find(f"{{{W}}}t")
            if t_el is None or t_el.text is None:
                continue
            text = t_el.text
            # 快速预筛：必须含潜在公式字符
            if not (_SIGNAL.search(text) or "_" in text or "^" in text):
                continue
            # 进一步判断本段是否真的含公式区间
            toks = split_text(text)
            if not any(k == "math" for k, _ in toks):
                continue
            new_elems = rebuild_run(child, text, stats)
            parent = p
            idx = list(parent).index(child)
            offset = 0
            for kind, el in new_elems:
                parent.insert(idx + offset, el)
                offset += 1
            parent.remove(child)

def add_math_ns(root_el):
    # oMath 元素自身已带 xmlns:m 声明，Word 按 URI 识别，无需在根重复声明
    pass

# ---------- 主处理：docx 文件 ----------
def convert_docx(src, dst, stats):
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    # 需要扫描的部件（主文档 + 页眉页脚，通常公式都在正文）
    targets = [n for n in names if n.endswith(".xml") and
               (n == "word/document.xml" or
                re.match(r"word/(header|footer)\d*\.xml", n))]
    for part in targets:
        xml = data[part].decode("utf-8")
        parser = etree.XMLParser(remove_blank_text=False)
        root = etree.fromstring(xml.encode("utf-8"), parser)
        add_math_ns(root)
        process_tree(root, stats)
        data[part] = etree.tostring(root, xml_declaration=True,
                                    encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])

# ---------- 文件选择 / 弹窗（ctypes，标准库，无需额外依赖） ----------
def msgbox(text, title=BRAND, style=0x40):
    """style: 0x40=信息, 0x30=警告, 0x10=错误。返回 None。"""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, str(text), str(title), style)
    except Exception:
        try:
            print(text)
        except Exception:
            pass

def pick_docx():
    """弹出 Windows 文件选择框，返回路径；取消/失败返回 None。"""
    try:
        import ctypes
        from ctypes import wintypes
        class OFN(ctypes.Structure):
            _fields_ = [
                ("lStructSize", wintypes.DWORD),
                ("hwndOwner", wintypes.HWND),
                ("hInstance", wintypes.HINSTANCE),
                ("lpstrFilter", wintypes.LPCWSTR),
                ("lpstrCustomFilter", wintypes.LPCWSTR),
                ("nMaxCustFilter", wintypes.DWORD),
                ("nFilterIndex", wintypes.DWORD),
                ("lpstrFile", wintypes.LPWSTR),
                ("nMaxFile", wintypes.DWORD),
                ("lpstrFileTitle", wintypes.LPWSTR),
                ("nMaxFileTitle", wintypes.DWORD),
                ("lpstrInitialDir", wintypes.LPCWSTR),
                ("lpstrTitle", wintypes.LPCWSTR),
                ("Flags", wintypes.DWORD),
                ("nFileOffset", wintypes.WORD),
                ("nFileExtension", wintypes.WORD),
                ("lpstrDefExt", wintypes.LPCWSTR),
                ("lCustData", wintypes.LPARAM),
                ("lpfnHook", wintypes.LPARAM),
                ("lpTemplateName", wintypes.LPCWSTR),
                ("pvReserved", wintypes.LPVOID),
                ("dwReserved", wintypes.DWORD),
                ("flagsEx", wintypes.DWORD),
            ]
        buf = ctypes.create_unicode_buffer(2048)
        ofn = OFN()
        ofn.lStructSize = ctypes.sizeof(ofn)
        ofn.lpstrFilter = "Word 文档 (*.docx)\0*.docx\0所有文件 (*.*)\0*.*\0\0"
        ofn.lpstrTitle = f"{BRAND} — 选择要转换的 Word 文档"
        ofn.Flags = 0x00080000 | 0x00001000   # OFN_EXPLORER | OFN_FILEMUSTEXIST
        ofn.lpstrFile = buf
        ofn.nMaxFile = 2048
        if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
            return buf.value
        return None
    except Exception:
        return None

def _write_log(exc):
    """把异常写到 exe 同目录 error.log，方便用户反馈。"""
    try:
        import traceback
        here = os.path.dirname(os.path.abspath(sys.argv[0])) if getattr(
            sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
        log = os.path.join(here, "公式转换_error.log")
        with open(log, "a", encoding="utf-8") as f:
            f.write("==== " + __import__("datetime").datetime.now().isoformat() + " ====\n")
            f.write(traceback.format_exc() + "\n\n")
    except Exception:
        pass

def do_convert(path):
    if not os.path.isfile(path):
        msgbox("找不到文件：\n" + path, style=0x10)
        return
    if not path.lower().endswith(".docx"):
        msgbox("仅支持 .docx 文件（不支持旧版 .doc）。\n\n可把 .docx 直接拖到本程序图标上运行。",
               style=0x10)
        return

    base, _ = os.path.splitext(path)
    dst = base + "_公式版.docx"
    if os.path.exists(dst) and os.path.abspath(dst) != os.path.abspath(path):
        dst = base + "_公式版_2.docx"
    stats = {"converted": [], "skip": []}
    convert_docx(path, dst, stats)

    msg = f"转换完成！\n\n成功转换公式 {len(stats['converted'])} 处。\n\n输出文件：\n{dst}"
    if stats["skip"]:
        msg += f"\n\n有 {len(stats['skip'])} 处未能解析，已按原文保留：\n"
        msg += "\n".join("  - " + s[:50] for s, _ in stats["skip"][:8])
    msg += f"\n\n—— {BRAND}"
    msgbox(msg)

def main():
    try:
        args = [a for a in sys.argv[1:] if not a.startswith("-")]
        path = args[0] if args else None
        if not path:
            path = pick_docx()
        if not path:
            # 用户取消对话框，或对话框不可用：给出明确指引，而不是“没反应”
            msgbox(f"未选择文件。\n\n—— {BRAND}\n\n用法二选一：\n"
                   "1) 把 Word 文档(.docx)直接拖到本程序图标上；\n"
                   "2) 在命令行运行：  公式一键转换.exe 文档路径.docx",
                   style=0x40)
            return
        do_convert(path)
        # console 模式（sys.stdout 存在）下暂停，便于查看结果；windowed 模式不暂停
        if sys.stdout is not None:
            input("\n按回车退出...")
    except Exception as e:
        import traceback
        _write_log(e)
        msgbox("程序运行出错（详细错误已写入同目录 error.log）：\n\n" + str(e)[:200],
               style=0x10)

if __name__ == "__main__":
    main()
