# -*- coding: utf-8 -*-
"""
word_com
========
对 win32com.client 常用对象/枚举的小封装,降低上层 skill 的拼写噪音。
把常量名集中在这里,避免散落 magic number。
"""

from __future__ import annotations

from typing import Any, Dict, Iterable, Optional

try:
    import win32com.client  # type: ignore
    from win32com.client import constants  # type: ignore
except ImportError:  # 静态分析兼容
    win32com = None  # type: ignore
    constants = None  # type: ignore


# ---------------------------------------------------------------------------
# Word 枚举常量(无法保证 constants 一定可用,提供 fallback)
# ---------------------------------------------------------------------------
# 段落对齐
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2
WD_ALIGN_PARAGRAPH_JUSTIFY = 3

# 分节符
WD_SECTION_CONTINUOUS = 0
WD_SECTION_NEW_COLUMN = 1
WD_SECTION_NEW_PAGE = 2
WD_SECTION_EVEN_PAGE = 3
WD_SECTION_ODD_PAGE = 4

# 分页符
WD_PAGE_BREAK = 7

# 行距规则
WD_LINE_SINGLE = 0
WD_LINE_1_5 = 1
WD_LINE_DOUBLE = 2
WD_LINE_AT_LEAST = 3
WD_LINE_EXACTLY = 4
WD_LINE_MULTIPLE = 5

# 域类型
WD_FIELD_PAGE = "PAGE"
WD_FIELD_NUMPAGES = "NUMPAGES"
WD_FIELD_TOC = "TOC"
WD_FIELD_STYLEREF = "STYLEREF"

# 单位
WD_TWIPS_PER_POINT = 20  # 1 pt = 20 twips


# ---------------------------------------------------------------------------
# 段落处理工具
# ---------------------------------------------------------------------------
def iter_paragraphs(doc) -> Iterable[Any]:
    """遍历所有段落,跨 Range/Selection 友好。"""
    return doc.Paragraphs


def para_text(p) -> str:
    return (p.Range.Text or "").rstrip("\r").rstrip("\n").strip()


def set_run_font(
    rng,
    *,
    east: Optional[str] = None,
    west: Optional[str] = None,
    size: Optional[float] = None,
    bold: Optional[bool] = None,
    color: Optional[int] = None,
) -> None:
    """设置一个 Range 的字体。"""
    f = rng.Font
    if east is not None:
        f.NameFarEast = east
    if west is not None:
        f.Name = west
    if size is not None:
        f.Size = float(size)
    if bold is not None:
        f.Bold = bool(bold)
    if color is not None:
        f.Color = int(color)


def insert_paragraph_at(doc, index: int, text: str = "", style: Optional[str] = None):
    """在指定位置前插入新段落,返回新段落对象。"""
    rng = doc.Range(0, 0) if index == 0 else doc.Paragraphs(index).Range.Duplicate
    rng.Collapse(1)  # wdCollapseStart
    rng.InsertBefore(text or "\r")
    new_para = doc.Paragraphs(index + 1)
    if style:
        try:
            new_para.Style = doc.Styles(style)
        except Exception:
            pass
    return new_para


def add_field(rng, field_code: str, default_text: str = "") -> None:
    """在 rng 处插入 Word 域(用于 PAGE / TOC / STYLEREF)。"""
    rng.Fields.Add(Range=rng, Type=-1, Text=field_code, PreserveFormatting=False)
    if default_text:
        rng.Text = default_text


def get_page_setup(doc):
    return doc.PageSetup


def get_sections(doc):
    return doc.Sections


def get_style(doc, *aliases: str):
    """按顺序尝试内置样式的多个别名,返回第一个存在的 Style。

    中文 Word 把内置样式的本地名翻译为 "标题 1"/"正文"/"表目录 1" 等,
    英文 Word 仍是 "Heading 1"/"Normal"/"TOC 1"。为了一次写代码两种
    Word 都能跑,所有需要 doc.Styles(name) 的地方都走这个函数。

    注意:win32com 在样式不存在时可能不抛异常,反而返回 Styles 集合本身,
    这里用 .Font 属性探测一次,确保返回的是真正的 Style 对象。
    """
    seen = set()
    for raw in aliases:
        if not raw:
            continue
        for name in (raw, _STYLE_ALIASES.get(raw, "")):
            if not name or name in seen:
                continue
            seen.add(name)
            try:
                s = doc.Styles(name)
                _ = s.Font  # 探测:集合对象无 .Font,会抛 AttributeError
                return s
            except Exception:
                continue
    return None


# 内置样式的"英文名 -> 中文名"补充别名,主查找传入英文也能解析中文 Word
_STYLE_ALIASES: Dict[str, str] = {
    "Heading 1": "标题 1", "Heading 2": "标题 2",
    "Heading 3": "标题 3", "Heading 4": "标题 4",
    "Normal": "正文",
    "TOC 1": "表目录 1", "TOC 2": "表目录 2", "TOC 3": "表目录 3",
    "TOC 4": "表目录 4", "TOC 5": "表目录 5", "TOC 6": "表目录 6",
    "TOC 7": "表目录 7", "TOC 8": "表目录 8", "TOC 9": "表目录 9",
}


def enable_odd_even_pages(doc) -> None:
    """开启页眉/页脚的奇偶页区分(影响所有 section)。"""
    doc.PageSetup.OddAndEvenPagesHeaderFooter = True


# ---------------------------------------------------------------------------
# 安全 Selection
# ---------------------------------------------------------------------------
class SafeSelection:
    """在 Word 中所有操作尽量走 Range,避免 Selection 闪屏。"""

    def __init__(self, doc) -> None:
        self._doc = doc

    def find_paragraph_starting_with(self, prefix: str) -> Optional[Any]:
        for p in self._doc.Paragraphs:
            if para_text(p).startswith(prefix):
                return p
        return None
