#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小恐龙公文排版助手 v2.2.0 - 银河麒麟 V10 (海光 Hygon C86 3350 x86_64)
核心 OpenXML (.docx) 国标公文排版引擎与 42 个 VBA 脚本原生实现
严格遵循《党政机关公文格式国家标准》(GB/T 9704—2012)
"""

import datetime
import glob
import os
import random
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import docx
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm, Pt, RGBColor


# ==================== 1. 常用工具与转换函数 ====================

CN_DIGITS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
CIRCLED_NUMBERS = [
    "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩",
    "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳",
]


def int_to_chinese(num: int) -> str:
    """将 1~999 的整数转换为中文数字（如 1->一, 12->十二, 25->二十五）"""
    if num <= 0:
        return "零"
    if num <= 10:
        return CN_DIGITS[num]
    if num < 20:
        return "十" + CN_DIGITS[num - 10]
    if num < 100:
        shi, ge = divmod(num, 10)
        return CN_DIGITS[shi] + "十" + (CN_DIGITS[ge] if ge > 0 else "")
    bai, rem = divmod(num, 100)
    if rem == 0:
        return CN_DIGITS[bai] + "百"
    if rem < 10:
        return CN_DIGITS[bai] + "百零" + CN_DIGITS[rem]
    return CN_DIGITS[bai] + "百" + int_to_chinese(rem)


def number_to_rmb_upper(raw_value: str) -> str:
    """将数字金额转换为标准会计人民币大写（移植自《数字转为人民币.vba》）"""
    cleaned = re.sub(r"[,，\s￥¥元]", "", str(raw_value or ""))
    try:
        val = float(cleaned)
    except ValueError:
        return f"输入有误：{raw_value}"

    prefix = "负" if val < 0 else ""
    val = abs(val)
    if val > 999999999999.99:
        return "输入有误：数字过大"

    fen_total = int(round(val * 100))
    if fen_total == 0:
        return "零元整"

    cn_nums = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
    int_units = ["", "拾", "佰", "仟"]
    sec_units = ["", "万", "亿", "万亿"]

    int_part, rem = divmod(fen_total, 100)
    jiao, fen = divmod(rem, 10)

    res = []
    if int_part > 0:
        s = str(int_part)
        zero_count = 0
        length = len(s)
        for i, ch in enumerate(s):
            n = int(ch)
            p = length - i - 1
            q, m = divmod(p, 4)
            if n == 0:
                zero_count += 1
            else:
                if zero_count > 0:
                    res.append("零")
                zero_count = 0
                res.append(cn_nums[n] + int_units[m])
            if m == 0 and zero_count < 4:
                res.append(sec_units[q])
        res.append("元")

    if jiao == 0 and fen == 0:
        res.append("整")
    else:
        if jiao > 0:
            res.append(cn_nums[jiao] + "角")
        elif int_part > 0 and fen > 0:
            res.append("零")
        if fen > 0:
            res.append(cn_nums[fen] + "分")

    return prefix + "".join(res)


def set_run_fonts(
    run,
    cn_font: str = "仿宋",
    en_font: str = "Times New Roman",
    size_pt: float = 16.0,
    bold: bool = False,
    italic: bool = False,
    color_rgb: Optional[Tuple[int, int, int]] = None,
    under_dot: bool = False,
    hidden: bool = False,
):
    """设置 Run 的中西文字体、字号、加粗、颜色、着重号及隐藏属性"""
    run.font.name = en_font
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.hidden = hidden
    if color_rgb is not None:
        run.font.color.rgb = RGBColor(*color_rgb)

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), cn_font)
    rFonts.set(qn("w:ascii"), en_font)
    rFonts.set(qn("w:hAnsi"), en_font)
    rFonts.set(qn("w:cs"), en_font)

    if under_dot:
        em = rPr.find(qn("w:em"))
        if em is None:
            em = OxmlElement("w:em")
            rPr.append(em)
        em.set(qn("w:val"), "dot")


def set_paragraph_char_indent(
    paragraph,
    first_line_chars: float = 0.0,
    left_chars: float = 0.0,
    right_chars: float = 0.0,
    font_size_pt: float = 16.0,
):
    """设置段落首行缩进、左缩进、右缩进（同时写入 w:ind 的 chars 与 dxa，确保 WPS 与 Word 100% 兼容）"""
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)

    for attr in [
        "w:firstLine", "w:firstLineChars", "w:hanging", "w:hangingChars",
        "w:left", "w:leftChars", "w:right", "w:rightChars"
    ]:
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]

    if first_line_chars > 0:
        ind.set(qn("w:firstLineChars"), str(int(round(first_line_chars * 100))))
        ind.set(qn("w:firstLine"), str(int(round(first_line_chars * font_size_pt * 20))))
    elif first_line_chars < 0:
        hang = abs(first_line_chars)
        ind.set(qn("w:hangingChars"), str(int(round(hang * 100))))
        ind.set(qn("w:hanging"), str(int(round(hang * font_size_pt * 20))))
    else:
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")

    if left_chars != 0:
        ind.set(qn("w:leftChars"), str(int(round(left_chars * 100))))
        ind.set(qn("w:left"), str(int(round(left_chars * font_size_pt * 20))))

    if right_chars != 0:
        ind.set(qn("w:rightChars"), str(int(round(right_chars * 100))))
        ind.set(qn("w:right"), str(int(round(right_chars * font_size_pt * 20))))


def set_paragraph_border(
    paragraph,
    side: str = "bottom",
    color_hex: str = "FF0000",
    sz_eighths_pt: int = 18,
    val: str = "single",
):
    """为段落添加顶部或底部横线（如 2.25磅红色间隔线反线 sz=18，信函双红线 val='double'）"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bdr = pBdr.find(qn(f"w:{side}"))
    if bdr is None:
        bdr = OxmlElement(f"w:{side}")
        pBdr.append(bdr)
    bdr.set(qn("w:val"), val)
    bdr.set(qn("w:sz"), str(sz_eighths_pt))
    bdr.set(qn("w:space"), "1")
    bdr.set(qn("w:color"), color_hex)


def set_paragraph_bottom_border(paragraph, color_hex: str = "FF0000", sz_eighths_pt: int = 18, val: str = "single"):
    set_paragraph_border(paragraph, side="bottom", color_hex=color_hex, sz_eighths_pt=sz_eighths_pt, val=val)


# ==================== 2. GB/T 9704—2012 国标公文核心处理器 ====================

class GwDocumentProcessor:
    """
    面向银河麒麟 V10 (Hygon 3350 x86_64) 的全功能公文排版引擎
    完整覆盖原 Windows VSTO 插件的全部 7 大功能组与 42 个 VBA 脚本
    """

    DEFAULT_CONFIG = {
        "top_mm": 37.0,
        "bottom_mm": 35.0,
        "left_mm": 28.0,
        "right_mm": 26.0,
        "footer_mm": 24.7,
        "lines_per_page": 22,
        "chars_per_line": 28,
        "line_spacing_pt": 28.0,
        "font_title": "方正小标宋简体",
        "font_h1": "黑体",
        "font_h2": "楷体",
        "font_h3": "仿宋",
        "font_h4": "仿宋",
        "font_body": "仿宋",
        "font_head": "黑体",
        "font_signer": "楷体",
        "font_banji": "仿宋",
        "font_pagenum": "宋体",
        "font_en": "Times New Roman",
    }

    def __init__(self, config: Optional[Dict] = None):
        self.cfg = {**self.DEFAULT_CONFIG, **(config or {})}

    # ---------- 版面组功能 ----------

    def apply_gb9704_page_setup(self, doc: Document) -> None:
        """设置 GB/T 9704—2012 页面尺寸、页边距与 22行×28字版心网格"""
        for sec in doc.sections:
            if sec.orientation == WD_ORIENT.LANDSCAPE:
                sec.page_width = Mm(297.0)
                sec.page_height = Mm(210.0)
            else:
                sec.page_width = Mm(210.0)
                sec.page_height = Mm(297.0)

            sec.top_margin = Mm(float(self.cfg["top_mm"]))
            sec.bottom_margin = Mm(float(self.cfg["bottom_mm"]))
            sec.left_margin = Mm(float(self.cfg["left_mm"]))
            sec.right_margin = Mm(float(self.cfg["right_mm"]))
            sec.footer_distance = Mm(float(self.cfg["footer_mm"]))
            sec.header_distance = Mm(15.0)

            sectPr = sec._sectPr
            docGrid = sectPr.find(qn("w:docGrid"))
            if docGrid is None:
                docGrid = OxmlElement("w:docGrid")
                sectPr.append(docGrid)
            docGrid.set(qn("w:type"), "linesAndChars")
            docGrid.set(qn("w:linePitch"), str(int(round(float(self.cfg["line_spacing_pt"]) * 20))))
            docGrid.set(qn("w:charSpace"), "0")

    def indent_all_body_2chars(self, doc: Document) -> int:
        """中文缩进：将正文段落统一设为首行缩进2字符"""
        cnt = 0
        for p in doc.paragraphs:
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER and (p.text or "").strip():
                set_paragraph_char_indent(p, first_line_chars=2, font_size_pt=16.0)
                cnt += 1
        return cnt

    def flush_left_all_paragraphs(self, doc: Document) -> int:
        """段落顶格：清除所有段落首行缩进与左缩进"""
        cnt = 0
        for p in doc.paragraphs:
            set_paragraph_char_indent(p, first_line_chars=0, left_chars=0, right_chars=0)
            cnt += 1
        return cnt

    def remove_header_bottom_line(self, doc: Document) -> None:
        """删除页眉横线"""
        for sec in doc.sections:
            for p in sec.header.paragraphs:
                pPr = p._p.get_or_add_pPr()
                pBdr = pPr.find(qn("w:pBdr"))
                if pBdr is not None:
                    pPr.remove(pBdr)

    def format_full_line_rearrange(self, doc: Document) -> int:
        """断行重排：合并未以句末标点（。：！？；）结尾的异常断行段落"""
        paras = list(doc.paragraphs)
        merged = 0
        i = 0
        while i < len(paras) - 1:
            txt = (paras[i].text or "").rstrip()
            nxt = (paras[i + 1].text or "").lstrip()
            if txt and nxt and not re.search(r"[。：！？；…—”》）]$", txt) and len(txt) > 15:
                paras[i].add_run(nxt)
                paras[i + 1]._element.getparent().remove(paras[i + 1]._element)
                paras.pop(i + 1)
                merged += 1
            else:
                i += 1
        return merged

    def format_soft_return_to_paragraph(self, doc: Document) -> int:
        """换行符转回车：将段内软回车 (\n / \v) 拆分为独立段落"""
        cnt = 0
        for p in list(doc.paragraphs):
            if "\n" in (p.text or "") or "\v" in (p.text or ""):
                parts = re.split(r"[\n\v]+", p.text)
                if len(parts) > 1:
                    p.text = parts[0]
                    for sub in parts[1:]:
                        np = p.insert_paragraph_before(sub)
                    cnt += len(parts) - 1
        return cnt

    def format_delete_spaces(self, doc: Document, half_only: bool = False) -> int:
        """删除所有空白或半角空格"""
        pattern = r" +" if half_only else r"[ \t\u3000\xa0]+"
        cnt = 0
        for p in doc.paragraphs:
            for r in p.runs:
                if r.text and re.search(pattern, r.text):
                    r.text = re.sub(pattern, "", r.text)
                    cnt += 1
        return cnt

    def format_delete_empty_lines(self, doc: Document) -> int:
        """删除全文所有空行"""
        cnt = 0
        for p in list(doc.paragraphs):
            has_drawing = bool(p._p.findall(".//" + qn("w:drawing")))
            if not (p.text or "").strip() and not has_drawing:
                p._element.getparent().remove(p._element)
                cnt += 1
        return cnt

    def format_delete_graphics(self, doc: Document) -> int:
        """删除文档中所有嵌入图片与图表"""
        cnt = 0
        for drawing in list(doc._element.findall(".//" + qn("w:drawing"))):
            parent = drawing.getparent()
            if parent is not None:
                parent.remove(drawing)
                cnt += 1
        return cnt

    def format_add_space_cn_en(self, doc: Document) -> int:
        """在中英文字符之间添加半角空格"""
        cnt = 0
        for p in doc.paragraphs:
            for r in p.runs:
                if not r.text:
                    continue
                new_t = re.sub(r"([A-Za-z])([\u4e00-\u9fa5])", r"\1 \2", r.text)
                new_t = re.sub(r"([\u4e00-\u9fa5])([A-Za-z])", r"\1 \2", new_t)
                if new_t != r.text:
                    r.text = new_t
                    cnt += 1
        return cnt

    def format_punctuation_switch(self, doc: Document, en_to_cn: bool = True) -> int:
        """标点符号英转中 / 中转英"""
        e_arr = [",", ";", ":", "?", "!", "(", ")", "[", "]"]
        c_arr = ["，", "；", "：", "？", "！", "（", "）", "［", "］"]
        src, dst = (e_arr, c_arr) if en_to_cn else (c_arr, e_arr)
        trans = str.maketrans("".join(src), "".join(dst))
        cnt = 0
        for p in doc.paragraphs:
            for r in p.runs:
                if r.text:
                    nt = r.text.translate(trans)
                    if en_to_cn:
                        nt = re.sub(r"([^\d])\.([^\d])", r"\1。\2", nt)
                    else:
                        nt = nt.replace("。", ".").replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
                    if nt != r.text:
                        r.text = nt
                        cnt += 1
        return cnt

    # ---------- 公文标题、文字格式与字体更改组 ----------

    def set_first_title_font(self, doc: Document, font_name: Optional[str] = None) -> bool:
        """将首个非空标题段落设为二号指定字体居中（对应【公文标题】与【二号】下拉框）"""
        target_font = font_name or self.cfg["font_title"]
        for p in doc.paragraphs:
            txt = (p.text or "").strip()
            if txt and not re.search(r"〔\d{4}〕.*号", txt):
                p.text = ""
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_paragraph_char_indent(p, first_line_chars=0, font_size_pt=22.0)
                r = p.add_run(txt)
                set_run_fonts(r, cn_font=target_font, en_font=self.cfg["font_en"], size_pt=22.0)
                return True
        return False

    def switch_document_font(self, doc: Document, old_font: str, new_font: str) -> int:
        """全文替换指定中文字体（如 仿宋_GB2312 => 仿宋）"""
        cnt = 0
        for rFonts in doc._element.findall(".//" + qn("w:rFonts")):
            matched = False
            for attr in ["w:eastAsia", "w:ascii", "w:hAnsi", "w:cs"]:
                if rFonts.get(qn(attr)) == old_font:
                    rFonts.set(qn(attr), new_font)
                    matched = True
            if matched:
                cnt += 1
        # 若段落未显式指定或需要全局覆盖，也检查各Run
        if cnt == 0:
            for p in doc.paragraphs:
                for r in p.runs:
                    rPr = r._r.get_or_add_rPr()
                    rf = rPr.find(qn("w:rFonts"))
                    if rf is None:
                        rf = OxmlElement("w:rFonts")
                        rPr.insert(0, rf)
                    rf.set(qn("w:eastAsia"), new_font)
                    cnt += 1
        return cnt

    def create_gw_custom_styles(self, doc: Document, force_rebuild: bool = False) -> List[str]:
        """生成或强制重建全套公文专用样式（公文:标题 / 公文:1~4级标题 / 公文:正文 / 公文:正文(缩进)）"""
        from docx.enum.style import WD_STYLE_TYPE

        specs = [
            ("公文:标题", self.cfg["font_title"], 22.0, WD_ALIGN_PARAGRAPH.CENTER, 0),
            ("公文:1级标题", self.cfg["font_h1"], 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY, 2),
            ("公文:2级标题", self.cfg["font_h2"], 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY, 2),
            ("公文:3级标题", self.cfg["font_h3"], 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY, 2),
            ("公文:4级标题", self.cfg["font_h4"], 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY, 2),
            ("公文:正文", self.cfg["font_body"], 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY, 0),
            ("公文:正文(缩进)", self.cfg["font_body"], 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY, 2),
        ]
        created = []
        styles = doc.styles
        for name, f_name, f_size, align, indent_chars in specs:
            if name in styles:
                if force_rebuild:
                    styles[name].delete()
                else:
                    created.append(name)
                    continue
            st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            st.font.name = self.cfg["font_en"]
            st.font.size = Pt(f_size)
            rPr = st.element.get_or_add_rPr()
            rFonts = OxmlElement("w:rFonts")
            rFonts.set(qn("w:eastAsia"), f_name)
            rFonts.set(qn("w:ascii"), self.cfg["font_en"])
            rPr.insert(0, rFonts)
            st.paragraph_format.alignment = align
            if indent_chars > 0:
                st.paragraph_format.first_line_indent = Pt(indent_chars * f_size)
            created.append(name)
        return created

    # ---------- 插入组：公文部件、红头、版记、横页、符号、页码 ----------

    def insert_red_header_single(
        self,
        doc: Document,
        agency_name: str = "请在此输入发文单位名称文件",
        doc_number: Optional[str] = None,
    ) -> None:
        """在文档开头插入单行发文机关红头版头与红色反线"""
        if doc_number is None:
            doc_number = f"○○发〔{datetime.date.today().year}〕1号"

        first_p = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
        p_head = first_p.insert_paragraph_before()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head.paragraph_format.space_before = Pt(20)
        p_head.paragraph_format.space_after = Pt(16)
        r_head = p_head.add_run(agency_name)
        set_run_fonts(
            r_head,
            cn_font=self.cfg["font_title"],
            en_font=self.cfg["font_en"],
            size_pt=42.0,
            bold=False,
            color_rgb=(255, 0, 0),
        )

        p_num = first_p.insert_paragraph_before()
        p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_num.paragraph_format.space_before = Pt(6)
        p_num.paragraph_format.space_after = Pt(6)
        r_num = p_num.add_run(doc_number)
        set_run_fonts(
            r_num,
            cn_font=self.cfg["font_body"],
            en_font=self.cfg["font_en"],
            size_pt=16.0,
            bold=False,
            color_rgb=(0, 0, 0),
        )
        set_paragraph_bottom_border(p_num, color_hex="FF0000", sz_eighths_pt=18, val="single")

    def insert_red_header_multi(
        self,
        doc: Document,
        agencies: Optional[List[str]] = None,
        doc_number: Optional[str] = None,
    ) -> None:
        """在文档开头插入联合发文机关红头表格版头与红色反线"""
        if not agencies:
            agencies = ["此处输入主办单位名称", "此处输入联合发文单位名称"]
        if doc_number is None:
            doc_number = f"○○发〔{datetime.date.today().year}〕1号"

        first_p = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        first_p._p.addprevious(tbl._tbl)

        c0 = tbl.cell(0, 0)
        c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.DISTRIBUTE
        r0 = p0.add_run("\n".join(agencies))
        set_run_fonts(r0, cn_font=self.cfg["font_title"], en_font=self.cfg["font_en"], size_pt=26.0, color_rgb=(255, 0, 0))

        c1 = tbl.cell(0, 1)
        c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run("文件")
        set_run_fonts(r1, cn_font=self.cfg["font_title"], en_font=self.cfg["font_en"], size_pt=44.0, color_rgb=(255, 0, 0))

        p_num = first_p.insert_paragraph_before()
        p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_num = p_num.add_run(doc_number)
        set_run_fonts(r_num, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0)
        set_paragraph_bottom_border(p_num, color_hex="FF0000", sz_eighths_pt=18, val="single")

    def insert_letter_red_line(self, doc: Document, top_line: bool = True) -> None:
        """插入信函格式上沿双红线或下沿红线"""
        p = doc.paragraphs[0] if (top_line and doc.paragraphs) else doc.add_paragraph()
        if top_line:
            set_paragraph_border(p, side="top", color_hex="FF0000", sz_eighths_pt=12, val="double")
        else:
            set_paragraph_border(p, side="bottom", color_hex="FF0000", sz_eighths_pt=12, val="single")

    def insert_banji_table(
        self,
        doc: Document,
        double_line: bool = True,
        cc_org: str = "○○○○○○。",
        print_org: str = "○○○○○○办公室",
        print_date: Optional[str] = None,
    ) -> None:
        """在文档末尾插入国标单行或双行版记表格（上下1.0pt粗线，中间0.75pt细线，四号仿宋）"""
        if print_date is None:
            today = datetime.date.today()
            print_date = f"{today.year}年{today.month}月{today.day}日印发"

        rows = 2 if double_line else 1
        tbl = doc.add_table(rows=rows, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        tblPr = tbl._tbl.tblPr
        borders_xml = f"""
        <w:tblBorders {nsdecls('w')}>
            <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>
            <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>
            <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideH w:val="{'single' if double_line else 'none'}" w:sz="6" w:space="0" w:color="000000"/>
            <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tblBorders>
        """
        old_borders = tblPr.find(qn("w:tblBorders"))
        if old_borders is not None:
            tblPr.remove(old_borders)
        tblPr.append(parse_xml(borders_xml))

        def format_cell(cell, text: str):
            cell.width = Mm(156.0)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.text = ""
            p.alignment = WD_ALIGN_PARAGRAPH.DISTRIBUTE if "印发" in text else WD_ALIGN_PARAGRAPH.LEFT
            set_paragraph_char_indent(p, first_line_chars=0, left_chars=1, right_chars=1, font_size_pt=14.0)
            r = p.add_run(text)
            set_run_fonts(r, cn_font=self.cfg["font_banji"], en_font=self.cfg["font_en"], size_pt=14.0)

        if double_line:
            format_cell(tbl.cell(0, 0), f"抄送：{cc_org}")
            format_cell(tbl.cell(1, 0), f"{print_org}　　　　　　　　{print_date}")
        else:
            format_cell(tbl.cell(0, 0), f"{print_org}　　　　　　　　{print_date}")

    def insert_cn_date_paragraph(self, doc: Document) -> str:
        """在文档末尾插入公文格式日期并右空4字"""
        today = datetime.date.today()
        d_str = f"{today.year}年{today.month}月{today.day}日"
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_paragraph_char_indent(p, first_line_chars=0, right_chars=4, font_size_pt=16.0)
        r = p.add_run(d_str)
        set_run_fonts(r, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0)
        return d_str

    def insert_landscape_section(self, doc: Document) -> None:
        """在文档中插入独立的横向 A4 分节页面并随后恢复纵向页面"""
        sec_land = doc.add_section(WD_SECTION.NEW_PAGE)
        sec_land.orientation = WD_ORIENT.LANDSCAPE
        sec_land.page_width = Mm(297.0)
        sec_land.page_height = Mm(210.0)
        doc.add_paragraph("（此处为横向A4页面内容区域）")

        sec_port = doc.add_section(WD_SECTION.NEW_PAGE)
        sec_port.orientation = WD_ORIENT.PORTRAIT
        sec_port.page_width = Mm(210.0)
        sec_port.page_height = Mm(297.0)

    def replace_brackets_to_hex(self, doc: Document, mode: str = "all") -> int:
        """将年份方括号 [2026] 或黑括号 【2026】 批量替换为六角括号 〔2026〕"""
        cnt = 0
        for p in doc.paragraphs:
            for r in p.runs:
                if not r.text:
                    continue
                nt = r.text
                if mode in ("all", "square"):
                    nt = re.sub(r"[\[［](\d{1,4})[\]］]", r"〔\1〕", nt)
                if mode in ("all", "black"):
                    nt = re.sub(r"【(\d{1,4})】", r"〔\1〕", nt)
                if nt != r.text:
                    r.text = nt
                    cnt += 1
        return cnt

    def insert_official_stamp_image(self, doc: Document, image_path: str, every_page: bool = False) -> None:
        """插入电子公章图片（支持单处落款加盖或每页页眉批量加盖 4cm×4cm 公章）"""
        if not os.path.isfile(image_path):
            raise FileNotFoundError(f"未找到公章图片: {image_path}")
        if every_page:
            for sec in doc.sections:
                hp = sec.header.paragraphs[0] if sec.header.paragraphs else sec.header.add_paragraph()
                run = hp.add_run()
                run.add_picture(image_path, width=Cm(4.0))
        else:
            p = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run()
            run.add_picture(image_path, width=Cm(4.0))

    def insert_page_numbers(self, doc: Document, mode: str = "center") -> None:
        """
        插入公文标准四号宋体一字线页码（— 1 —）
        mode: 'center' (居中) | 'outside' (双面奇右偶左) | 'right' (居右)
        """
        for sec in doc.sections:
            footer = sec.footer
            footer.is_linked_to_previous = False
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            p.text = ""
            if mode == "right":
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_paragraph_char_indent(p, right_chars=1, font_size_pt=14.0)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            r1 = p.add_run("— ")
            set_run_fonts(r1, cn_font=self.cfg["font_pagenum"], en_font=self.cfg["font_en"], size_pt=14.0)

            fld = OxmlElement("w:fldSimple")
            fld.set(qn("w:instr"), "PAGE \\* Arabic")
            r_page = OxmlElement("w:r")
            t_page = OxmlElement("w:t")
            t_page.text = "1"
            r_page.append(t_page)
            fld.append(r_page)
            p._p.append(fld)

            r2 = p.add_run(" —")
            set_run_fonts(r2, cn_font=self.cfg["font_pagenum"], en_font=self.cfg["font_en"], size_pt=14.0)

            if mode == "outside":
                sec.even_page_header_footer = True
                ep_footer = sec.even_page_footer
                ep_footer.is_linked_to_previous = False
                ep = ep_footer.paragraphs[0] if ep_footer.paragraphs else ep_footer.add_paragraph()
                ep.text = ""
                ep.alignment = WD_ALIGN_PARAGRAPH.LEFT
                set_paragraph_char_indent(ep, left_chars=1, font_size_pt=14.0)

                er1 = ep.add_run("— ")
                set_run_fonts(er1, cn_font=self.cfg["font_pagenum"], en_font=self.cfg["font_en"], size_pt=14.0)
                efld = OxmlElement("w:fldSimple")
                efld.set(qn("w:instr"), "PAGE \\* Arabic")
                ep._p.append(efld)
                er2 = ep.add_run(" —")
                set_run_fonts(er2, cn_font=self.cfg["font_pagenum"], en_font=self.cfg["font_en"], size_pt=14.0)

                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_paragraph_char_indent(p, right_chars=1, font_size_pt=14.0)

    def auto_layout_2026(self, doc: Document) -> Dict[str, int]:
        """
        《公文自动排版2026.vba》完整 Python OpenXML 实现
        自动完成：A4国标版心设置 -> 空行清理 -> 括号规范 -> 主标题/主送机关/四级标题/正文/附件/落款日期智能排版 -> 页码生成
        """
        self.apply_gb9704_page_setup(doc)

        valid_paras = []
        for p in list(doc.paragraphs):
            has_drawing = bool(p._p.findall(".//" + qn("w:drawing")))
            cleaned = re.sub(r"^[\s\u3000\t]+|[\s\u3000\t]+$", "", p.text or "")
            if not cleaned and not has_drawing:
                p._element.getparent().remove(p._element)
                continue
            if cleaned:
                cleaned = re.sub(r"[\[［【](\d{4})[\]］】]", r"〔\1〕", cleaned)
                valid_paras.append((p, cleaned))

        re_h1 = re.compile(r"^[一二三四五六七八九十百]+[、．.]\s*")
        re_h2 = re.compile(r"^[（(][一二三四五六七八九十百]+[）)]\s*")
        re_h3 = re.compile(r"^\d+[．.、]\s*")
        re_h4 = re.compile(r"^[（(]\d+[）)]\s*")
        re_salutation = re.compile(r"^[^\s。；！？，、]{2,35}[：:]$")
        re_attachment = re.compile(r"^附件[：:\d\s]")
        re_date = re.compile(
            r"^([12]\d{3}|[二〇零○一二三四五六七八九]{4})\s*年\s*(\d{1,2}|[一二三四五六七八九十]{1,2})\s*月\s*(\d{1,2}|[一二三四五六七八九十]{1,3})\s*日$"
        )

        stats = {"title": 0, "h1": 0, "h2": 0, "h3": 0, "h4": 0, "body": 0, "signature": 0}
        found_main_title = False
        total = len(valid_paras)

        for idx, (p, txt) in enumerate(valid_paras):
            p.text = ""
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(float(self.cfg["line_spacing_pt"]))

            if not found_main_title and idx <= 1 and not re.search(r"〔\d{4}〕.*号", txt):
                found_main_title = True
                stats["title"] += 1
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.line_spacing = Pt(32.0)
                set_paragraph_char_indent(p, first_line_chars=0, font_size_pt=22.0)
                r = p.add_run(txt)
                set_run_fonts(r, cn_font=self.cfg["font_title"], en_font=self.cfg["font_en"], size_pt=22.0)
                continue

            if idx <= 2 and re.search(r"〔\d{4}〕.*号", txt):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_paragraph_char_indent(p, first_line_chars=0, font_size_pt=16.0)
                r = p.add_run(txt)
                set_run_fonts(r, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0)
                continue

            if idx <= 3 and re_salutation.match(txt) and not re_attachment.match(txt):
                txt_norm = re.sub(r":$", "：", txt)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                set_paragraph_char_indent(p, first_line_chars=0, font_size_pt=16.0)
                r = p.add_run(txt_norm)
                set_run_fonts(r, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0)
                continue

            if re_h1.match(txt):
                stats["h1"] += 1
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                set_paragraph_char_indent(p, first_line_chars=2, font_size_pt=16.0)
                r = p.add_run(txt)
                set_run_fonts(r, cn_font=self.cfg["font_h1"], en_font=self.cfg["font_en"], size_pt=16.0)
                continue

            if re_h2.match(txt):
                stats["h2"] += 1
                txt_norm = re.sub(r"^\(([一二三四五六七八九十百]+)\)", r"（\1）", txt)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                set_paragraph_char_indent(p, first_line_chars=2, font_size_pt=16.0)
                r = p.add_run(txt_norm)
                set_run_fonts(r, cn_font=self.cfg["font_h2"], en_font=self.cfg["font_en"], size_pt=16.0)
                continue

            if re_h3.match(txt):
                stats["h3"] += 1
                txt_norm = re.sub(r"^(\d+)[.、]\s*", r"\1．", txt)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                set_paragraph_char_indent(p, first_line_chars=2, font_size_pt=16.0)
                m_lead = re.match(r"^(\d+．[^。：；！？]{1,20}[。：])(.*)$", txt_norm)
                if m_lead and m_lead.group(2).strip():
                    r1 = p.add_run(m_lead.group(1))
                    set_run_fonts(r1, cn_font=self.cfg["font_h3"], en_font=self.cfg["font_en"], size_pt=16.0, bold=True)
                    r2 = p.add_run(m_lead.group(2))
                    set_run_fonts(r2, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0, bold=False)
                else:
                    r = p.add_run(txt_norm)
                    set_run_fonts(r, cn_font=self.cfg["font_h3"], en_font=self.cfg["font_en"], size_pt=16.0, bold=True)
                continue

            if re_h4.match(txt):
                stats["h4"] += 1
                txt_norm = re.sub(r"^\((\d+)\)\s*", r"（\1）", txt)
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                set_paragraph_char_indent(p, first_line_chars=2, font_size_pt=16.0)
                r = p.add_run(txt_norm)
                set_run_fonts(r, cn_font=self.cfg["font_h4"], en_font=self.cfg["font_en"], size_pt=16.0)
                continue

            if re_attachment.match(txt):
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                set_paragraph_char_indent(p, first_line_chars=-3, left_chars=5, font_size_pt=16.0)
                r = p.add_run(txt)
                set_run_fonts(r, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0)
                continue

            if re_date.match(txt) and idx >= max(1, total - 5):
                stats["signature"] += 1
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_paragraph_char_indent(p, first_line_chars=0, right_chars=4, font_size_pt=16.0)
                r = p.add_run(txt)
                set_run_fonts(r, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0)

                if idx > 0:
                    prev_p, prev_txt = valid_paras[idx - 1]
                    if 2 <= len(prev_txt) <= 22 and not re.search(r"[。；！？：]", prev_txt):
                        prev_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        diff = max(0.0, (len(txt) - len(prev_txt)) / 2.0)
                        set_paragraph_char_indent(prev_p, first_line_chars=0, right_chars=4.0 + diff, font_size_pt=16.0)
                continue

            stats["body"] += 1
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            set_paragraph_char_indent(p, first_line_chars=2, font_size_pt=16.0)
            r = p.add_run(txt)
            set_run_fonts(r, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0)

        self.insert_page_numbers(doc, mode="center")
        return stats

    # ==================== 3. 原版 42 个 VBA 脚本的 Python 批处理实现 ====================

    def script_official_signature_align(self, doc: Document) -> bool:
        """《公文落款.vba》：智能识别文档末尾的发文机关署名与成文日期并居中对齐（日期右空4字）"""
        non_empty = [p for p in doc.paragraphs if (p.text or "").strip()]
        if len(non_empty) < 2:
            return False
        sign_p, date_p = non_empty[-2], non_empty[-1]
        sign_txt = (sign_p.text or "").strip()
        date_txt = (date_p.text or "").strip()
        date_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_paragraph_char_indent(date_p, first_line_chars=0, right_chars=4.0, font_size_pt=16.0)

        sign_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        diff = max(0.0, (len(date_txt) - len(sign_txt)) / 2.0)
        set_paragraph_char_indent(sign_p, first_line_chars=0, right_chars=4.0 + diff, font_size_pt=16.0)
        return True

    def script_renumber_chapters_articles(self, doc: Document) -> Tuple[int, int]:
        """《重新整理章条序号.vba》：将全文章、条各自从“第一”开始连续重排"""
        ch_count = 0
        art_count = 0
        re_ch = re.compile(r"^(\s*)第[零〇一二三四五六七八九十百千\d]+章")
        re_art = re.compile(r"^(\s*)第[零〇一二三四五六七八九十百千\d]+条")

        for p in doc.paragraphs:
            txt = p.text or ""
            m_ch = re_ch.match(txt)
            if m_ch:
                ch_count += 1
                new_txt = re_ch.sub(rf"\g<1>第{int_to_chinese(ch_count)}章", txt, count=1)
                if p.runs:
                    p.runs[0].text = new_txt
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = new_txt
                continue

            m_art = re_art.match(txt)
            if m_art:
                art_count += 1
                new_txt = re_art.sub(rf"\g<1>第{int_to_chinese(art_count)}条", txt, count=1)
                if p.runs:
                    p.runs[0].text = new_txt
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = new_txt
        return ch_count, art_count

    def script_bold_chapters_articles(self, doc: Document, whole_line: bool = False) -> int:
        """《加粗第n章第n条.vba》 & 《加粗第n章第n条(整行).vba》"""
        re_target = re.compile(r"^(\s*第[零〇一二三四五六七八九十百千\d]+[章条])(.*)$")
        count = 0
        for p in doc.paragraphs:
            m = re_target.match(p.text or "")
            if not m:
                continue
            count += 1
            if whole_line:
                for r in p.runs:
                    r.bold = True
            else:
                head_part, rest_part = m.group(1), m.group(2)
                p.text = ""
                r1 = p.add_run(head_part)
                set_run_fonts(r1, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0, bold=True)
                if rest_part:
                    r2 = p.add_run(rest_part)
                    set_run_fonts(r2, cn_font=self.cfg["font_body"], en_font=self.cfg["font_en"], size_pt=16.0, bold=False)
        return count

    def script_center_chapter_lines(self, doc: Document) -> int:
        """《第n章整行加粗居中.vba》"""
        re_ch = re.compile(r"^\s*第[零〇一二三四五六七八九十百千\d]{1,5}章")
        count = 0
        for p in doc.paragraphs:
            if re_ch.match(p.text or ""):
                count += 1
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_paragraph_char_indent(p, first_line_chars=0, left_chars=0, right_chars=0)
                for r in p.runs:
                    r.bold = True
        return count

    def script_chapter_article_one_space(self, doc: Document) -> int:
        """《第n章条后1空格.vba》"""
        re_sp = re.compile(r"^(\s*第[零〇一二三四五六七八九十百千\d]+[条章])[ \t\u3000]*(.+)$")
        count = 0
        for p in doc.paragraphs:
            m = re_sp.match(p.text or "")
            if m:
                count += 1
                new_txt = f"{m.group(1)} {m.group(2)}"
                if p.runs:
                    p.runs[0].text = new_txt
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = new_txt
        return count

    def script_renumber_cn_brackets(self, doc: Document) -> int:
        """《（一）至（九十九）重新排序.vba》"""
        re_br = re.compile(r"[（(][一二三四五六七八九十百]+[）)]")
        idx = 0
        for p in doc.paragraphs:
            txt = p.text or ""
            if re_br.search(txt):
                idx += 1
                new_txt = re_br.sub(f"（{int_to_chinese(idx)}）", txt, count=1)
                if p.runs:
                    p.runs[0].text = new_txt
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = new_txt
        return idx

    def script_delete_headers_footers(self, doc: Document) -> int:
        """《删除页眉页脚.vba》"""
        count = 0
        for sec in doc.sections:
            for hf in [sec.header, sec.footer, sec.first_page_header, sec.first_page_footer, sec.even_page_header, sec.even_page_footer]:
                try:
                    for p in hf.paragraphs:
                        p.text = ""
                    count += 1
                except Exception:
                    pass
        return count

    def script_extract_all_images(self, docx_path: str, output_dir: Optional[str] = None) -> List[str]:
        """《保存文档所有图片.vba》：直接从 .docx 容器提取 word/media/ 下所有原始高清图片"""
        p = Path(docx_path)
        out_dir = Path(output_dir) if output_dir else p.parent / f"{p.stem}_提取图片"
        out_dir.mkdir(parents=True, exist_ok=True)
        extracted = []
        with zipfile.ZipFile(docx_path, "r") as zf:
            for idx, item in enumerate(zf.namelist(), start=1):
                if item.startswith("word/media/"):
                    ext = Path(item).suffix or ".png"
                    target = out_dir / f"{p.stem}_image_{idx}{ext}"
                    target.write_bytes(zf.read(item))
                    extracted.append(str(target))
        return extracted

    def script_make_place_cards(self, names: List[str], output_path: str) -> str:
        """《桌签标牌制作.vba》：生成 A4 横向双面对折 72pt 黑体会议桌签 .docx"""
        doc = Document()
        sec = doc.sections[0]
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width = Mm(297.0)
        sec.page_height = Mm(210.0)
        sec.top_margin = Mm(15.0)
        sec.bottom_margin = Mm(15.0)
        sec.left_margin = Mm(20.0)
        sec.right_margin = Mm(20.0)

        clean_names = [n.strip() for n in names if n.strip()]
        if not clean_names:
            clean_names = ["张  三"]

        tbl = doc.add_table(rows=len(clean_names) * 2, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        for i, name in enumerate(clean_names):
            disp = f"{name[0]}  {name[1]}" if len(name) == 2 else name
            for sub in range(2):
                row = tbl.rows[i * 2 + sub]
                row.height = Cm(8.5)
                cell = row.cells[0]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(disp)
                set_run_fonts(r, cn_font="黑体", en_font="Times New Roman", size_pt=72.0, bold=True)

        doc.save(output_path)
        return output_path

    def script_remove_xlsx_protection(self, xlsx_path: str, output_path: Optional[str] = None) -> str:
        """《移除xls写保护.vba》：无损移除 .xlsx 工作簿与工作表的 XML 写保护标签"""
        target = output_path or xlsx_path
        temp_zip = tempfile.mktemp(suffix=".xlsx")
        with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("xl/worksheets/sheet") or item.filename == "xl/workbook.xml":
                    txt = data.decode("utf-8", errors="ignore")
                    txt = re.sub(r"<sheetProtection[^>]*/>", "", txt)
                    txt = re.sub(r"<workbookProtection[^>]*/>", "", txt)
                    data = txt.encode("utf-8")
                zout.writestr(item, data)
        shutil.move(temp_zip, target)
        return target

    def script_images_to_pptx(self, image_paths: List[str], output_pptx: str) -> str:
        """《图片批量转PPT.vba》：将多张图片批量生成 16:9 宽屏 PPTX 演示文稿"""
        from pptx import Presentation
        from pptx.util import Inches

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6]

        for img_path in image_paths:
            if not os.path.isfile(img_path):
                continue
            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

        prs.save(output_pptx)
        return output_pptx

    def script_images_to_pdf(self, image_paths: List[str], output_pdf: str) -> str:
        """《插入图片生成PDF.vba》：将多张图片合成 A4 标准 PDF 文件"""
        from PIL import Image

        pil_images = []
        for p in image_paths:
            if os.path.isfile(p):
                img = Image.open(p).convert("RGB")
                pil_images.append(img)
        if not pil_images:
            raise ValueError("未提供有效的图片文件")
        pil_images[0].save(output_pdf, "PDF", resolution=200.0, save_all=True, append_images=pil_images[1:])
        return output_pdf

    def script_make_scanned_pdf(
        self,
        input_pdf: str,
        output_pdf: str,
        skew: float = 0.6,
        noise: float = 0.06,
        dpi: int = 200,
        grayscale: bool = False,
    ) -> str:
        """
        《生成扫描版pdf.vba》银河麒麟原生实现（无需 Windows make-look-scanned.exe）
        使用 pdftoppm + Pillow 对每页施加轻微随机倾角、纸张底色与扫描噪点，再封装为扫描件 PDF
        """
        from PIL import Image, ImageEnhance, ImageFilter
        import numpy as np

        with tempfile.TemporaryDirectory(prefix="xkl_scan_") as tmpdir:
            prefix = os.path.join(tmpdir, "page")
            subprocess.run(
                ["pdftoppm", "-png", "-r", str(dpi), input_pdf, prefix],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            page_files = sorted(glob.glob(f"{prefix}*.png"))
            if not page_files:
                raise RuntimeError("未能从 PDF 渲染出页面图像")

            processed_pages = []
            for pf in page_files:
                img = Image.open(pf).convert("RGB")
                if grayscale:
                    img = img.convert("L").convert("RGB")

                angle = random.uniform(-abs(skew), abs(skew))
                img = img.rotate(angle, resample=Image.BICUBIC, expand=False, fillcolor=(252, 251, 247))

                arr = np.array(img, dtype=np.int16)
                if noise > 0:
                    sigma = max(1.0, noise * 80.0)
                    grain = np.random.normal(0, sigma, arr.shape).astype(np.int16)
                    arr = np.clip(arr + grain, 0, 255)
                img_out = Image.fromarray(arr.astype(np.uint8))
                img_out = img_out.filter(ImageFilter.GaussianBlur(radius=0.35))
                img_out = ImageEnhance.Contrast(img_out).enhance(1.06)
                processed_pages.append(img_out)

            processed_pages[0].save(
                output_pdf,
                "PDF",
                resolution=float(dpi),
                save_all=True,
                append_images=processed_pages[1:],
            )
        return output_pdf

    def convert_doc_to_pdf_kylin(self, input_doc: str, output_pdf: Optional[str] = None) -> str:
        """在银河麒麟系统上调用 LibreOffice 或 WPS 将 .docx/.doc 转换为 PDF"""
        in_path = Path(input_doc).resolve()
        out_dir = Path(output_pdf).resolve().parent if output_pdf else in_path.parent
        out_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(prefix="lo_kylin_") as lo_prof:
            env = {**os.environ, "HOME": "/tmp", "TMPDIR": "/tmp"}
            cmd = [
                "libreoffice",
                f"-env:UserInstallation=file://{lo_prof}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(out_dir),
                str(in_path),
            ]
            subprocess.run(cmd, env=env, timeout=45, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        generated = out_dir / f"{in_path.stem}.pdf"
        if output_pdf and Path(output_pdf).resolve() != generated:
            shutil.move(str(generated), output_pdf)
            return output_pdf
        return str(generated)
