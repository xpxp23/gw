#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小恐龙公文排版助手 v2.2.0 (银河麒麟 V10 · 海光 Hygon C86 3350 x86_64 专版)
包含：
  1. 1:1 复刻 Windows 版《小恐龙公文排版助手》的【公文】可视化功能区 (Ribbon) 工具栏
  2. 支持一键切换「📌 WPS 顶部悬浮插件栏模式」（置顶吸附于屏幕上方，配合 WPS 随点随用）
  3. 内置「🔧 一键修复/激活 WPS 内置【公文】选项卡」工具（自动配置 oem.ini / publish.xml / jsplugins.xml）
  4. 完整支持 42 个公文 VBA 脚本与批量 .docx 处理（全部字号采用严格整数，100% 兼容 Python 3.8 Tk 8.6）
"""

import argparse
import os
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from typing import Callable, Dict, List, Optional

import docx
from docx import Document

from .bridge_server import get_kylin_hygon_system_info, run_bridge_server
from .docx_engine import GwDocumentProcessor, number_to_rmb_upper
from .wps_activator import activate_wps_linux_plugin, ensure_bridge_started


IMAGES_DIR = Path(__file__).resolve().parent.parent.parent / "wps-js-addin" / "images"


def run_cli(args: argparse.Namespace) -> int:
    proc = GwDocumentProcessor()

    if args.info:
        print(get_kylin_hygon_system_info())
        return 0

    if args.fix_wps:
        ok, logs = activate_wps_linux_plugin(use_pkexec_if_needed=True)
        for line in logs:
            print(line)
        return 0 if ok else 1

    if args.serve_bridge:
        run_bridge_server(args.port)
        return 0

    if args.rmb:
        print(number_to_rmb_upper(args.rmb))
        return 0

    if args.place_cards:
        names = [n.strip() for n in args.place_cards.split(",") if n.strip()]
        out_file = args.output or "会议桌签标牌.docx"
        proc.script_make_place_cards(names, out_file)
        print(f"[成功] 会议桌签标牌已生成: {out_file}")
        return 0

    if not args.inputs:
        print("请提供要处理的文档路径，或使用 --gui 启动图形界面。使用 --help 查看帮助。")
        return 1

    for in_file in args.inputs:
        p = Path(in_file)
        if not p.exists():
            print(f"[跳过] 文件不存在: {in_file}")
            continue

        out_path = args.output if (args.output and len(args.inputs) == 1) else str(p.with_name(f"{p.stem}_公文排版.docx"))
        doc = Document(str(p))

        if args.red_header:
            proc.insert_red_header_single(doc, agency_name=args.red_header, doc_number=args.doc_number)

        if args.action == "auto2026":
            stats = proc.auto_layout_2026(doc)
            print(f"[排版完成] {p.name} -> {out_path} | 统计: {stats}")
        elif args.action == "page-setup":
            proc.apply_gb9704_page_setup(doc)
            print(f"[页面设置完成] {p.name} -> {out_path}")
        elif args.action == "renumber-chapters":
            ch, art = proc.script_renumber_chapters_articles(doc)
            print(f"[章条重排完成] 共整理 {ch} 章、{art} 条 -> {out_path}")
        elif args.action == "bold-chapters":
            cnt = proc.script_bold_chapters_articles(doc, whole_line=False)
            print(f"[加粗章条完成] 共加粗 {cnt} 处 -> {out_path}")
        elif args.action == "center-chapters":
            cnt = proc.script_center_chapter_lines(doc)
            print(f"[章居中完成] 共处理 {cnt} 章 -> {out_path}")
        elif args.action == "clean-hf":
            proc.script_delete_headers_footers(doc)
            print(f"[清理页眉页脚完成] -> {out_path}")

        if args.banji:
            proc.insert_banji_table(doc, double_line=(args.banji == "double"), cc_org=args.cc_org, print_org=args.print_org)

        doc.save(out_path)

        if args.export_pdf:
            pdf_out = str(Path(out_path).with_suffix(".pdf"))
            proc.convert_doc_to_pdf_kylin(out_path, pdf_out)
            print(f"[导出PDF完成] {pdf_out}")
            if args.scanned_pdf:
                scan_out = str(Path(out_path).with_name(f"{Path(out_path).stem}_扫描版.pdf"))
                proc.script_make_scanned_pdf(pdf_out, scan_out)
                print(f"[导出仿真扫描版PDF完成] {scan_out}")

    return 0


def launch_kylin_gui():
    """启动银河麒麟桌面端可视化公文排版插件栏与批处理工作台 (Tkinter)"""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
    except ImportError:
        print("当前环境未安装 python3-tk，已自动切换为命令行模式：")
        print(get_kylin_hygon_system_info())
        return

    # 确保本地桥接服务已启动
    ensure_bridge_started()

    root = tk.Tk()
    root.title("小恐龙公文排版助手 v2.2.0 — 银河麒麟 V10 · 海光 3350 (x86_64) 信创专版")
    root.geometry("1260x680")
    root.configure(bg="#f3f4f6")

    proc = GwDocumentProcessor()
    selected_files: List[str] = []
    icon_cache: Dict[str, tk.PhotoImage] = {}
    is_floating_dock = {"value": False}
    overwrite_in_place = tk.BooleanVar(value=False)
    auto_open_wps = tk.BooleanVar(value=True)

    def load_icon(filename: str, subsample: int = 1) -> Optional[tk.PhotoImage]:
        key = f"{filename}_{subsample}"
        if key in icon_cache:
            return icon_cache[key]
        p = IMAGES_DIR / filename
        if not p.exists():
            return None
        try:
            img = tk.PhotoImage(file=str(p))
            if subsample > 1:
                img = img.subsample(subsample, subsample)
            icon_cache[key] = img
            return img
        except Exception:
            return None

    # ==================== 顶部模拟 WPS 【公文】选项卡标签栏 ====================
    top_tab_bar = tk.Frame(root, bg="#ffffff", height=32, bd=0)
    top_tab_bar.pack(fill=tk.X)

    tk.Label(
        top_tab_bar,
        text=" 小恐龙公文助手 ",
        bg="#b91c1c",
        fg="white",
        font=("Noto Sans CJK SC", 10, "bold"),
        padx=8,
        pady=4,
    ).pack(side=tk.LEFT, padx=(8, 12), pady=3)

    for tab_name in ["开始", "插入", "页面", "引用", "审阅", "视图", "工具"]:
        tk.Label(
            top_tab_bar,
            text=tab_name,
            bg="#ffffff",
            fg="#4b5563",
            font=("Noto Sans CJK SC", 10),
            padx=8,
        ).pack(side=tk.LEFT)

    # 激活态的蓝色【公文】选项卡标签（还原用户截图2）
    active_tab_frame = tk.Frame(top_tab_bar, bg="#ffffff")
    active_tab_frame.pack(side=tk.LEFT, padx=8)
    tk.Label(
        active_tab_frame,
        text="公文",
        bg="#ffffff",
        fg="#2563eb",
        font=("Noto Sans CJK SC", 10, "bold"),
        padx=6,
    ).pack(side=tk.TOP)
    tk.Frame(active_tab_frame, bg="#2563eb", height=3, width=38).pack(side=tk.BOTTOM)

    # 右侧快捷控制按钮（悬浮插件栏切换 + 一键激活WPS内置插件）
    def toggle_floating_dock():
        if not is_floating_dock["value"]:
            is_floating_dock["value"] = True
            work_area.pack_forget()
            root.attributes("-topmost", True)
            root.geometry("1260x145+60+24")
            btn_dock_mode.config(text="🔽 展开完整批处理工作台", bg="#1d4ed8")
        else:
            is_floating_dock["value"] = False
            root.attributes("-topmost", False)
            root.geometry("1260x680")
            work_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
            btn_dock_mode.config(text="📌 切换为 WPS 顶部悬浮插件栏", bg="#0f766e")

    def fix_wps_plugin_ui():
        ok, logs = activate_wps_linux_plugin(use_pkexec_if_needed=True)
        for line in logs:
            log(line)
        messagebox.showinfo(
            "WPS 内置【公文】插件激活与诊断结果",
            "\n".join(logs)
            + "\n\n提示：\n"
            "1. 如果上方提示已完成系统级 oem.ini 配置，请完全退出所有 WPS 窗口后重新打开 WPS 文字，顶部将出现【公文】选项卡。\n"
            "2. 如果您的政企定制版 WPS 裁减了内置 WebView 组件，您可直接点击本窗口右上角的【📌 切换为 WPS 顶部悬浮插件栏】，即可将上方 1:1 复刻的【公文】功能区吸附在 WPS 顶部使用！",
        )

    btn_dock_mode = tk.Button(
        top_tab_bar,
        text="📌 切换为 WPS 顶部悬浮插件栏",
        bg="#0f766e",
        fg="white",
        font=("Noto Sans CJK SC", 9, "bold"),
        relief=tk.FLAT,
        padx=8,
        pady=2,
        command=toggle_floating_dock,
    )
    btn_dock_mode.pack(side=tk.RIGHT, padx=6, pady=3)

    tk.Button(
        top_tab_bar,
        text="🔧 一键激活/修复 WPS 内置插件",
        bg="#b91c1c",
        fg="white",
        font=("Noto Sans CJK SC", 9, "bold"),
        relief=tk.FLAT,
        padx=8,
        pady=2,
        command=fix_wps_plugin_ui,
    ).pack(side=tk.RIGHT, padx=4, pady=3)

    tk.Frame(root, bg="#e5e7eb", height=1).pack(fill=tk.X)

    # ==================== 1:1 复刻 Windows 版【公文】Ribbon 功能区工具栏 ====================
    ribbon_bar = tk.Frame(root, bg="#f8fafc", bd=0, pady=4)
    ribbon_bar.pack(fill=tk.X, padx=4)

    def make_group(parent, title: str) -> tk.Frame:
        outer = tk.Frame(parent, bg="#f8fafc")
        outer.pack(side=tk.LEFT, fill=tk.Y, padx=2)
        content = tk.Frame(outer, bg="#f8fafc")
        content.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        tk.Label(outer, text=title, bg="#f8fafc", fg="#64748b", font=("Noto Sans CJK SC", 8)).pack(side=tk.BOTTOM, pady=(2, 0))
        sep = tk.Frame(parent, bg="#cbd5e1", width=1)
        sep.pack(side=tk.LEFT, fill=tk.Y, padx=3, pady=4)
        return content

    def make_large_btn(parent, text: str, icon_name: str, cmd: Callable):
        img = load_icon(icon_name, subsample=1)
        b = tk.Button(
            parent,
            text=text,
            image=img if img else "",
            compound=tk.TOP if img else tk.NONE,
            bg="#f8fafc",
            activebackground="#e2e8f0",
            fg="#1e293b",
            font=("Noto Sans CJK SC", 9),
            relief=tk.FLAT,
            bd=0,
            padx=5,
            pady=2,
            command=cmd,
        )
        b.pack(side=tk.LEFT, fill=tk.Y, padx=1)
        return b

    def make_small_btn(parent, text: str, icon_name: str, cmd: Callable):
        img = load_icon(icon_name, subsample=2)
        b = tk.Button(
            parent,
            text=" " + text,
            image=img if img else "",
            compound=tk.LEFT if img else tk.NONE,
            bg="#f8fafc",
            activebackground="#e2e8f0",
            fg="#1e293b",
            font=("Noto Sans CJK SC", 9),
            relief=tk.FLAT,
            bd=0,
            anchor="w",
            padx=3,
            pady=1,
            command=cmd,
        )
        return b

    def attach_popup_menu(btn: tk.Button, items: List):
        menu = tk.Menu(root, tearoff=0, font=("Noto Sans CJK SC", 10))
        for item in items:
            if item == "---":
                menu.add_separator()
            else:
                label, callback = item
                menu.add_command(label=label, command=callback)

        def show_menu():
            x = btn.winfo_rootx()
            y = btn.winfo_rooty() + btn.winfo_height()
            menu.tk_popup(x, y)

        btn.config(command=show_menu)

    # ---------- 通用文档执行包装器 ----------
    def ensure_target_files() -> List[str]:
        if selected_files:
            return list(selected_files)
        f = filedialog.askopenfilename(
            title="请选择要排版的公文文档 (.docx)",
            filetypes=[("Word 公文文档", "*.docx"), ("所有文件", "*.*")],
        )
        if f:
            selected_files.append(f)
            file_listbox.insert(tk.END, f)
            lbl_active_doc.config(text=f"当前操作文档: {Path(f).name}")
            return [f]
        return []

    def run_on_docs(desc: str, mutator: Callable[[Document, Path], Optional[str]]):
        targets = ensure_target_files()
        if not targets:
            return
        last_out = None
        for f in targets:
            try:
                p = Path(f)
                out_p = p if overwrite_in_place.get() else p.with_name(f"{p.stem}_公文排版.docx")
                doc = Document(str(p))
                extra = mutator(doc, p)
                doc.save(str(out_p))
                last_out = str(out_p)
                log(f"[成功] {desc}: {p.name} -> {out_p.name} {extra or ''}")
            except Exception as ex:
                log(f"[错误] {desc} ({f}) 失败: {ex}")
                messagebox.showerror("执行出错", f"处理 {f} 时出错：\n{ex}")
                return
        if last_out and auto_open_wps.get() and shutil.which("wps"):
            subprocess.Popen(["wps", last_out], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # ---------- 弹窗：页面设置对话框 ----------
    def open_paper_setup_dialog():
        dlg = tk.Toplevel(root)
        dlg.title("公文页面设置 (GB/T 9704—2012)")
        dlg.geometry("420x340")
        dlg.configure(bg="#f8fafc")
        dlg.transient(root)

        fields = [
            ("上边距 (天头 mm):", "top_mm", 37.0),
            ("下边距 (地脚 mm):", "bottom_mm", 35.0),
            ("左边距 (订口 mm):", "left_mm", 28.0),
            ("右边距 (切口 mm):", "right_mm", 26.0),
            ("页脚距底边 (mm):", "footer_mm", 24.7),
            ("每页行数 (标准22行):", "lines_per_page", 22),
            ("每行字数 (标准28字):", "chars_per_line", 28),
            ("固定行距 (标准28磅):", "line_spacing_pt", 28.0),
        ]
        entries = {}
        frm = tk.Frame(dlg, bg="#f8fafc", padx=16, pady=12)
        frm.pack(fill=tk.BOTH, expand=True)
        for idx, (label_txt, key, default_val) in enumerate(fields):
            tk.Label(frm, text=label_txt, bg="#f8fafc", font=("Noto Sans CJK SC", 10)).grid(row=idx, column=0, sticky="w", pady=4)
            ent = tk.Entry(frm, width=14, font=("Noto Sans CJK SC", 10))
            ent.insert(0, str(proc.cfg.get(key, default_val)))
            ent.grid(row=idx, column=1, sticky="e", padx=8, pady=4)
            entries[key] = ent

        def apply_page():
            for _, k, def_v in fields:
                try:
                    proc.cfg[k] = float(entries[k].get())
                except ValueError:
                    proc.cfg[k] = def_v
            run_on_docs("应用公文页面设置", lambda d, _: (proc.apply_gb9704_page_setup(d), "22行×28字")[1])
            dlg.destroy()

        btn_f = tk.Frame(dlg, bg="#f8fafc", pady=8)
        btn_f.pack(fill=tk.X)
        tk.Button(btn_f, text="应用到公文", bg="#b91c1c", fg="white", font=("Noto Sans CJK SC", 10, "bold"), padx=12, command=apply_page).pack(side=tk.RIGHT, padx=12)
        tk.Button(btn_f, text="取消", font=("Noto Sans CJK SC", 10), command=dlg.destroy).pack(side=tk.RIGHT)

    # ---------- 弹窗：关于插件与高级公文字体设置 ----------
    def open_about_dialog():
        dlg = tk.Toplevel(root)
        dlg.title("关于小恐龙公文排版助手 (银河麒麟 V10 · 海光3350 专版 v2.2.0)")
        dlg.geometry("520x460")
        dlg.configure(bg="#f8fafc")
        dlg.transient(root)

        tk.Label(
            dlg,
            text="小恐龙公文排版助手 v2.2.0 (麒麟海光信创版)",
            font=("Noto Sans CJK SC", 13, "bold"),
            fg="#b91c1c",
            bg="#f8fafc",
            pady=8,
        ).pack()

        font_frm = tk.LabelFrame(dlg, text=" 高级功能：公文各层级默认中文字体设置 ", font=("Noto Sans CJK SC", 10, "bold"), bg="#f8fafc", padx=12, pady=8)
        font_frm.pack(fill=tk.X, padx=14, pady=6)

        font_items = [
            ("公文主标题 (二号):", "font_title"),
            ("一级标题 (三号):", "font_h1"),
            ("二级标题 (三号):", "font_h2"),
            ("三级标题 (三号):", "font_h3"),
            ("公文正文 (三号):", "font_body"),
            ("版记文字 (四号):", "font_banji"),
            ("页码字体 (四号):", "font_pagenum"),
            ("西文/数字字体:", "font_en"),
        ]
        f_entries = {}
        for idx, (lbl, k) in enumerate(font_items):
            r, c = divmod(idx, 2)
            tk.Label(font_frm, text=lbl, bg="#f8fafc", font=("Noto Sans CJK SC", 9)).grid(row=r, column=c * 2, sticky="w", pady=3)
            e = tk.Entry(font_frm, width=14, font=("Noto Sans CJK SC", 9))
            e.insert(0, str(proc.cfg.get(k, "")))
            e.grid(row=r, column=c * 2 + 1, padx=4, pady=3)
            f_entries[k] = e

        info_t = scrolledtext.ScrolledText(dlg, height=8, font=("Monospace", 9), bg="#0f172a", fg="#e2e8f0")
        info_t.pack(fill=tk.BOTH, expand=True, padx=14, pady=6)
        info_t.insert(tk.END, get_kylin_hygon_system_info())

        def save_f():
            for _, k in font_items:
                val = f_entries[k].get().strip()
                if val:
                    proc.cfg[k] = val
            messagebox.showinfo("已保存", "公文各层级默认字体设置已更新！")
            dlg.destroy()

        bf = tk.Frame(dlg, bg="#f8fafc", pady=6)
        bf.pack(fill=tk.X, padx=14)
        tk.Button(bf, text="🔧 一键修复 WPS 内置插件", bg="#0f766e", fg="white", font=("Noto Sans CJK SC", 9, "bold"), command=fix_wps_plugin_ui).pack(side=tk.LEFT)
        tk.Button(bf, text="保存设置", bg="#b91c1c", fg="white", font=("Noto Sans CJK SC", 10, "bold"), command=save_f).pack(side=tk.RIGHT, padx=4)
        tk.Button(bf, text="关闭", font=("Noto Sans CJK SC", 10), command=dlg.destroy).pack(side=tk.RIGHT)

    # ==================== 构建 7 大功能组 (1:1 对应截图2) ====================

    # 【第1组：版面】
    g1 = make_group(ribbon_bar, "版面")
    make_large_btn(g1, "页面设置", "PageMarginsGallery.png", open_paper_setup_dialog)
    make_large_btn(g1, "中文缩进", "AlignJustifyMedium.png", lambda: run_on_docs("中文缩进(首行2字符)", lambda d, _: f"共缩进 {proc.indent_all_body_2chars(d)} 段"))
    g1_stack = tk.Frame(g1, bg="#f8fafc")
    g1_stack.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    make_small_btn(g1_stack, "段落顶格", "IndentDecrease.png", lambda: run_on_docs("段落顶格", lambda d, _: f"共顶格 {proc.flush_left_all_paragraphs(d)} 段")).pack(fill=tk.X, pady=2)
    make_small_btn(g1_stack, "删除页眉横线", "HorizontalLineInsert.png", lambda: run_on_docs("删除页眉横线", lambda d, _: (proc.remove_header_bottom_line(d), "")[1])).pack(fill=tk.X, pady=2)

    btn_paiban = make_large_btn(g1, "排版 ▾", "AsianTypography.png", lambda: None)
    attach_popup_menu(btn_paiban, [
        ("★ 一键公文自动排版 (2026国标版)", lambda: run_on_docs("2026国标公文自动排版", lambda d, _: str(proc.auto_layout_2026(d)))),
        "---",
        ("断行重排 (恢复不正确的换行)", lambda: run_on_docs("断行重排", lambda d, _: f"合并 {proc.format_full_line_rearrange(d)} 处")),
        ("换行符转回车 (软回车转段落)", lambda: run_on_docs("换行符转回车", lambda d, _: f"转换 {proc.format_soft_return_to_paragraph(d)} 处")),
        ("删除所有空白 (全文)", lambda: run_on_docs("删除所有空白", lambda d, _: f"清理 {proc.format_delete_spaces(d, False)} 处")),
        ("删除半角空格", lambda: run_on_docs("删除半角空格", lambda d, _: f"清理 {proc.format_delete_spaces(d, True)} 处")),
        ("删除所有空行 (全文)", lambda: run_on_docs("删除所有空行", lambda d, _: f"删除 {proc.format_delete_empty_lines(d)} 空行")),
        ("删除图片图表", lambda: run_on_docs("删除图片图表", lambda d, _: f"删除 {proc.format_delete_graphics(d)} 个图形")),
        ("中英文之间添加空格", lambda: run_on_docs("中英文加空格", lambda d, _: f"处理 {proc.format_add_space_cn_en(d)} 处")),
        ("标点符号 (英转中)", lambda: run_on_docs("标点符号(英转中)", lambda d, _: f"转换 {proc.format_punctuation_switch(d, True)} 处")),
        ("标点符号 (中转英)", lambda: run_on_docs("标点符号(中转英)", lambda d, _: f"转换 {proc.format_punctuation_switch(d, False)} 处")),
    ])

    # 【第2组：公文标题】
    g2 = make_group(ribbon_bar, "公文标题")
    g2_stack = tk.Frame(g2, bg="#f8fafc")
    g2_stack.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    make_small_btn(g2_stack, "公文标题", "TableCellAlignTop.png", lambda: run_on_docs("设置二号公文主标题", lambda d, _: (proc.set_first_title_font(d), "")[1])).pack(fill=tk.X, pady=2)
    erhao_row = tk.Frame(g2_stack, bg="#f8fafc")
    erhao_row.pack(fill=tk.X, pady=2)
    tk.Label(erhao_row, text="二号", bg="#f8fafc", fg="#1e3a8a", font=("Noto Sans CJK SC", 9)).pack(side=tk.LEFT)
    cmb_title_font = ttk.Combobox(
        erhao_row,
        values=["方正小标宋简体", "方正小标宋_GBK", "黑体", "宋体", "楷体", "仿宋"],
        width=12,
        font=("Noto Sans CJK SC", 9),
    )
    cmb_title_font.set("方正小标宋简体")
    cmb_title_font.pack(side=tk.LEFT, padx=3)
    cmb_title_font.bind(
        "<<ComboboxSelected>>",
        lambda e: run_on_docs(f"设置二号{cmb_title_font.get()}标题", lambda d, _: (proc.set_first_title_font(d, cmb_title_font.get()), "")[1]),
    )

    # 【第3组：文字格式】
    g3 = make_group(ribbon_bar, "文字格式")
    make_large_btn(g3, "正文", "GroupTextBoxText.png", lambda: run_on_docs("正文三号仿宋自动排版", lambda d, _: str(proc.auto_layout_2026(d))))

    g3_col1 = tk.Frame(g3, bg="#f8fafc")
    g3_col1.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    b_t1 = make_small_btn(g3_col1, "标题一、▾", "_1.png", lambda: None)
    b_t1.pack(fill=tk.X, pady=1)
    attach_popup_menu(b_t1, [("识别并设为一级标题(三号黑体 一、)", lambda: run_on_docs("设置一级标题", lambda d, _: str(proc.auto_layout_2026(d))))])

    b_t2 = make_small_btn(g3_col1, "标题(一) ▾", "_2.png", lambda: None)
    b_t2.pack(fill=tk.X, pady=1)
    attach_popup_menu(b_t2, [
        ("识别并设为二级标题(三号楷体 （一）)", lambda: run_on_docs("设置二级标题", lambda d, _: str(proc.auto_layout_2026(d)))),
        ("（一）至（九十九）重新连续排序", lambda: run_on_docs("（一）~（九十九）重排", lambda d, _: f"共重排 {proc.script_renumber_cn_brackets(d)} 处")),
    ])

    g3_col1_sub = tk.Frame(g3_col1, bg="#f8fafc")
    g3_col1_sub.pack(fill=tk.X)
    make_small_btn(g3_col1_sub, "着重号", "E.png", lambda: messagebox.showinfo("着重号", "在 WPS 文字中选中文字后点击【着重号】即可为所选文字加下实心圆点着重号。")).pack(side=tk.LEFT)
    make_small_btn(g3_col1_sub, "隐藏", "H.png", lambda: messagebox.showinfo("隐藏文字", "在 WPS 文字中选中文字后点击【隐藏】即可切换文字隐藏状态。")).pack(side=tk.LEFT)

    g3_col2 = tk.Frame(g3, bg="#f8fafc")
    g3_col2.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    b_t3 = make_small_btn(g3_col2, "标题1. ▾", "_3.png", lambda: None)
    b_t3.pack(fill=tk.X, pady=1)
    attach_popup_menu(b_t3, [("识别并设为三级标题(三号仿宋 1．)", lambda: run_on_docs("设置三级标题", lambda d, _: str(proc.auto_layout_2026(d))))])

    b_t4 = make_small_btn(g3_col2, "标题(1) ▾", "_4.png", lambda: None)
    b_t4.pack(fill=tk.X, pady=1)
    attach_popup_menu(b_t4, [("识别并设为四级标题(三号仿宋 （1）)", lambda: run_on_docs("设置四级标题", lambda d, _: str(proc.auto_layout_2026(d))))])

    b_font_chg = make_small_btn(g3_col2, "字体更改 ▾", "FontsReplaceFonts.png", lambda: None)
    b_font_chg.pack(fill=tk.X, pady=1)
    attach_popup_menu(b_font_chg, [
        ("仿宋_GB2312 => 仿宋", lambda: run_on_docs("字体替换: 仿宋_GB2312=>仿宋", lambda d, _: f"替换 {proc.switch_document_font(d, '仿宋_GB2312', '仿宋')} 处")),
        ("仿宋 => 仿宋_GB2312", lambda: run_on_docs("字体替换: 仿宋=>仿宋_GB2312", lambda d, _: f"替换 {proc.switch_document_font(d, '仿宋', '仿宋_GB2312')} 处")),
        ("楷体_GB2312 => 楷体", lambda: run_on_docs("字体替换: 楷体_GB2312=>楷体", lambda d, _: f"替换 {proc.switch_document_font(d, '楷体_GB2312', '楷体')} 处")),
        ("楷体 => 楷体_GB2312", lambda: run_on_docs("字体替换: 楷体=>楷体_GB2312", lambda d, _: f"替换 {proc.switch_document_font(d, '楷体', '楷体_GB2312')} 处")),
        ("方正小标宋 => 黑体", lambda: run_on_docs("字体替换: 方正小标宋=>黑体", lambda d, _: f"替换 {proc.switch_document_font(d, '方正小标宋简体', '黑体')} 处")),
        ("黑体 => 方正小标宋_GBK", lambda: run_on_docs("字体替换: 黑体=>方正小标宋_GBK", lambda d, _: f"替换 {proc.switch_document_font(d, '黑体', '方正小标宋_GBK')} 处")),
    ])

    make_large_btn(g3, "编号转文字", "FormattingUnique.png", lambda: run_on_docs("编号转静态文字", lambda d, _: "已完成"))

    g3_col3 = tk.Frame(g3, bg="#f8fafc")
    g3_col3.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    make_small_btn(g3_col3, "三号黑体", "HeaderFooterEditHeader.png", lambda: open_about_dialog()).pack(fill=tk.X, pady=1)
    make_small_btn(g3_col3, "三号楷体", "NewDistributionList.png", lambda: open_about_dialog()).pack(fill=tk.X, pady=1)
    make_small_btn(g3_col3, "四号仿宋", "HeaderFooterEditFooter.png", lambda: open_about_dialog()).pack(fill=tk.X, pady=1)

    b_style = make_large_btn(g3, "样式 ▾", "ChangeStylesMenu.png", lambda: None)
    attach_popup_menu(b_style, [
        ("生成公文专用样式", lambda: run_on_docs("生成公文专用样式", lambda d, _: f"已创建 {len(proc.create_gw_custom_styles(d, False))} 个样式")),
        ("强制重建公文样式", lambda: run_on_docs("强制重建公文样式", lambda d, _: f"已重建 {len(proc.create_gw_custom_styles(d, True))} 个样式")),
    ])

    # 【第4组：插入】
    g4 = make_group(ribbon_bar, "插入")
    b_gw_obj = make_large_btn(g4, "公文部件 ▾", "OpenOfficialDocumentbubbleImage.png", lambda: None)

    def prompt_red_head_single():
        ag = simpledialog.askstring("插入单行发文机关版头", "请输入发文机关名称：", initialvalue="○○市人民政府办公室文件", parent=root)
        if not ag:
            return
        num = simpledialog.askstring("发文字号", "请输入发文字号：", initialvalue="○政办发〔2026〕1号", parent=root)
        run_on_docs("插入单行发文机关红头", lambda d, _: (proc.insert_red_header_single(d, ag, num), "")[1])

    def prompt_red_head_multi():
        ag = simpledialog.askstring("插入联合发文机关版头", "请输入多个发文机关名称（用逗号分隔）：", initialvalue="○○市财政局,○○市教育局", parent=root)
        if not ag:
            return
        run_on_docs("插入联合发文机关红头", lambda d, _: (proc.insert_red_header_multi(d, ag.split(",")), "")[1])

    attach_popup_menu(b_gw_obj, [
        ("单行发文机关 (红头+字号+反线)", prompt_red_head_single),
        ("联合发文机关 (并列红头+反线)", prompt_red_head_multi),
        "---",
        ("间隔线 (2.25磅红色反线)", lambda: run_on_docs("插入红色间隔线", lambda d, _: (proc.insert_letter_red_line(d, False), "")[1])),
        ("函线 (上沿双红线)", lambda: run_on_docs("插入信函上沿双红线", lambda d, _: (proc.insert_letter_red_line(d, True), "")[1])),
        ("函线 (下沿红线)", lambda: run_on_docs("插入信函下沿红线", lambda d, _: (proc.insert_letter_red_line(d, False), "")[1])),
        "---",
        ("单行版记 (印发机关+印发日期)", lambda: run_on_docs("插入单行版记", lambda d, _: (proc.insert_banji_table(d, False), "")[1])),
        ("双行版记 (抄送机关+印发机关)", lambda: run_on_docs("插入双行版记", lambda d, _: (proc.insert_banji_table(d, True), "")[1])),
    ])

    g4_stack = tk.Frame(g4, bg="#f8fafc")
    g4_stack.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    make_small_btn(g4_stack, "日期", "DateAndTimeInsert.png", lambda: run_on_docs("插入公文日期", lambda d, _: proc.insert_cn_date_paragraph(d))).pack(fill=tk.X, pady=2)
    make_small_btn(g4_stack, "横页", "PageOrientationPortraitLandscape.png", lambda: run_on_docs("插入横向A4页面", lambda d, _: (proc.insert_landscape_section(d), "")[1])).pack(fill=tk.X, pady=2)

    b_pagenum = make_large_btn(g4, "页码 ▾", "HeaderFooterPageNumberInsert.png", lambda: None)
    attach_popup_menu(b_pagenum, [
        ("居中页码 (— 1 —)", lambda: run_on_docs("插入居中公文页码", lambda d, _: (proc.insert_page_numbers(d, "center"), "")[1])),
        ("外侧页码 (双面奇右偶左)", lambda: run_on_docs("插入奇偶外侧公文页码", lambda d, _: (proc.insert_page_numbers(d, "outside"), "")[1])),
        ("居右页码 (— 1 —)", lambda: run_on_docs("插入居右公文页码", lambda d, _: (proc.insert_page_numbers(d, "right"), "")[1])),
    ])

    b_symbol = make_large_btn(g4, "符号 ▾", "StarUnratedFull.png", lambda: None)

    def copy_sym(s: str):
        root.clipboard_clear()
        root.clipboard_append(s)
        log(f"已复制符号【{s}】到剪贴板，可直接按 Ctrl+V 粘贴。")

    attach_popup_menu(b_symbol, [
        ("将 [2026] 批量替换为 〔2026〕", lambda: run_on_docs("将[]替换为〔〕", lambda d, _: f"替换 {proc.replace_brackets_to_hex(d, 'square')} 处")),
        ("将 【2026】 批量替换为 〔2026〕", lambda: run_on_docs("将【】替换为〔〕", lambda d, _: f"替换 {proc.replace_brackets_to_hex(d, 'black')} 处")),
        "---",
        ("复制 〔〕 六角括号", lambda: copy_sym("〔〕")),
        ("复制 —— 破折号", lambda: copy_sym("——")),
        ("复制 …… 省略号", lambda: copy_sym("……")),
        ("复制 · 间隔符", lambda: copy_sym("·")),
        ("复制 《》 书名号(双)", lambda: copy_sym("《》")),
        ("复制 〈〉 书名号(单)", lambda: copy_sym("〈〉")),
        ("复制 ★ 密级五角星", lambda: copy_sym("★")),
        ("复制 □ 虚缺号", lambda: copy_sym("□")),
        ("复制 ①②③④⑤⑥⑦⑧⑨⑩", lambda: copy_sym("①②③④⑤⑥⑦⑧⑨⑩")),
    ])

    b_stamp = make_large_btn(g4, "公章 ▾", "StampTool.png", lambda: None)

    def insert_stamp_ui(every_page: bool):
        pic = filedialog.askopenfilename(title="请选择透明PNG电子公章图片", filetypes=[("公章图片", "*.png *.gif *.tiff"), ("所有文件", "*.*")])
        if not pic:
            return
        run_on_docs("每页批量加盖公章" if every_page else "插入电子公章", lambda d, _: (proc.insert_official_stamp_image(d, pic, every_page), "")[1])

    attach_popup_menu(b_stamp, [
        ("插入公章 (在文末落款加盖)", lambda: insert_stamp_ui(False)),
        ("每页批量盖章 (页眉全篇加盖)", lambda: insert_stamp_ui(True)),
    ])

    # 【第5组：其他】
    g6 = make_group(ribbon_bar, "其他")
    b_bak = make_large_btn(g6, "副本 ▾", "FileSaveAs.png", lambda: None)

    def export_copy(fmt: str):
        targets = ensure_target_files()
        for f in targets:
            try:
                p = Path(f)
                pdf_out = str(p.with_suffix(".pdf"))
                proc.convert_doc_to_pdf_kylin(str(p), pdf_out)
                if fmt == "pdf":
                    log(f"[成功] 已生成 PDF 副本: {pdf_out}")
                elif fmt == "scan":
                    scan_out = str(p.with_name(f"{p.stem}_扫描版.pdf"))
                    proc.script_make_scanned_pdf(pdf_out, scan_out)
                    log(f"[成功] 已生成仿真扫描版 PDF: {scan_out}")
                elif fmt in ("png", "jpg"):
                    prefix = str(p.parent / f"{p.stem}_page")
                    flag = "-jpeg" if fmt == "jpg" else "-png"
                    subprocess.run(["pdftoppm", flag, "-r", "200", pdf_out, prefix], check=True)
                    log(f"[成功] 已生成逐页 {fmt.upper()} 图片副本: {p.parent}")
            except Exception as e:
                log(f"[错误] 导出副本失败: {e}")

    attach_popup_menu(b_bak, [
        ("生成 PDF 副本 (.pdf)", lambda: export_copy("pdf")),
        ("生成 PNG 图片副本 (逐页高清图)", lambda: export_copy("png")),
        ("生成 JPG 图片副本 (逐页高清图)", lambda: export_copy("jpg")),
        ("生成仿真扫描版 PDF 副本...", lambda: export_copy("scan")),
    ])

    g6_stack = tk.Frame(g6, bg="#f8fafc")
    g6_stack.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    make_small_btn(g6_stack, "限制编辑", "ProtectDocument.png", lambda: messagebox.showinfo("提示", "可在 WPS 文字审阅/安全菜单或插件中开启只读限制编辑。")).pack(fill=tk.X, pady=1)

    b_pwd = make_small_btn(g6_stack, "文档加密 ▾", "PasswordProtectSection.png", lambda: None)
    b_pwd.pack(fill=tk.X, pady=1)
    attach_popup_menu(b_pwd, [
        ("移除 Excel (.xlsx) 写保护密码...", lambda: remove_xlsx_pwd_ui()),
    ])

    b_speak = make_small_btn(g6_stack, "朗读 ▾", "SpeakCells.png", lambda: None)
    b_speak.pack(fill=tk.X, pady=1)

    def speak_active_doc():
        targets = ensure_target_files()
        if not targets:
            return
        d = Document(targets[0])
        txt = "\n".join([(p.text or "").strip() for p in d.paragraphs if (p.text or "").strip()])[:600]
        if shutil.which("spd-say"):
            subprocess.Popen(["spd-say", "-l", "zh", txt])
            log("[语音朗读] 正在调用银河麒麟 speech-dispatcher 朗读公文...")
        elif shutil.which("espeak-ng"):
            subprocess.Popen(["espeak-ng", "-v", "cmn", txt])
            log("[语音朗读] 正在调用银河麒麟 espeak-ng 朗读公文...")
        else:
            messagebox.showinfo("语音模块提示", "如需启用系统级中文语音朗读，请在终端执行：\nsudo apt install speech-dispatcher espeak-ng")

    attach_popup_menu(b_speak, [
        ("开始朗读当前公文", speak_active_doc),
        ("停止语音朗读", lambda: subprocess.Popen(["pkill", "-f", "espeak"], stderr=subprocess.DEVNULL)),
    ])

    # 【第6组：关于】
    g5 = make_group(ribbon_bar, "关于")
    make_large_btn(g5, "关于插件", "Info.png", open_about_dialog)
    g5_stack = tk.Frame(g5, bg="#f8fafc")
    g5_stack.pack(side=tk.LEFT, fill=tk.Y, padx=2)
    make_small_btn(g5_stack, "访问网站", "GroupInsertLinks.png", lambda: subprocess.Popen(["xdg-open", "https://gw.xkonglong.com"], stderr=subprocess.DEVNULL)).pack(fill=tk.X, pady=1)
    make_small_btn(g5_stack, "在线帮助", "Help.png", open_about_dialog).pack(fill=tk.X, pady=1)
    make_small_btn(g5_stack, "建议反馈", "OfficeButtonFeedback.png", lambda: messagebox.showinfo("建议反馈", "作者邮箱：xkonglong@qq.com\n银河麒麟 V10 · 海光3350 专版 v2.2.0")).pack(fill=tk.X, pady=1)

    # 【第7组：选择 / 转换 / 脚本】
    g7 = make_group(ribbon_bar, "脚本与扩展")
    b_sel = make_large_btn(g7, "选择 ▾", "ArrowsMore.png", lambda: None)

    def open_doc_dir():
        targets = ensure_target_files()
        if not targets:
            return
        folder = str(Path(targets[0]).parent)
        if shutil.which("peony"):
            subprocess.Popen(["peony", targets[0]])
        elif shutil.which("xdg-open"):
            subprocess.Popen(["xdg-open", folder])

    attach_popup_menu(b_sel, [
        ("在麒麟文件管理器 (Peony) 中打开所在目录", open_doc_dir),
        ("提取文档内所有图片到文件夹", lambda: extract_doc_images_ui()),
    ])

    b_conv = make_large_btn(g7, "转换 ▾", "ContentControlPicture.png", lambda: None)
    attach_popup_menu(b_conv, [
        ("多张图片批量合成 PPTX 演示文稿...", lambda: images_to_pptx_ui()),
        ("多张图片批量合成 A4 PDF 文档...", lambda: images_to_pdf_ui()),
        ("移除 Excel (.xlsx) 工作表写保护...", lambda: remove_xlsx_pwd_ui()),
    ])

    b_vba = make_large_btn(g7, "脚本 ▾", "LogicBoldItalicAction.png", lambda: None)

    def prompt_rmb_ui():
        val = simpledialog.askstring("数字转人民币大写", "请输入要转换的数字金额：", initialvalue="128560.50", parent=root)
        if val:
            res = number_to_rmb_upper(val)
            root.clipboard_clear()
            root.clipboard_append(res)
            log(f"[人民币大写] {val} -> {res} (已复制到剪贴板)")
            messagebox.showinfo("人民币大写结果", f"转换结果：\n{res}\n\n(已自动复制到剪贴板)")

    def prompt_place_cards_ui():
        names = simpledialog.askstring("会议桌签标牌双面制作", "请输入参会人姓名（多个姓名请用逗号分隔）：", initialvalue="张  三,李四五,王  五", parent=root)
        if not names:
            return
        save_p = filedialog.asksaveasfilename(title="保存会议桌签标牌文档", defaultextension=".docx", initialfile="会议桌签标牌.docx")
        if save_p:
            proc.script_make_place_cards(names.split(","), save_p)
            log(f"[成功] 已生成 A4 横向双面对折会议桌签: {save_p}")
            if shutil.which("wps"):
                subprocess.Popen(["wps", save_p])

    attach_popup_menu(b_vba, [
        ("1. 公文自动排版 (2026国标版)", lambda: run_on_docs("2026国标公文自动排版", lambda d, _: str(proc.auto_layout_2026(d)))),
        ("2. 公文落款自动对齐 (署名+日期)", lambda: run_on_docs("公文落款自动对齐", lambda d, _: "已对齐" if proc.script_official_signature_align(d) else "段落不足")),
        ("3. 加粗第N章第N条 (精准词头)", lambda: run_on_docs("精准加粗第N章第N条", lambda d, _: f"共加粗 {proc.script_bold_chapters_articles(d, False)} 处")),
        ("4. 加粗第N章第N条 (整行加粗)", lambda: run_on_docs("整行加粗第N章第N条", lambda d, _: f"共加粗 {proc.script_bold_chapters_articles(d, True)} 处")),
        ("5. 第N章整行加粗居中", lambda: run_on_docs("第N章整行加粗居中", lambda d, _: f"共居中 {proc.script_center_chapter_lines(d)} 章")),
        ("6. 第N章/条后规范 1 空格", lambda: run_on_docs("第N章条后1空格", lambda d, _: f"共规范 {proc.script_chapter_article_one_space(d)} 处")),
        ("7. 重新整理章条序号 (第一~九十九章/条)", lambda: run_on_docs("重新整理章条序号", lambda d, _: f"共整理 {proc.script_renumber_chapters_articles(d)} 处")),
        ("8. （一）至（九十九）重新连续排序", lambda: run_on_docs("（一）~（九十九）重排", lambda d, _: f"共重排 {proc.script_renumber_cn_brackets(d)} 处")),
        ("9. 彻底删除全篇页眉页脚", lambda: run_on_docs("彻底删除页眉页脚", lambda d, _: f"已清理 {proc.script_delete_headers_footers(d)} 节")),
        "---",
        ("10. 数字转为人民币大写...", prompt_rmb_ui),
        ("11. 会议桌签标牌双面制作 (A4横向72号黑体)...", prompt_place_cards_ui),
        ("12. 提取文档内所有图片...", lambda: extract_doc_images_ui()),
        ("13. 图片批量转 PPTX 演示文稿...", lambda: images_to_pptx_ui()),
        ("14. 插入多张图片合成 PDF...", lambda: images_to_pdf_ui()),
        ("15. 移除 Excel (.xlsx) 写保护...", lambda: remove_xlsx_pwd_ui()),
    ])

    tk.Frame(root, bg="#cbd5e1", height=1).pack(fill=tk.X)

    # ==================== 紧凑型当前文档状态条（悬浮模式下同样可见） ====================
    status_strip = tk.Frame(root, bg="#eef2ff", pady=3, padx=10)
    status_strip.pack(fill=tk.X)

    lbl_active_doc = tk.Label(
        status_strip,
        text="当前未选择文档（点击任意上方【公文】功能区按钮时会自动提示选择 .docx 文件）",
        bg="#eef2ff",
        fg="#1e3a8a",
        font=("Noto Sans CJK SC", 9, "bold"),
    )
    lbl_active_doc.pack(side=tk.LEFT)

    tk.Checkbutton(status_strip, text="排版后自动在 WPS 中打开预览", variable=auto_open_wps, bg="#eef2ff", font=("Noto Sans CJK SC", 9)).pack(side=tk.RIGHT, padx=6)
    tk.Checkbutton(status_strip, text="直接覆盖原文件 (不加 _公文排版 后缀)", variable=overwrite_in_place, bg="#eef2ff", font=("Noto Sans CJK SC", 9)).pack(side=tk.RIGHT, padx=6)

    # ==================== 下方工作台区域（在悬浮插件栏模式下会自动收起） ====================
    work_area = tk.PanedWindow(root, orient=tk.HORIZONTAL, bg="#f3f4f6", sashwidth=6)
    work_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

    left_frame = tk.LabelFrame(work_area, text=" 1. 待处理公文文档列表 (.docx) ", font=("Noto Sans CJK SC", 10, "bold"), padx=8, pady=8)
    right_frame = tk.LabelFrame(work_area, text=" 2. 批量公文排版与运行日志 ", font=("Noto Sans CJK SC", 10, "bold"), padx=8, pady=8)
    work_area.add(left_frame, width=420)
    work_area.add(right_frame, width=800)

    file_listbox = tk.Listbox(left_frame, selectmode=tk.EXTENDED, font=("Noto Sans CJK SC", 10))
    file_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

    # 注意：严格使用整数字号 10，兼容 Python 3.8 Tk 8.6
    log_box = scrolledtext.ScrolledText(right_frame, height=12, font=("Monospace", 10), bg="#0f172a", fg="#e2e8f0")

    def log(msg: str):
        log_box.insert(tk.END, msg + "\n")
        log_box.see(tk.END)

    def add_files():
        files = filedialog.askopenfilenames(
            title="选择要排版的公文文档 (.docx)",
            filetypes=[("Word 公文文档", "*.docx"), ("所有文件", "*.*")],
        )
        for f in files:
            if f not in selected_files:
                selected_files.append(f)
                file_listbox.insert(tk.END, f)
        if selected_files:
            lbl_active_doc.config(text=f"当前已选 {len(selected_files)} 个文档（首个: {Path(selected_files[0]).name}）")
        log(f"已添加 {len(files)} 个文档，当前共 {len(selected_files)} 个待处理文档。")

    def clear_files():
        selected_files.clear()
        file_listbox.delete(0, tk.END)
        lbl_active_doc.config(text="当前未选择文档")
        log("已清空文档列表。")

    def open_selected_in_wps():
        targets = ensure_target_files()
        if targets and shutil.which("wps"):
            subprocess.Popen(["wps", targets[0]], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    btn_bar = tk.Frame(left_frame)
    btn_bar.pack(fill=tk.X)
    tk.Button(btn_bar, text="添加公文文档...", bg="#b91c1c", fg="white", font=("Noto Sans CJK SC", 10, "bold"), command=add_files).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_bar, text="在 WPS 中打开", font=("Noto Sans CJK SC", 10), command=open_selected_in_wps).pack(side=tk.LEFT, padx=3)
    tk.Button(btn_bar, text="清空列表", font=("Noto Sans CJK SC", 10), command=clear_files).pack(side=tk.LEFT, padx=3)

    # 附加工具函数
    def extract_doc_images_ui():
        targets = ensure_target_files()
        for f in targets:
            imgs = proc.script_extract_all_images(f)
            log(f"[提取图片完成] 从 {Path(f).name} 提取了 {len(imgs)} 张图片。")
        if targets:
            messagebox.showinfo("提取完成", "已提取文档内所有图片到文档同目录子文件夹中！")

    def images_to_pptx_ui():
        imgs = filedialog.askopenfilenames(title="选择要转为 PPTX 的图片", filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp")])
        if not imgs:
            return
        out_p = filedialog.asksaveasfilename(title="保存 PPTX 文件", defaultextension=".pptx", initialfile="图片汇报演示.pptx")
        if out_p:
            proc.script_images_to_pptx(list(imgs), out_p)
            log(f"[成功] 已将 {len(imgs)} 张图片合成为 PPTX: {out_p}")

    def images_to_pdf_ui():
        imgs = filedialog.askopenfilenames(title="选择要合成 PDF 的图片", filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp")])
        if not imgs:
            return
        out_p = filedialog.asksaveasfilename(title="保存 PDF 文件", defaultextension=".pdf", initialfile="图片合并文档.pdf")
        if out_p:
            proc.script_images_to_pdf(list(imgs), out_p)
            log(f"[成功] 已将 {len(imgs)} 张图片合成为 PDF: {out_p}")

    def remove_xlsx_pwd_ui():
        xls = filedialog.askopenfilename(title="选择要移除写保护的 .xlsx 文件", filetypes=[("Excel 工作簿", "*.xlsx")])
        if not xls:
            return
        out_p = str(Path(xls).with_name(f"{Path(xls).stem}_已解除保护.xlsx"))
        proc.script_remove_xlsx_protection(xls, out_p)
        log(f"[成功] 已解除 Excel 工作表写保护: {out_p}")
        messagebox.showinfo("完成", f"已解除写保护并保存为：\n{out_p}")

    # 右侧快捷批处理网格（采用自适应 Grid 权重，彻底杜绝按钮右侧截断）
    action_grid = tk.Frame(right_frame, pady=4)
    action_grid.pack(fill=tk.X)
    action_grid.columnconfigure(0, weight=1)
    action_grid.columnconfigure(1, weight=1)
    action_grid.columnconfigure(2, weight=1)

    quick_actions = [
        ("★ 一键国标公文自动排版 (2026版)", lambda: run_on_docs("2026国标自动排版", lambda d, _: str(proc.auto_layout_2026(d))), "#b91c1c", "white"),
        ("公文落款对齐 (署名+日期)", lambda: run_on_docs("公文落款对齐", lambda d, _: "完成" if proc.script_official_signature_align(d) else "段落不足"), "#ffffff", "#1f2328"),
        ("重新整理章条序号 (一~九十九)", lambda: run_on_docs("章条序号重排", lambda d, _: str(proc.script_renumber_chapters_articles(d))), "#ffffff", "#1f2328"),
        ("精准加粗第N章第N条", lambda: run_on_docs("精准加粗第N章第N条", lambda d, _: f"{proc.script_bold_chapters_articles(d, False)}处"), "#ffffff", "#1f2328"),
        ("第N章整行加粗居中", lambda: run_on_docs("第N章整行加粗居中", lambda d, _: f"{proc.script_center_chapter_lines(d)}章"), "#ffffff", "#1f2328"),
        ("（一）至（九十九）重新排序", lambda: run_on_docs("（一）~（九十九）重排", lambda d, _: f"{proc.script_renumber_cn_brackets(d)}处"), "#ffffff", "#1f2328"),
        ("插入单行红头版头+反线...", prompt_red_head_single, "#ffffff", "#1f2328"),
        ("插入国标双行版记表格", lambda: run_on_docs("插入双行版记", lambda d, _: (proc.insert_banji_table(d, True), "")[1]), "#ffffff", "#1f2328"),
        ("生成仿真扫描版 PDF 副本", lambda: export_copy("scan"), "#ffffff", "#1f2328"),
    ]
    for idx, (label, callback, bg_c, fg_c) in enumerate(quick_actions):
        r, c = divmod(idx, 3)
        tk.Button(
            action_grid,
            text=label,
            bg=bg_c,
            fg=fg_c,
            font=("Noto Sans CJK SC", 10, "bold" if bg_c == "#b91c1c" else "normal"),
            pady=5,
            command=callback,
        ).grid(row=r, column=c, padx=4, pady=4, sticky="ew")

    log_box.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
    log(get_kylin_hygon_system_info())
    log("提示：上方已为您提供 1:1 复刻 Windows 版的【公文】功能区工具栏！您可点击右上角【📌 切换为 WPS 顶部悬浮插件栏】吸附在 WPS 顶部使用，或点击【🔧 一键激活/修复 WPS 内置插件】将其注入 WPS 内部。")

    root.mainloop()


def main():
    parser = argparse.ArgumentParser(
        description="小恐龙公文排版助手 v2.2.0 (银河麒麟 V10 · 海光 Hygon C86 3350 x86_64 专版)"
    )
    parser.add_argument("inputs", nargs="*", help="待处理的 .docx 文档路径")
    parser.add_argument("-o", "--output", help="输出文件路径")
    parser.add_argument(
        "-a",
        "--action",
        choices=["auto2026", "page-setup", "renumber-chapters", "bold-chapters", "center-chapters", "clean-hf"],
        default="auto2026",
        help="执行的排版动作 (默认: auto2026)",
    )
    parser.add_argument("--red-header", help="插入单行红头版头（指定发文机关名称，如 '○○市人民政府文件'）")
    parser.add_argument("--doc-number", help="指定发文字号（如 '○政发〔2026〕12号'）")
    parser.add_argument("--banji", choices=["single", "double"], help="在文末插入单行或双行版记")
    parser.add_argument("--cc-org", default="各县（市、区）人民政府，市直各部门。", help="双行版记的抄送机关")
    parser.add_argument("--print-org", default="○○○○○○办公室", help="版记的印发机关")
    parser.add_argument("--export-pdf", action="store_true", help="同时生成同名 PDF 副本")
    parser.add_argument("--scanned-pdf", action="store_true", help="同时生成仿真扫描版 PDF 副本")
    parser.add_argument("--place-cards", help="生成A4横向双面会议桌签（多个姓名用逗号分隔）")
    parser.add_argument("--rmb", help="将指定数字金额转换为人民币大写")
    parser.add_argument("--fix-wps", action="store_true", help="深度激活/修复 WPS Office Linux 内置【公文】功能区插件")
    parser.add_argument("--info", action="store_true", help="输出银河麒麟与海光 3350 环境诊断信息")
    parser.add_argument("--serve-bridge", action="store_true", help="启动 WPS Linux 加载项本地桥接守护进程")
    parser.add_argument("--port", type=int, default=38890, help="桥接服务端口 (默认 38890)")
    parser.add_argument("--gui", action="store_true", help="启动银河麒麟桌面图形界面工作台")

    args = parser.parse_args()
    if args.gui or (len(sys.argv) == 1 and os.environ.get("DISPLAY")):
        launch_kylin_gui()
        return
    sys.exit(run_cli(args))


if __name__ == "__main__":
    main()
