#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小恐龙公文排版助手 (银河麒麟 V10 / 海光 Hygon C86 3350 x86_64 专版)
独立桌面端 GUI 工作台与 CLI 命令行批处理工具
支持在不依赖 Windows VSTO 的情况下，对单个或多个 .docx / .xlsx / 图片进行国标公文排版、红头版记生成与 42 项脚本处理
"""

import argparse
import os
import sys
import threading
from pathlib import Path
from typing import List

import docx
from docx import Document

from .bridge_server import get_kylin_hygon_system_info, run_bridge_server
from .docx_engine import GwDocumentProcessor, number_to_rmb_upper


def run_cli(args: argparse.Namespace) -> int:
    proc = GwDocumentProcessor()

    if args.info:
        print(get_kylin_hygon_system_info())
        return 0

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
            cnt = proc.script_delete_headers_footers(doc)
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
    """启动银河麒麟桌面端可视化公文排版工作台 (Tkinter/TTK)"""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, ttk
    except ImportError:
        print("当前环境未安装 python3-tk，已自动切换为命令行帮助模式：")
        print(get_kylin_hygon_system_info())
        return

    root = tk.Tk()
    root.title("小恐龙公文排版助手 v2.1.0 — 银河麒麟 V10 · 海光 3350 (x86_64) 信创专版")
    root.geometry("920x640")
    root.configure(bg="#f5f7fa")

    proc = GwDocumentProcessor()
    selected_files: List[str] = []

    # 顶部标题栏
    header = tk.Frame(root, bg="#b91c1c", height=64)
    header.pack(fill=tk.X)
    tk.Label(
        header,
        text="小恐龙公文排版助手 (银河麒麟 V10 · 海光 Hygon C86 3350 专版)",
        bg="#b91c1c",
        fg="white",
        font=("Noto Sans CJK SC", 15, "bold"),
        pady=12,
    ).pack(side=tk.LEFT, padx=18)

    # 主分区
    main_pane = tk.PanedWindow(root, orient=tk.HORIZONTAL, bg="#f5f7fa", sashwidth=6)
    main_pane.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

    left_frame = tk.LabelFrame(main_pane, text=" 1. 待处理公文文档列表 (.docx) ", font=("Noto Sans CJK SC", 11, "bold"), padx=10, pady=10)
    right_frame = tk.LabelFrame(main_pane, text=" 2. 国标公文排版与脚本批处理操作 ", font=("Noto Sans CJK SC", 11, "bold"), padx=10, pady=10)
    main_pane.add(left_frame, width=380)
    main_pane.add(right_frame, width=500)

    file_listbox = tk.Listbox(left_frame, selectmode=tk.EXTENDED, font=("Noto Sans CJK SC", 10))
    file_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

    log_box = scrolledtext.ScrolledText(right_frame, height=10, font=("Monospace", 9.5), bg="#0f172a", fg="#e2e8f0")

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
        log(f"已添加 {len(files)} 个文档，当前共 {len(selected_files)} 个待处理文档。")

    def clear_files():
        selected_files.clear()
        file_listbox.delete(0, tk.END)
        log("已清空文档列表。")

    btn_bar = tk.Frame(left_frame)
    btn_bar.pack(fill=tk.X)
    tk.Button(btn_bar, text="添加公文文档...", bg="#b91c1c", fg="white", command=add_files).pack(side=tk.LEFT, padx=4)
    tk.Button(btn_bar, text="清空列表", command=clear_files).pack(side=tk.LEFT, padx=4)

    # 右侧选项
    var_red_head = tk.BooleanVar(value=False)
    var_agency = tk.StringVar(value="○○○○○○委员会文件")
    var_banji = tk.StringVar(value="none")
    var_pdf = tk.BooleanVar(value=False)
    var_scan = tk.BooleanVar(value=False)

    opt_frame = tk.Frame(right_frame)
    opt_frame.pack(fill=tk.X, pady=4)
    tk.Checkbutton(opt_frame, text="自动插入单行红头版头，发文机关名称：", variable=var_red_head).grid(row=0, column=0, sticky="w")
    tk.Entry(opt_frame, textvariable=var_agency, width=24).grid(row=0, column=1, sticky="w", padx=4)

    tk.Label(opt_frame, text="版记选项：").grid(row=1, column=0, sticky="w", pady=4)
    banji_sub = tk.Frame(opt_frame)
    banji_sub.grid(row=1, column=1, sticky="w")
    tk.Radiobutton(banji_sub, text="不加版记", variable=var_banji, value="none").pack(side=tk.LEFT)
    tk.Radiobutton(banji_sub, text="单行版记", variable=var_banji, value="single").pack(side=tk.LEFT)
    tk.Radiobutton(banji_sub, text="双行版记", variable=var_banji, value="double").pack(side=tk.LEFT)

    tk.Checkbutton(opt_frame, text="同时生成同名 PDF 副本", variable=var_pdf).grid(row=2, column=0, sticky="w")
    tk.Checkbutton(opt_frame, text="同时生成仿真扫描版 PDF", variable=var_scan).grid(row=2, column=1, sticky="w")

    def execute_batch(mode: str):
        if not selected_files:
            messagebox.showwarning("提示", "请先在左侧点击【添加公文文档】选择要处理的 .docx 文件！")
            return
        for f in selected_files:
            try:
                p = Path(f)
                out_p = p.with_name(f"{p.stem}_麒麟公文排版.docx")
                doc = Document(str(p))
                if var_red_head.get():
                    proc.insert_red_header_single(doc, agency_name=var_agency.get())
                if mode == "auto2026":
                    st = proc.auto_layout_2026(doc)
                    log(f"[成功] 2026国标自动排版: {p.name} -> {out_p.name} ({st})")
                elif mode == "renumber_ch":
                    ch, art = proc.script_renumber_chapters_articles(doc)
                    log(f"[成功] 章条序号重排 ({ch}章, {art}条): {out_p.name}")
                elif mode == "bold_ch":
                    cnt = proc.script_bold_chapters_articles(doc, whole_line=False)
                    log(f"[成功] 精准加粗第N章第N条 ({cnt}处): {out_p.name}")
                elif mode == "center_ch":
                    cnt = proc.script_center_chapter_lines(doc)
                    log(f"[成功] 第N章整行加粗居中 ({cnt}处): {out_p.name}")
                elif mode == "renumber_cn":
                    cnt = proc.script_renumber_cn_brackets(doc)
                    log(f"[成功] （一）至（九十九）重新编号 ({cnt}处): {out_p.name}")
                elif mode == "clean_hf":
                    proc.script_delete_headers_footers(doc)
                    log(f"[成功] 清除页眉页脚: {out_p.name}")

                if var_banji.get() in ("single", "double"):
                    proc.insert_banji_table(doc, double_line=(var_banji.get() == "double"))

                doc.save(str(out_p))

                if var_pdf.get() or var_scan.get():
                    pdf_file = str(out_p.with_suffix(".pdf"))
                    proc.convert_doc_to_pdf_kylin(str(out_p), pdf_file)
                    log(f"  -> 已导出 PDF: {Path(pdf_file).name}")
                    if var_scan.get():
                        scan_file = str(out_p.with_name(f"{out_p.stem}_扫描版.pdf"))
                        proc.script_make_scanned_pdf(pdf_file, scan_file)
                        log(f"  -> 已生成仿真扫描件: {Path(scan_file).name}")
            except Exception as ex:
                log(f"[错误] 处理 {f} 失败: {ex}")
        messagebox.showinfo("处理完成", "批量公文排版处理已完成！")

    action_grid = tk.Frame(right_frame, pady=8)
    action_grid.pack(fill=tk.X)
    actions = [
        ("★ 一键国标公文自动排版 (2026版)", "auto2026", "#b91c1c", "white"),
        ("重新整理章条序号 (第一~九十九章/条)", "renumber_ch", "#ffffff", "#1f2328"),
        ("精准加粗第N章第N条", "bold_ch", "#ffffff", "#1f2328"),
        ("第N章整行加粗居中", "center_ch", "#ffffff", "#1f2328"),
        ("（一）至（九十九）重新排序", "renumber_cn", "#ffffff", "#1f2328"),
        ("彻底删除全篇页眉页脚", "clean_hf", "#ffffff", "#1f2328"),
    ]
    for idx, (label, cmd_key, bg_c, fg_c) in enumerate(actions):
        r, c = divmod(idx, 2)
        tk.Button(
            action_grid,
            text=label,
            bg=bg_c,
            fg=fg_c,
            width=28,
            pady=5,
            command=lambda k=cmd_key: execute_batch(k),
        ).grid(row=r, column=c, padx=4, pady=4, sticky="ew")

    log_box.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
    log(get_kylin_hygon_system_info())
    log("提示：选择左侧 .docx 文档后点击上方功能按钮即可批量排版。")

    # 后台启动 WPS Linux 桥接服务
    t = threading.Thread(target=run_bridge_server, daemon=True)
    t.start()

    root.mainloop()


def main():
    parser = argparse.ArgumentParser(
        description="小恐龙公文排版助手 (银河麒麟 V10 · 海光 Hygon C86 3350 x86_64 专版)"
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
