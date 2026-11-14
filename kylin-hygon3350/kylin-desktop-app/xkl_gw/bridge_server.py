#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小恐龙公文排版助手 - 银河麒麟本地桥接服务与 WPS Linux 加载项托管守护进程
监听地址: http://127.0.0.1:38890
功能:
  1. 托管 wps-js-addin 静态资源 (供 WPS Linux 加载项在线/本地加载)
  2. 提供银河麒麟系统级能力接口:
     - /api/system-info      检测银河麒麟版本与海光 Hygon C86 3350 CPU 信息
     - /api/tts-speak        调用 speech-dispatcher (spd-say) 或 espeak-ng 朗读中文公文
     - /api/tts-stop         停止语音朗读
     - /api/pdf-to-images    调用 pdftoppm / Pillow 将 PDF 逐页转为高清 PNG/JPG
     - /api/make-scanned-pdf 生成仿真扫描版 PDF
     - /api/open-folder      调用麒麟 Peony 文件管理器或 xdg-open 打开文档目录
     - /api/pick-file        调用 zenity / kdialog / tkinter 打开原生文件选择框
     - /api/copy-clipboard   复制文本到系统剪贴板 (xclip / wl-copy / tkinter)
     - /api/launch-desktop-gui 启动独立桌面端公文批处理工作台
"""

import glob
import http.server
import json
import os
import platform
import shutil
import socketserver
import subprocess
import sys
import threading
from pathlib import Path
from urllib.parse import urlparse

from .docx_engine import GwDocumentProcessor


PORT = 38890
ADDIN_ROOT = Path(__file__).resolve().parent.parent.parent / "wps-js-addin"


def get_kylin_hygon_system_info() -> str:
    """获取银河麒麟操作系统版本与海光 3350 CPU 硬件信息"""
    os_name = "银河麒麟桌面操作系统 (Kylin OS)"
    for release_file in ["/etc/kylin-release", "/etc/os-release"]:
        if os.path.isfile(release_file):
            try:
                content = Path(release_file).read_text(encoding="utf-8", errors="ignore")
                for line in content.splitlines():
                    if line.startswith("PRETTY_NAME="):
                        os_name = line.split("=", 1)[1].strip().strip('"')
                        break
                    elif "Kylin" in line and "=" not in line:
                        os_name = line.strip()
                        break
            except Exception:
                pass

    cpu_model = platform.processor() or platform.machine()
    cpu_cores = os.cpu_count() or 8
    if os.path.isfile("/proc/cpuinfo"):
        try:
            for line in Path("/proc/cpuinfo").read_text(encoding="utf-8", errors="ignore").splitlines():
                if line.lower().startswith("model name"):
                    cpu_model = line.split(":", 1)[1].strip()
                    break
        except Exception:
            pass

    wps_path = shutil.which("wps") or "/opt/kingsoft/wps-office/office6/wps"
    wps_status = f"已安装 ({wps_path})" if os.path.exists(wps_path) else "未检测到默认路径 (可使用独立桌面端模式)"

    return (
        f"操作系统: {os_name}\n"
        f"内核架构: {platform.system()} {platform.release()} ({platform.machine()})\n"
        f"处理器  : {cpu_model} ({cpu_cores} 逻辑核心)\n"
        f"WPS状态 : {wps_status}\n"
        f"Python  : {platform.python_version()} | 桥接服务端口: 127.0.0.1:{PORT}"
    )


class GwBridgeHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ADDIN_ROOT), **kwargs)

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def _send_json(self, data: dict, status: int = 200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length).decode("utf-8", errors="ignore") if length > 0 else "{}"
        try:
            payload = json.loads(raw) if raw else {}
        except Exception:
            payload = {}

        route = parsed.path
        try:
            if route == "/api/system-info":
                self._send_json({"ok": True, "info": get_kylin_hygon_system_info()})
                return

            if route == "/api/tts-speak":
                text = payload.get("text", "")
                rate = int(payload.get("rate", 175))
                if text:
                    if shutil.which("spd-say"):
                        subprocess.Popen(["spd-say", "-l", "zh", "-r", str(max(-100, min(100, rate - 175))), text])
                    elif shutil.which("espeak-ng"):
                        subprocess.Popen(["espeak-ng", "-v", "cmn", "-s", str(rate), text])
                    elif shutil.which("espeak"):
                        subprocess.Popen(["espeak", "-v", "zh", "-s", str(rate), text])
                self._send_json({"ok": True})
                return

            if route == "/api/tts-stop":
                if shutil.which("spd-say"):
                    subprocess.Popen(["spd-say", "-S"])
                subprocess.Popen(["pkill", "-f", "espeak"], stderr=subprocess.DEVNULL)
                self._send_json({"ok": True})
                return

            if route == "/api/open-folder":
                target_path = payload.get("path", "")
                folder = os.path.dirname(target_path) if os.path.isfile(target_path) else target_path
                if shutil.which("peony"):
                    subprocess.Popen(["peony", target_path if os.path.exists(target_path) else folder])
                elif shutil.which("xdg-open"):
                    subprocess.Popen(["xdg-open", folder])
                self._send_json({"ok": True, "folder": folder})
                return

            if route == "/api/pick-file":
                title = payload.get("title", "选择文件")
                selected = ""
                if shutil.which("zenity"):
                    res = subprocess.run(
                        ["zenity", "--file-selection", f"--title={title}"],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                    )
                    if res.returncode == 0:
                        selected = res.stdout.strip()
                self._send_json({"ok": bool(selected), "path": selected})
                return

            if route == "/api/copy-clipboard":
                txt = payload.get("text", "")
                if shutil.which("xclip"):
                    p = subprocess.Popen(["xclip", "-selection", "clipboard"], stdin=subprocess.PIPE)
                    p.communicate(input=txt.encode("utf-8"))
                elif shutil.which("wl-copy"):
                    p = subprocess.Popen(["wl-copy"], stdin=subprocess.PIPE)
                    p.communicate(input=txt.encode("utf-8"))
                self._send_json({"ok": True})
                return

            if route == "/api/pdf-to-images":
                pdf_path = payload.get("pdfPath", "")
                out_dir = payload.get("outputDir", os.path.dirname(pdf_path))
                base_name = payload.get("baseName", "page")
                fmt = payload.get("format", "png").lower()
                dpi = int(payload.get("dpi", 200))
                remove_tmp = bool(payload.get("removeTempPdf", False))

                prefix = os.path.join(out_dir, f"{base_name}_page")
                flag = "-jpeg" if fmt in ("jpg", "jpeg") else "-png"
                subprocess.run(["pdftoppm", flag, "-r", str(dpi), pdf_path, prefix], check=True)
                ext = "jpg" if fmt in ("jpg", "jpeg") else "png"
                files = sorted(glob.glob(f"{prefix}*.{ext}"))
                if remove_tmp and os.path.isfile(pdf_path):
                    try:
                        os.remove(pdf_path)
                    except Exception:
                        pass
                self._send_json({"ok": True, "count": len(files), "files": files})
                return

            if route == "/api/make-scanned-pdf":
                pdf_path = payload.get("pdfPath", "")
                out_path = payload.get("outputPath", pdf_path.replace(".pdf", "_scanned.pdf"))
                skew = float(payload.get("skew", 0.6))
                noise = float(payload.get("noise", 0.06))
                dpi = int(payload.get("dpi", 200))
                remove_tmp = bool(payload.get("removeTempPdf", False))

                proc = GwDocumentProcessor()
                proc.script_make_scanned_pdf(pdf_path, out_path, skew=skew, noise=noise, dpi=dpi)
                if remove_tmp and os.path.isfile(pdf_path):
                    try:
                        os.remove(pdf_path)
                    except Exception:
                        pass
                self._send_json({"ok": True, "outputPath": out_path})
                return

            if route == "/api/launch-desktop-gui":
                subprocess.Popen([sys.executable, "-m", "xkl_gw.gui_app"])
                self._send_json({"ok": True})
                return

            self._send_json({"ok": False, "error": f"Unknown endpoint: {route}"}, status=404)
        except Exception as e:
            self._send_json({"ok": False, "error": str(e)}, status=500)


def run_bridge_server(port: int = PORT):
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("127.0.0.1", port), GwBridgeHandler) as httpd:
        print(f"[小恐龙公文排版助手-麒麟桥接服务] 已启动: http://127.0.0.1:{port}/")
        httpd.serve_forever()


if __name__ == "__main__":
    run_bridge_server()
