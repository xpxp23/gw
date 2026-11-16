#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小恐龙公文排版助手 v2.2.0 — 银河麒麟 V10 SP1 WPS Office Linux 加载项深度激活与诊断模块
解决政企版/专业版/社区版 WPS Linux 默认关闭 JsApiPlugin 或因路径差异不显示【公文】功能区选项卡的问题
"""

import os
import shutil
import socket
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Tuple


ADDIN_DIR = Path(__file__).resolve().parent.parent.parent / "wps-js-addin"
BRIDGE_PORT = 38890


def is_bridge_running(port: int = BRIDGE_PORT) -> bool:
    """检查 127.0.0.1:38890 本地桥接服务是否已启动"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", port)) == 0


def ensure_bridge_started() -> bool:
    """若本地桥接服务未运行，则在后台拉起守护进程"""
    if is_bridge_running():
        return True
    try:
        app_root = Path(__file__).resolve().parent.parent
        env = {**os.environ, "PYTHONPATH": f"{app_root}:{os.environ.get('PYTHONPATH', '')}"}
        subprocess.Popen(
            [sys.executable, "-m", "xkl_gw.bridge_server"],
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
        return True
    except Exception:
        return False


def build_plugin_xml_content(local_install_dir: str) -> str:
    """生成同时兼容 WPS 2019 / 2023 / 信创版 Linux 的 publish.xml 与 jsplugins.xml 内容"""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
    <jspluginonline name="xkl_gw_addin" type="wps" url="http://127.0.0.1:{BRIDGE_PORT}/" version="2.2.0" enable="enable_dev" install="xkl_gw_addin_2.2.0"/>
    <jsplugin name="xkl_gw_addin" type="wps" url="http://127.0.0.1:{BRIDGE_PORT}/" version="2.2.0" enable="enable_dev"/>
</jsplugins>
"""


def activate_wps_linux_plugin(use_pkexec_if_needed: bool = False) -> Tuple[bool, List[str]]:
    """
    全方位注册并激活 WPS Office Linux 的【公文】功能区加载项：
      1. 确保 127.0.0.1:38890 桥接与静态资源服务常驻运行
      2. 向 ~/.local/share/Kingsoft/wps/jsaddons 与 ~/.kingsoft/wps/jsaddons 双路径部署插件目录及 publish.xml / jsplugins.xml
      3. 配置 ~/.config/Kingsoft/Office.conf 开启 JS 加载项支持
      4. 检查并配置 /opt/kingsoft/wps-office/office6/cfgs/oem.ini 中的 JsApiPlugin=true 与 JSPluginsServer
    """
    logs: List[str] = []

    if not ADDIN_DIR.exists():
        return False, [f"[错误] 未找到加载项源目录: {ADDIN_DIR}"]

    # 1. 启动后台桥接服务
    ensure_bridge_started()
    logs.append(f"[1/4] 本地桥接与插件托管服务 (127.0.0.1:{BRIDGE_PORT}): {'运行中' if is_bridge_running() else '已触发后台启动'}")

    # 2. 部署用户级双路径 jsaddons
    home = Path.home()
    user_addon_roots = [
        home / ".local" / "share" / "Kingsoft" / "wps" / "jsaddons",
        home / ".kingsoft" / "wps" / "jsaddons",
    ]
    for root in user_addon_roots:
        try:
            root.mkdir(parents=True, exist_ok=True)
            for sub_name in ["xkl_gw_addin_2.2.0", "xkl_gw_addin_2.1.0", "xkl_gw_addin"]:
                dest = root / sub_name
                if dest.exists():
                    shutil.rmtree(dest)
                shutil.copytree(str(ADDIN_DIR), str(dest))

            xml_text = build_plugin_xml_content(str(root / "xkl_gw_addin_2.2.0"))
            (root / "publish.xml").write_text(xml_text, encoding="utf-8")
            (root / "jsplugins.xml").write_text(xml_text, encoding="utf-8")
            (root / "auth.xml").write_text(
                '<?xml version="1.0" encoding="UTF-8"?>\n<jsplugins>\n    <jsplugin name="xkl_gw_addin" enable="enable"/>\n</jsplugins>\n',
                encoding="utf-8",
            )
            logs.append(f"[2/4] 已写入用户级 WPS 插件目录: {root}")
        except Exception as e:
            logs.append(f"[警告] 写入 {root} 失败: {e}")

    # 3. 更新 ~/.config/Kingsoft/Office.conf
    office_conf = home / ".config" / "Kingsoft" / "Office.conf"
    try:
        office_conf.parent.mkdir(parents=True, exist_ok=True)
        conf_text = office_conf.read_text(encoding="utf-8", errors="ignore") if office_conf.exists() else ""
        if "jsaddons" not in conf_text:
            with open(office_conf, "a", encoding="utf-8") as f:
                f.write(f"\n[jsaddons]\nEnableJsAddons=true\nJSPluginsServer=http://127.0.0.1:{BRIDGE_PORT}/jsplugins.xml\n")
        logs.append(f"[3/4] 已更新用户配置文件: {office_conf}")
    except Exception as e:
        logs.append(f"[警告] 更新 Office.conf 提示: {e}")

    # 4. 检查并配置 /opt/kingsoft/wps-office/office6/cfgs/oem.ini
    oem_candidates = [
        Path("/opt/kingsoft/wps-office/office6/cfgs/oem.ini"),
        Path("/usr/lib/office6/cfgs/oem.ini"),
    ]
    oem_patched = False
    for oem_path in oem_candidates:
        if not oem_path.exists():
            continue
        try:
            raw = oem_path.read_text(encoding="utf-8", errors="ignore")
            needs_patch = ("JsApiPlugin=true" not in raw) or ("JSPluginsServer=" not in raw)
            if not needs_patch:
                logs.append(f"[4/4] 系统级 WPS 配置 {oem_path} 已开启 JsApiPlugin=true")
                oem_patched = True
                break

            # 构造更新后的 oem.ini 内容
            lines = raw.splitlines()
            out_lines = []
            in_support = False
            support_found = False
            js_added = False
            for line in lines:
                stripped = line.strip()
                if stripped.startswith("[") and stripped.endswith("]"):
                    if in_support and not js_added:
                        out_lines.append("JsApiPlugin=true")
                        out_lines.append("JsApiShowWebDebugger=true")
                        out_lines.append(f"JSPluginsServer=http://127.0.0.1:{BRIDGE_PORT}/jsplugins.xml")
                        js_added = True
                    in_support = stripped.lower() == "[support]"
                    if in_support:
                        support_found = True
                if in_support and stripped.startswith(("JsApiPlugin=", "JSPluginsServer=")):
                    continue
                out_lines.append(line)

            if in_support and not js_added:
                out_lines.append("JsApiPlugin=true")
                out_lines.append("JsApiShowWebDebugger=true")
                out_lines.append(f"JSPluginsServer=http://127.0.0.1:{BRIDGE_PORT}/jsplugins.xml")
            elif not support_found:
                out_lines.extend([
                    "",
                    "[Support]",
                    "JsApiPlugin=true",
                    "JsApiShowWebDebugger=true",
                    f"JSPluginsServer=http://127.0.0.1:{BRIDGE_PORT}/jsplugins.xml",
                ])

            new_oem_text = "\n".join(out_lines) + "\n"

            if os.access(oem_path, os.W_OK):
                oem_path.write_text(new_oem_text, encoding="utf-8")
                logs.append(f"[4/4] 已直接写入系统级配置: {oem_path}")
                oem_patched = True
            elif use_pkexec_if_needed and shutil.which("pkexec"):
                tmp_oem = Path("/tmp/xkl_wps_oem.ini")
                tmp_oem.write_text(new_oem_text, encoding="utf-8")
                res = subprocess.run(
                    ["pkexec", "cp", "-f", str(tmp_oem), str(oem_path)],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                )
                if res.returncode == 0:
                    logs.append(f"[4/4] 已通过麒麟授权写入系统级配置: {oem_path}")
                    oem_patched = True
                else:
                    logs.append(
                        f"[提示] {oem_path} 需要管理员权限开启 JsApiPlugin=true。\n"
                        f"      请在终端执行: sudo bash {ADDIN_DIR.parent / 'scripts' / 'enable_wps_linux_plugin.sh'}"
                    )
            else:
                logs.append(
                    f"[提示] 检测到 {oem_path} 尚未开启 JsApiPlugin=true（需 sudo 权限写入）。\n"
                    f"      请执行: sudo bash {ADDIN_DIR.parent / 'scripts' / 'enable_wps_linux_plugin.sh'}"
                )
        except Exception as e:
            logs.append(f"[警告] 检查 {oem_path} 异常: {e}")

    return True, logs
