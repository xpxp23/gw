#!/usr/bin/env bash
# ==============================================================================
# 小恐龙公文排版助手 v2.2.0 — WPS Office Linux 内置【公文】功能区插件深度激活脚本
# 解决银河麒麟 V10 SP1 下 WPS 默认关闭 JsApiPlugin 或未读取 jsaddons 的问题
# ==============================================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
ADDIN_SRC="${PROJECT_ROOT}/wps-js-addin"
BRIDGE_PORT=38890

echo "=================================================================="
echo "  正在激活 WPS Office Linux 内置【公文】功能区插件 (v2.2.0) ..."
echo "=================================================================="

# 1. 部署到当前用户与真实登录用户的双路径 (~/.local/share/Kingsoft/wps/jsaddons 与 ~/.kingsoft/wps/jsaddons)
TARGET_USER="${SUDO_USER:-$USER}"
TARGET_HOME="$(eval echo "~${TARGET_USER}")"

for addon_root in "${TARGET_HOME}/.local/share/Kingsoft/wps/jsaddons" "${TARGET_HOME}/.kingsoft/wps/jsaddons"; do
    mkdir -p "${addon_root}"
    for folder_name in "xkl_gw_addin_2.2.0" "xkl_gw_addin_2.1.0" "xkl_gw_addin"; do
        rm -rf "${addon_root}/${folder_name}"
        cp -rf "${ADDIN_SRC}" "${addon_root}/${folder_name}"
    done
    cp -f "${ADDIN_SRC}/publish.xml" "${addon_root}/publish.xml"
    cp -f "${ADDIN_SRC}/jsplugins.xml" "${addon_root}/jsplugins.xml"
    cat << 'EOF' > "${addon_root}/auth.xml"
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
    <jsplugin name="xkl_gw_addin" enable="enable"/>
</jsplugins>
EOF
    chown -R "${TARGET_USER}" "${addon_root}" 2>/dev/null || true
    echo "[✓] 已部署用户级 WPS 加载项目录: ${addon_root}"
done

# 2. 更新 ~/.config/Kingsoft/Office.conf
OFFICE_CONF="${TARGET_HOME}/.config/Kingsoft/Office.conf"
mkdir -p "$(dirname "${OFFICE_CONF}")"
if ! grep -q "JSPluginsServer" "${OFFICE_CONF}" 2>/dev/null; then
    cat << EOF >> "${OFFICE_CONF}"

[jsaddons]
EnableJsAddons=true
JSPluginsServer=http://127.0.0.1:${BRIDGE_PORT}/jsplugins.xml
EOF
    chown "${TARGET_USER}" "${OFFICE_CONF}" 2>/dev/null || true
fi
echo "[✓] 已更新用户级 WPS 配置: ${OFFICE_CONF}"

# 3. 写入系统级 /opt/kingsoft/wps-office/office6/cfgs/oem.ini (开启 JsApiPlugin=true)
for oem_file in "/opt/kingsoft/wps-office/office6/cfgs/oem.ini" "/usr/lib/office6/cfgs/oem.ini"; do
    if [[ -f "${oem_file}" ]]; then
        if [[ -w "${oem_file}" ]]; then
            python3 - "${oem_file}" "${BRIDGE_PORT}" << 'PYEOF'
import sys, pathlib
oem_path = pathlib.Path(sys.argv[1])
port = sys.argv[2]
raw = oem_path.read_text(encoding="utf-8", errors="ignore")
lines = [l for l in raw.splitlines() if not l.strip().startswith(("JsApiPlugin=", "JsApiShowWebDebugger=", "JSPluginsServer="))]
out = []
inserted = False
for l in lines:
    out.append(l)
    if l.strip().lower() == "[support]" and not inserted:
        out.append("JsApiPlugin=true")
        out.append("JsApiShowWebDebugger=true")
        out.append(f"JSPluginsServer=http://127.0.0.1:{port}/jsplugins.xml")
        inserted = True
if not inserted:
    out.extend(["", "[Support]", "JsApiPlugin=true", "JsApiShowWebDebugger=true", f"JSPluginsServer=http://127.0.0.1:{port}/jsplugins.xml"])
oem_path.write_text("\n".join(out) + "\n", encoding="utf-8")
PYEOF
            echo "[✓] 已开启系统级 WPS 插件开关: ${oem_file}"
        else
            echo "[!] 提示: 修改 ${oem_file} 需要管理员权限，尝试使用 sudo 写入..."
            sudo python3 - "${oem_file}" "${BRIDGE_PORT}" << 'PYEOF' || true
import sys, pathlib
oem_path = pathlib.Path(sys.argv[1])
port = sys.argv[2]
raw = oem_path.read_text(encoding="utf-8", errors="ignore")
lines = [l for l in raw.splitlines() if not l.strip().startswith(("JsApiPlugin=", "JsApiShowWebDebugger=", "JSPluginsServer="))]
out = []
inserted = False
for l in lines:
    out.append(l)
    if l.strip().lower() == "[support]" and not inserted:
        out.append("JsApiPlugin=true")
        out.append("JsApiShowWebDebugger=true")
        out.append(f"JSPluginsServer=http://127.0.0.1:{port}/jsplugins.xml")
        inserted = True
if not inserted:
    out.extend(["", "[Support]", "JsApiPlugin=true", "JsApiShowWebDebugger=true", f"JSPluginsServer=http://127.0.0.1:{port}/jsplugins.xml"])
oem_path.write_text("\n".join(out) + "\n", encoding="utf-8")
PYEOF
            echo "[✓] 已通过 sudo 开启系统级 WPS 插件开关: ${oem_file}"
        fi
    fi
done

# 4. 立即在后台启动 127.0.0.1:38890 本地插件托管与桥接服务
pkill -f "xkl_gw.bridge_server" 2>/dev/null || true
pkill -f "serve-bridge" 2>/dev/null || true
export PYTHONPATH="${PROJECT_ROOT}/kylin-desktop-app:${PYTHONPATH}"
nohup python3 -m xkl_gw.bridge_server >/tmp/xkl-gw-bridge.log 2>&1 &
sleep 0.5
echo "[✓] 已在后台启动本地桥接服务 (http://127.0.0.1:${BRIDGE_PORT}/)"
echo "=================================================================="
echo "  WPS 内置【公文】功能区插件已激活！请彻底关闭所有 WPS 窗口后重新打开 WPS 文字。"
echo "=================================================================="
