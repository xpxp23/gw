#!/usr/bin/env bash
# ==============================================================================
# 小恐龙公文排版助手 v2.1.0 — 银河麒麟 V10 (海光 Hygon C86 3350 x86_64) 一键部署脚本
# 支持：
#   1. 检测海光 Hygon C86 3350 (x86_64 / amd64) CPU 与银河麒麟 V10 环境
#   2. 部署 WPS Office Linux 信创版/专业版 JS 加载项 (含本地离线包与 publish.xml)
#   3. 部署银河麒麟独立桌面端工作台、命令行工具 (xkl-gw-cli / xkl-gw-gui) 与后台桥接服务
#   4. 配置 GB/T 9704—2012 国标公文字体映射与桌面快捷方式
# ==============================================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

INSTALL_DIR="${HOME}/.local/share/xkl-gw-kylin"
BIN_DIR="${HOME}/.local/bin"
WPS_ADDONS_DIR="${HOME}/.local/share/Kingsoft/wps/jsaddons"
DESKTOP_DIR="${HOME}/.local/share/applications"
AUTOSTART_DIR="${HOME}/.config/autostart"

echo "=================================================================="
echo "  小恐龙公文排版助手 v2.1.0 (银河麒麟 V10 · 海光 3350 专版) 安装程序"
echo "=================================================================="

# 1. 硬件架构与操作系统检测
ARCH="$(uname -m)"
CPU_MODEL="$(grep -m1 'model name' /proc/cpuinfo 2>/dev/null | cut -d: -f2 | xargs || echo 'x86_64 Processor')"
CPU_VENDOR="$(grep -m1 'vendor_id' /proc/cpuinfo 2>/dev/null | cut -d: -f2 | xargs || echo 'Unknown')"

echo "[1/5] 检测处理器与操作系统环境..."
echo "      架构: ${ARCH}"
echo "      型号: ${CPU_MODEL} (${CPU_VENDOR})"
if [[ "${ARCH}" != "x86_64" && "${ARCH}" != "amd64" ]]; then
    echo "      [提示] 当前架构为 ${ARCH}，本版本针对海光 3350 (x86_64) 优化，仍可兼容运行。"
fi

# 2. 拷贝核心文件到用户安装目录
echo "[2/5] 安装公文排版引擎与 42 个公文自动化脚本至 ${INSTALL_DIR} ..."
mkdir -p "${INSTALL_DIR}" "${BIN_DIR}" "${DESKTOP_DIR}" "${AUTOSTART_DIR}"
cp -rf "${PROJECT_ROOT}/wps-js-addin" "${INSTALL_DIR}/"
cp -rf "${PROJECT_ROOT}/kylin-desktop-app" "${INSTALL_DIR}/"
cp -rf "${PROJECT_ROOT}/vba-scripts" "${INSTALL_DIR}/"
cp -rf "${PROJECT_ROOT}/scripts" "${INSTALL_DIR}/"

# 3. 注册 WPS Office Linux 加载项
echo "[3/5] 注册 WPS Office Linux 原生 JS 加载项..."
mkdir -p "${WPS_ADDONS_DIR}/xkl_gw_addin_2.1.0"
cp -rf "${PROJECT_ROOT}/wps-js-addin/"* "${WPS_ADDONS_DIR}/xkl_gw_addin_2.1.0/"
cp -f "${PROJECT_ROOT}/wps-js-addin/jsplugins.xml" "${WPS_ADDONS_DIR}/jsplugins.xml"
cp -f "${PROJECT_ROOT}/wps-js-addin/publish.xml" "${WPS_ADDONS_DIR}/publish.xml"

# 4. 生成可执行启动脚本与桌面图标
echo "[4/5] 创建命令行入口 (xkl-gw-cli / xkl-gw-gui / xkl-gw-bridge) 与麒麟桌面图标..."

cat << EOF > "${BIN_DIR}/xkl-gw-cli"
#!/usr/bin/env bash
export PYTHONPATH="${INSTALL_DIR}/kylin-desktop-app:\${PYTHONPATH}"
exec python3 -m xkl_gw.gui_app "\$@"
EOF
chmod +x "${BIN_DIR}/xkl-gw-cli"

cat << EOF > "${BIN_DIR}/xkl-gw-gui"
#!/usr/bin/env bash
export PYTHONPATH="${INSTALL_DIR}/kylin-desktop-app:\${PYTHONPATH}"
exec python3 -m xkl_gw.gui_app --gui "\$@"
EOF
chmod +x "${BIN_DIR}/xkl-gw-gui"

cat << EOF > "${BIN_DIR}/xkl-gw-bridge"
#!/usr/bin/env bash
export PYTHONPATH="${INSTALL_DIR}/kylin-desktop-app:\${PYTHONPATH}"
exec python3 -m xkl_gw.gui_app --serve-bridge "\$@"
EOF
chmod +x "${BIN_DIR}/xkl-gw-bridge"

cat << EOF > "${DESKTOP_DIR}/xkl-gw-assistant.desktop"
[Desktop Entry]
Name=小恐龙公文排版助手(麒麟版)
Comment=GB/T 9704 党政机关公文自动排版与批处理工具 (海光3350专版)
Exec=${BIN_DIR}/xkl-gw-gui
Icon=${INSTALL_DIR}/wps-js-addin/images/OfficialTemplate.png
Terminal=false
Type=Application
Categories=Office;WordProcessor;
StartupNotify=true
EOF

cat << EOF > "${AUTOSTART_DIR}/xkl-gw-bridge.desktop"
[Desktop Entry]
Name=小恐龙公文助手WPS桥接服务
Comment=WPS Office Linux 加载项本地托管与系统桥接守护进程
Exec=${BIN_DIR}/xkl-gw-bridge
Icon=${INSTALL_DIR}/wps-js-addin/images/OfficialTemplate.png
Terminal=false
Type=Application
X-GNOME-Autostart-enabled=true
EOF

# 如果用户桌面目录存在，同时放置桌面快捷方式
USER_DESKTOP="${HOME}/桌面"
[[ ! -d "${USER_DESKTOP}" && -d "${HOME}/Desktop" ]] && USER_DESKTOP="${HOME}/Desktop"
if [[ -d "${USER_DESKTOP}" ]]; then
    cp -f "${DESKTOP_DIR}/xkl-gw-assistant.desktop" "${USER_DESKTOP}/"
    chmod +x "${USER_DESKTOP}/xkl-gw-assistant.desktop"
fi

# 5. 配置国标公文字体映射
echo "[5/5] 配置 GB/T 9704 国标公文字体别名映射..."
bash "${SCRIPT_DIR}/setup_gb9704_fonts.sh"

echo "=================================================================="
echo "  安装成功！"
echo "  1. WPS插件模式：重启 WPS 文字，顶部功能区将出现【公文(麒麟版)】选项卡"
echo "  2. 独立桌面模式：运行 ~/.local/bin/xkl-gw-gui 或双击桌面快捷方式"
echo "  3. 命令行批处理：运行 ~/.local/bin/xkl-gw-cli --help 查看参数"
echo "=================================================================="
