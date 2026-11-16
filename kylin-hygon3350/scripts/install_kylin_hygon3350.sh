#!/usr/bin/env bash
# ==============================================================================
# 小恐龙公文排版助手 v2.2.0 — 银河麒麟 V10 (海光 Hygon C86 3350 x86_64) 一键安装/升级脚本
# 自动识别已有安装路径 (/data/usershare/APP/xkl-gw-kylin 或 ~/.local/share/xkl-gw-kylin)
# ==============================================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

# 若用户此前已安装在 /data/usershare/APP/xkl-gw-kylin 且可写，则优先原地升级该目录
if [[ -d "/data/usershare/APP/xkl-gw-kylin" && -w "/data/usershare/APP/xkl-gw-kylin" ]]; then
    INSTALL_DIR="/data/usershare/APP/xkl-gw-kylin"
else
    INSTALL_DIR="${HOME}/.local/share/xkl-gw-kylin"
fi

BIN_DIR="${HOME}/.local/bin"
DESKTOP_DIR="${HOME}/.local/share/applications"
AUTOSTART_DIR="${HOME}/.config/autostart"

echo "=================================================================="
echo "  小恐龙公文排版助手 v2.2.0 (银河麒麟 V10 · 海光 3350 专版) 安装/升级程序"
echo "  目标安装路径: ${INSTALL_DIR}"
echo "=================================================================="

mkdir -p "${INSTALL_DIR}" "${BIN_DIR}" "${DESKTOP_DIR}" "${AUTOSTART_DIR}"

if [[ "${PROJECT_ROOT}" != "${INSTALL_DIR}" ]]; then
    cp -rf "${PROJECT_ROOT}/wps-js-addin" "${INSTALL_DIR}/"
    cp -rf "${PROJECT_ROOT}/kylin-desktop-app" "${INSTALL_DIR}/"
    cp -rf "${PROJECT_ROOT}/vba-scripts" "${INSTALL_DIR}/"
    cp -rf "${PROJECT_ROOT}/scripts" "${INSTALL_DIR}/"
fi

chmod +x "${INSTALL_DIR}/scripts/"*.sh

# 生成可执行启动脚本
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
Comment=GB/T 9704 党政机关公文自动排版与悬浮功能区插件 (海光3350专版 v2.2.0)
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

USER_DESKTOP="${HOME}/桌面"
[[ ! -d "${USER_DESKTOP}" && -d "${HOME}/Desktop" ]] && USER_DESKTOP="${HOME}/Desktop"
if [[ -d "${USER_DESKTOP}" ]]; then
    cp -f "${DESKTOP_DIR}/xkl-gw-assistant.desktop" "${USER_DESKTOP}/"
    chmod +x "${USER_DESKTOP}/xkl-gw-assistant.desktop"
fi

# 配置国标公文字体映射并激活 WPS Linux 内置插件
bash "${INSTALL_DIR}/scripts/setup_gb9704_fonts.sh"
bash "${INSTALL_DIR}/scripts/enable_wps_linux_plugin.sh"

echo "=================================================================="
echo "  v2.2.0 升级完成！"
echo "  1. 桌面端已内置 1:1 复刻 Windows 版的【公文】可视化功能区工具栏，"
echo "     支持点击右上角【📌 切换为 WPS 顶部悬浮插件栏】吸附在 WPS 顶部使用！"
echo "  2. WPS 内置加载项已同步激活（且后台 38890 桥接服务已立即拉起），"
echo "     完全关闭并重开 WPS 文字即可加载内置【公文】选项卡。"
echo "=================================================================="
