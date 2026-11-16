#!/usr/bin/env bash
# ==============================================================================
# 构建银河麒麟桌面操作系统 V10 (海光 Hygon C86 3350 / amd64) 标准 .deb 安装包 (v2.2.0)
# ==============================================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
BUILD_ROOT="/tmp/xkl-gw-deb-build"
PKG_NAME="xkl-gw-assistant_2.2.0-kylin1_amd64.deb"

rm -rf "${BUILD_ROOT}" "${PROJECT_ROOT}"/*.deb
mkdir -p "${BUILD_ROOT}/DEBIAN"
mkdir -p "${BUILD_ROOT}/opt/xkl-gw-kylin"
mkdir -p "${BUILD_ROOT}/usr/bin"
mkdir -p "${BUILD_ROOT}/usr/share/applications"
mkdir -p "${BUILD_ROOT}/etc/xdg/autostart"

cp -rf "${PROJECT_ROOT}/wps-js-addin" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -rf "${PROJECT_ROOT}/kylin-desktop-app" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -rf "${PROJECT_ROOT}/vba-scripts" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -rf "${PROJECT_ROOT}/scripts" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -f "${PROJECT_ROOT}/README.md" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -f "${PROJECT_ROOT}/MIT_License.md" "${BUILD_ROOT}/opt/xkl-gw-kylin/"

cat << 'EOF' > "${BUILD_ROOT}/usr/bin/xkl-gw-cli"
#!/usr/bin/env bash
export PYTHONPATH="/opt/xkl-gw-kylin/kylin-desktop-app:${PYTHONPATH}"
exec python3 -m xkl_gw.gui_app "$@"
EOF
chmod 0755 "${BUILD_ROOT}/usr/bin/xkl-gw-cli"

cat << 'EOF' > "${BUILD_ROOT}/usr/bin/xkl-gw-gui"
#!/usr/bin/env bash
export PYTHONPATH="/opt/xkl-gw-kylin/kylin-desktop-app:${PYTHONPATH}"
exec python3 -m xkl_gw.gui_app --gui "$@"
EOF
chmod 0755 "${BUILD_ROOT}/usr/bin/xkl-gw-gui"

cat << 'EOF' > "${BUILD_ROOT}/usr/bin/xkl-gw-bridge"
#!/usr/bin/env bash
export PYTHONPATH="/opt/xkl-gw-kylin/kylin-desktop-app:${PYTHONPATH}"
exec python3 -m xkl_gw.gui_app --serve-bridge "$@"
EOF
chmod 0755 "${BUILD_ROOT}/usr/bin/xkl-gw-bridge"

cat << 'EOF' > "${BUILD_ROOT}/usr/share/applications/xkl-gw-assistant.desktop"
[Desktop Entry]
Name=小恐龙公文排版助手(麒麟版)
Comment=GB/T 9704 党政机关公文自动排版与悬浮功能区插件 (海光3350专版 v2.2.0)
Exec=/usr/bin/xkl-gw-gui
Icon=/opt/xkl-gw-kylin/wps-js-addin/images/OfficialTemplate.png
Terminal=false
Type=Application
Categories=Office;WordProcessor;
StartupNotify=true
EOF

cat << 'EOF' > "${BUILD_ROOT}/etc/xdg/autostart/xkl-gw-bridge.desktop"
[Desktop Entry]
Name=小恐龙公文助手WPS桥接服务
Comment=WPS Office Linux 加载项本地托管与系统桥接守护进程
Exec=/usr/bin/xkl-gw-bridge
Icon=/opt/xkl-gw-kylin/wps-js-addin/images/OfficialTemplate.png
Terminal=false
Type=Application
X-GNOME-Autostart-enabled=true
EOF

cat << 'EOF' > "${BUILD_ROOT}/DEBIAN/control"
Package: xkl-gw-assistant
Version: 2.2.0-kylin1
Section: editors
Priority: optional
Architecture: amd64
Depends: python3 (>= 3.6)
Recommends: python3-docx, python3-lxml, python3-pil, python3-tk, poppler-utils, libreoffice-writer, xclip, speech-dispatcher, espeak-ng
Maintainer: xkonglong <xkonglong@qq.com>
Homepage: https://gw.xkonglong.com
Description: 小恐龙公文排版助手 v2.2.0 (银河麒麟 V10 · 海光 Hygon C86 3350 专版)
 根据《党政机关公文格式国家标准》(GB/T 9704—2012) 快速设置页面版式、
 四级公文标题、红头版头、版记、公文页码、电子公章及 42 项公文自动化脚本。
 同时包含 WPS Office Linux 原生 JS 加载项与 1:1 复刻 Windows 功能区的顶部悬浮插件栏。
EOF

cat << 'EOF' > "${BUILD_ROOT}/DEBIAN/postinst"
#!/usr/bin/env bash
set -e
bash /opt/xkl-gw-kylin/scripts/enable_wps_linux_plugin.sh || true
exit 0
EOF
chmod 0755 "${BUILD_ROOT}/DEBIAN/postinst"

dpkg-deb --build "${BUILD_ROOT}" "${PROJECT_ROOT}/${PKG_NAME}"
rm -rf "${BUILD_ROOT}"
echo "[成功] 已构建适用于银河麒麟 + 海光 3350 (amd64) 的 DEB 安装包: ${PROJECT_ROOT}/${PKG_NAME}"
