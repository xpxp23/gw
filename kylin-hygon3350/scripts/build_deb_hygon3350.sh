#!/usr/bin/env bash
# ==============================================================================
# 构建银河麒麟桌面操作系统 V10 (海光 Hygon C86 3350 / amd64) 标准 .deb 安装包
# ==============================================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
BUILD_ROOT="/tmp/xkl-gw-deb-build"
PKG_NAME="xkl-gw-assistant_2.1.0-kylin1_amd64.deb"

rm -rf "${BUILD_ROOT}"
mkdir -p "${BUILD_ROOT}/DEBIAN"
mkdir -p "${BUILD_ROOT}/opt/xkl-gw-kylin"
mkdir -p "${BUILD_ROOT}/usr/bin"
mkdir -p "${BUILD_ROOT}/usr/share/applications"
mkdir -p "${BUILD_ROOT}/etc/xdg/autostart"

# 1. 复制项目源码与资源到 /opt/xkl-gw-kylin
cp -rf "${PROJECT_ROOT}/wps-js-addin" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -rf "${PROJECT_ROOT}/kylin-desktop-app" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -rf "${PROJECT_ROOT}/vba-scripts" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -rf "${PROJECT_ROOT}/scripts" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -f "${PROJECT_ROOT}/README.md" "${BUILD_ROOT}/opt/xkl-gw-kylin/"
cp -f "${PROJECT_ROOT}/MIT_License.md" "${BUILD_ROOT}/opt/xkl-gw-kylin/"

# 2. 创建系统级可执行命令 (/usr/bin/xkl-gw-cli, /usr/bin/xkl-gw-gui, /usr/bin/xkl-gw-bridge)
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

# 3. 桌面图标与自启动服务
cat << 'EOF' > "${BUILD_ROOT}/usr/share/applications/xkl-gw-assistant.desktop"
[Desktop Entry]
Name=小恐龙公文排版助手(麒麟版)
Comment=GB/T 9704 党政机关公文自动排版与批处理工具 (海光3350专版)
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

# 4. 编写 DEBIAN/control 与 postinst
cat << 'EOF' > "${BUILD_ROOT}/DEBIAN/control"
Package: xkl-gw-assistant
Version: 2.1.0-kylin1
Section: editors
Priority: optional
Architecture: amd64
Depends: python3 (>= 3.6)
Recommends: python3-docx, python3-lxml, python3-pil, python3-tk, poppler-utils, libreoffice-writer
Maintainer: xkonglong <xkonglong@qq.com>
Homepage: https://gw.xkonglong.com
Description: 小恐龙公文排版助手 (银河麒麟 V10 · 海光 Hygon C86 3350 专版)
 根据《党政机关公文格式国家标准》(GB/T 9704—2012) 快速设置页面版式、
 四级公文标题、红头版头、版记、公文页码、电子公章及 42 项公文自动化脚本。
 同时包含 WPS Office Linux 原生 JS 加载项与银河麒麟独立桌面批处理工作台。
EOF

cat << 'EOF' > "${BUILD_ROOT}/DEBIAN/postinst"
#!/usr/bin/env bash
set -e
# 为当前登录用户自动注册 WPS Office Linux 加载项并配置字体映射
for user_home in /home/* /root; do
    if [[ -d "${user_home}" ]]; then
        wps_dir="${user_home}/.local/share/Kingsoft/wps/jsaddons"
        mkdir -p "${wps_dir}/xkl_gw_addin_2.1.0"
        cp -rf /opt/xkl-gw-kylin/wps-js-addin/* "${wps_dir}/xkl_gw_addin_2.1.0/" 2>/dev/null || true
        cp -f /opt/xkl-gw-kylin/wps-js-addin/jsplugins.xml "${wps_dir}/jsplugins.xml" 2>/dev/null || true
        cp -f /opt/xkl-gw-kylin/wps-js-addin/publish.xml "${wps_dir}/publish.xml" 2>/dev/null || true
        owner="$(stat -c '%U:%G' "${user_home}" 2>/dev/null || echo 'root:root')"
        chown -R "${owner}" "${user_home}/.local/share/Kingsoft" 2>/dev/null || true
    fi
done
exit 0
EOF
chmod 0755 "${BUILD_ROOT}/DEBIAN/postinst"

dpkg-deb --build "${BUILD_ROOT}" "${PROJECT_ROOT}/${PKG_NAME}"
rm -rf "${BUILD_ROOT}"
echo "[成功] 已构建适用于银河麒麟 + 海光 3350 (amd64) 的 DEB 安装包: ${PROJECT_ROOT}/${PKG_NAME}"
