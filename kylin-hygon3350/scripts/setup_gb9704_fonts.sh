#!/usr/bin/env bash
# ==============================================================================
# 银河麒麟桌面操作系统 V10 (海光 Hygon C86 3350) 国标公文字体映射配置脚本
# 解决 Linux 下《党政机关公文格式》(GB/T 9704—2012) 常见字体名称别名兼容问题：
#   - 方正小标宋简体 <=> 方正小标宋_GBK <=> FZXiaoBiaoSong-B05S
#   - 仿宋 <=> 仿宋_GB2312 <=> 方正仿宋_GBK <=> FangSong
#   - 楷体 <=> 楷体_GB2312 <=> 方正楷体_GBK <=> KaiTi
#   - 黑体 <=> 方正黑体_GBK <=> SimHei
#   - 宋体 <=> 方正书宋_GBK <=> SimSun
# ==============================================================================
set -e

FONTCONF_DIR="${HOME}/.config/fontconfig/conf.d"
mkdir -p "${FONTCONF_DIR}"

CONF_FILE="${FONTCONF_DIR}/99-xkl-gw-gb9704-fonts.conf"

cat << 'EOF' > "${CONF_FILE}"
<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
    <!-- 1. 公文标题：方正小标宋简体 / 方正小标宋_GBK -->
    <alias>
        <family>方正小标宋简体</family>
        <prefer>
            <family>方正小标宋简体</family>
            <family>方正小标宋_GBK</family>
            <family>FZXiaoBiaoSong-B05S</family>
            <family>FZXiaoBiaoSong-B05</family>
            <family>Noto Serif CJK SC</family>
            <family>SimSun</family>
        </prefer>
    </alias>
    <alias>
        <family>方正小标宋_GBK</family>
        <prefer>
            <family>方正小标宋_GBK</family>
            <family>方正小标宋简体</family>
            <family>FZXiaoBiaoSong-B05</family>
            <family>Noto Serif CJK SC</family>
        </prefer>
    </alias>

    <!-- 2. 公文正文与三级/四级标题：仿宋 / 仿宋_GB2312 -->
    <alias>
        <family>仿宋</family>
        <prefer>
            <family>仿宋</family>
            <family>仿宋_GB2312</family>
            <family>方正仿宋_GBK</family>
            <family>FangSong</family>
            <family>FangSong_GB2312</family>
            <family>FZFangSong-Z02S</family>
            <family>Noto Serif CJK SC</family>
        </prefer>
    </alias>
    <alias>
        <family>仿宋_GB2312</family>
        <prefer>
            <family>仿宋_GB2312</family>
            <family>仿宋</family>
            <family>方正仿宋_GBK</family>
            <family>FangSong_GB2312</family>
            <family>FangSong</family>
            <family>Noto Serif CJK SC</family>
        </prefer>
    </alias>

    <!-- 3. 二级标题与签发人：楷体 / 楷体_GB2312 -->
    <alias>
        <family>楷体</family>
        <prefer>
            <family>楷体</family>
            <family>楷体_GB2312</family>
            <family>方正楷体_GBK</family>
            <family>KaiTi</family>
            <family>KaiTi_GB2312</family>
            <family>FZKai-Z03S</family>
            <family>Noto Serif CJK SC</family>
        </prefer>
    </alias>
    <alias>
        <family>楷体_GB2312</family>
        <prefer>
            <family>楷体_GB2312</family>
            <family>楷体</family>
            <family>方正楷体_GBK</family>
            <family>KaiTi_GB2312</family>
            <family>KaiTi</family>
            <family>Noto Serif CJK SC</family>
        </prefer>
    </alias>

    <!-- 4. 一级标题与密级：黑体 -->
    <alias>
        <family>黑体</family>
        <prefer>
            <family>黑体</family>
            <family>方正黑体_GBK</family>
            <family>SimHei</family>
            <family>FZHei-B01S</family>
            <family>Noto Sans CJK SC</family>
        </prefer>
    </alias>

    <!-- 5. 页码字体：宋体 -->
    <alias>
        <family>宋体</family>
        <prefer>
            <family>宋体</family>
            <family>方正书宋_GBK</family>
            <family>SimSun</family>
            <family>FZShuSong-Z01S</family>
            <family>Noto Serif CJK SC</family>
        </prefer>
    </alias>
</fontconfig>
EOF

if command -v fc-cache >/dev/null 2>&1; then
    fc-cache -f >/dev/null 2>&1 || true
fi

echo "[完成] 国标公文字体映射规则已写入: ${CONF_FILE}"
