## 小恐龙公文排版助手 for Word/WPS 免费开源

官网地址:  [https://gw.xkonglong.com](https://gw.xkonglong.com)

![小恐龙公文排版助手](https://gw.xkonglong.com/img/xklbanner.jpg)

根据《党政机关公文格式国家标准》(GB/T 9704—2012)的格式要求，快速设置页面版式、快速设置常用文字格式、插入常用符号、版头、版记、公文页码、红线等。

支持对文字的快速排版，中文首行缩进2字符、删除空白，断句重排等。

支持对文档盖章、可以在文档同目录下生成 PDF/ doc /图片 格式的副本。

可广泛用于各类规范文本的排版。

---

### 🐧 新增：银河麒麟 V10 · 海光 Hygon 3350 信创专版 (v2.1.0)

针对**银河麒麟桌面操作系统 V10 / V10 SP1**与**海光 Hygon C86 3350（x86_64 / amd64 架构）**处理器进行了深度重构与双引擎移植：

- 📂 **专版目录**：[`kylin-hygon3350/`](kylin-hygon3350/)
- 📖 **详细技术与使用文档**：[`kylin-hygon3350/README.md`](kylin-hygon3350/README.md)
- 📦 **预构建安装包**：`kylin-hygon3350/xkl-gw-assistant_2.1.0-kylin1_amd64.deb`
- 🚀 **主要架构特性**：
  1. **WPS Office Linux 原生 JS 加载项**：完整移植原版 Ribbon 功能区（7个功能组、60+交互功能、5个原生 HTML5 对话框）。
  2. **42 个公文自动化脚本全量移植**：支持 WPS JS 原生执行及 Python 跨平台引擎，并附带 UTF-8 / GB18030 双编码脚本库。
  3. **银河麒麟独立桌面端与 CLI 批处理**：支持无需启动 WPS 即可通过图形界面或命令行执行国标排版、双行版记、红头版头、会议桌签与扫描版 PDF 生成。
  4. **国标字体自动映射**：内置 `setup_gb9704_fonts.sh`，自动匹配银河麒麟系统下的方正小标宋、仿宋、楷体、黑体等别名。

---

### Windows 原版说明

插件支持 VBA 脚本，可以在此处 `VBA Script` 文件夹查看和下载 VBA 脚本，也可到 [VBA脚本讨论组](https://github.com/xkonglong/gw/discussions)
VBA 脚本可以实现自动排版等复杂的且个性化的操作。
![vba](https://gw.xkonglong.com/img/screenshots.gif)

点此查看 [VBA指南](VBA指南.md)

本插件完全免费，承诺不含任何恶意代码。

从软件从2026年10月21日起开源 [MIT](MIT_License.md)
