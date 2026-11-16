## 小恐龙公文排版助手 for Word/WPS 免费开源

官网地址:  [https://gw.xkonglong.com](https://gw.xkonglong.com)

![小恐龙公文排版助手](https://gw.xkonglong.com/img/xklbanner.jpg)

根据《党政机关公文格式国家标准》(GB/T 9704—2012)的格式要求，快速设置页面版式、快速设置常用文字格式、插入常用符号、版头、版记、公文页码、红线等。

支持对文字的快速排版，中文首行缩进2字符、删除空白，断句重排等。

支持对文档盖章、可以在文档同目录下生成 PDF/ doc /图片 格式的副本。

可广泛用于各类规范文本的排版。

---

### 🐧 银河麒麟 V10 / SP1 · 海光 Hygon 3350 信创专版 (v2.2.0)

针对**银河麒麟桌面操作系统 V10 / V10 SP1**与**海光 Hygon C86 3350（x86_64 / amd64 架构）**处理器进行了深度重构与双引擎移植，并在 v2.2.0 全面解决 WPS 插件加载与界面兼容问题：

- 📂 **专版目录**：[`kylin-hygon3350/`](kylin-hygon3350/)
- 📖 **详细技术与更新文档**：[`kylin-hygon3350/README.md`](kylin-hygon3350/README.md)
- 📦 **预构建安装包**：`kylin-hygon3350/xkl-gw-assistant_2.2.0-kylin1_amd64.deb`
- 🚀 **主要架构特性与 v2.2.0 升级点**：
  1. **彻底解决 WPS 不显示【公文】选项卡**：
     - 重写符合 WPS Linux 严格 Schema 的 `ribbon.xml`，移除不兼容的 `<menuSeparator>`。
     - 自动向 `/opt/kingsoft/wps-office/office6/cfgs/oem.ini` 写入 `JsApiPlugin=true` 与本地插件服务器配置。
     - 部署时立即拉起本地 38890 插件托管与系统级桥接服务。
  2. **1:1 复刻 Windows 版功能区 + 「📌 WPS 顶部悬浮插件栏模式」**：
     - 桌面端直接内置 1:1 还原 Windows 版完整公文图标工具栏（7 大功能分组）。
     - 支持一键切换为置顶吸附在 WPS 顶部的悬浮插件栏，彻底无惧政企定制版 WPS 内部限制。
  3. **42 个公文自动化脚本全量移植**：支持 WPS JS 原生执行及 Python 跨平台引擎，并附带 UTF-8 / GB18030 双编码脚本库。
  4. **国标字体自动映射与 UI 自适应**：内置 `setup_gb9704_fonts.sh`，修复 Python 3.8 / Tk 8.6 下字号与按钮截断问题。

---

### Windows 原版说明

插件支持 VBA 脚本，可以在此处 `VBA Script` 文件夹查看和下载 VBA 脚本，也可到 [VBA脚本讨论组](https://github.com/xkonglong/gw/discussions)
VBA 脚本可以实现自动排版等复杂的且个性化的操作。
![vba](https://gw.xkonglong.com/img/screenshots.gif)

点此查看 [VBA指南](VBA指南.md)

本插件完全免费，承诺不含任何恶意代码。

从软件从2026年10月21日起开源 [MIT](MIT_License.md)
