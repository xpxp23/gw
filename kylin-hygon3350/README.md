# 小恐龙公文排版助手 v2.2.0（银河麒麟 V10 SP1 · 海光 Hygon C86 3350 信创增强版）

## 本次 v2.2.0 针对实机测试反馈的核心升级

1. **为什么之前打开 WPS 没有看到【公文】选项卡？已彻底解决！**
   - **原因 1（WPS Linux 系统级开关）**：银河麒麟下的 WPS Office (`/opt/kingsoft/wps-office/office6/cfgs/oem.ini`) 默认未开启 `JsApiPlugin=true` 与 `JSPluginsServer`，导致 WPS 启动时不读取 `~/.local/share/Kingsoft/wps/jsaddons/`。
   - **原因 2（WPS Linux XML Schema 校验）**：WPS Linux 的 `customUI` 解析器不支持 Windows VSTO 的 `<menuSeparator>` 标签，遇到该标签会静默跳过渲染整个选项卡。
   - **原因 3（后台 38890 托管服务未在安装时立即拉起）**：首次安装后尚未重启登录桌面时，`127.0.0.1:38890` 尚未启动。
   - **v2.2.0 解决方案**：
     - 重写了符合 WPS Linux 严格 Schema 且 **1:1 还原 Windows 版 7 大功能组布局** 的 `wps-js-addin/ribbon.xml` 与 `publish.xml`。
     - 新增 `scripts/enable_wps_linux_plugin.sh`（及界面右上角的 **「🔧 一键激活/修复 WPS 内置插件」** 按钮），自动向 `/opt/kingsoft/wps-office/office6/cfgs/oem.ini` 写入 `JsApiPlugin=true` 和 `JSPluginsServer=http://127.0.0.1:38890/jsplugins.xml`，同时覆盖 `~/.local/share/Kingsoft/wps/jsaddons` 与 `~/.kingsoft/wps/jsaddons` 双路径，并在安装时立即后台启动 `127.0.0.1:38890` 服务。

2. **桌面端新增 1:1 复刻 Windows 版的【公文】可视化功能区 + 「📌 WPS 顶部悬浮插件栏模式」**
   - 即使部分单位的政企定制版 WPS 裁减了内置 WebView 组件，您现在打开桌面【小恐龙公文排版助手(麒麟版)】，顶部直接内置了 **与 Windows 版完全一致的【公文】图标功能区**（版面、公文标题、文字格式、插入、其他、关于、选择/转换/脚本 全部 7 大分组及下拉菜单）。
   - 点击窗口右上角的 **「📌 切换为 WPS 顶部悬浮插件栏」**，窗口会自动收起底部面板并置顶吸附在 WPS 文字窗口上方，随点随用！

3. **修复 Python 3.8 Tk 8.6 字号兼容与右侧按钮截断问题**
   - 全面移除小数字号 `9.5`，统一采用严格整数字号（`8`、`9`、`10`、`13`）。
   - 底部快捷按钮区改为 3 列自适应网格 (`columnconfigure(weight=1)`)，彻底消除右侧文字截断。

---

## 快速升级/安装命令

```bash
chmod +x scripts/install_kylin_hygon3350.sh
./scripts/install_kylin_hygon3350.sh
```
（脚本会自动检测您现有的 `/data/usershare/APP/xkl-gw-kylin` 目录并原地升级，同时配置系统级 WPS `oem.ini` 插件开关并拉起后台服务。）
