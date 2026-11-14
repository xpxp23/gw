# 小恐龙公文排版助手 v2.1.0（银河麒麟 V10 · 海光 Hygon C86 3350 信创专版）

本工程是基于开源项目 **小恐龙公文排版助手 (`xkonglong/gw`)** 针对 **银河麒麟桌面操作系统 V10 / V10 SP1 (Kylin OS)** 与 **海光 Hygon C86 3350 处理器 (`x86_64` / `amd64` 架构)** 深度重构与完整移植的双引擎公文排版系统。

---

## 一、为什么原版项目无法在银河麒麟 + 海光 3350 上直接运行？

| 维度 | 原版项目 (`xkonglong/gw` v2.0.0.1) | 银河麒麟 V10 + 海光 3350 专版 (`v2.1.0-kylin-hygon3350`) |
| :--- | :--- | :--- |
| **宿主插件架构** | Windows 专属 `.NET Framework 4.8 VSTO` COM 加载项 (`Microsoft.Office.Tools.Word`) | **WPS Office Linux 原生 JS 加载项** (`ribbon.xml` + `wps.WpsApplication()` JSAPI) + **银河麒麟独立桌面端/CLI** |
| **UI 窗体界面** | Windows Forms (`System.Windows.Forms` + `Win32 Hwnd`) | **5 个原生 HTML5/CSS3 插件对话框** + **Tkinter/TTK 麒麟桌面端批处理工作台** |
| **42 个 VBA 脚本** | 依赖 Windows `Microsoft.Vbe.Interop` (`VBProject`) 动态注入且强制要求 ANSI/GB2312 编码 | **全量移植为 WPS JS 原生引擎 + Python OpenXML 批处理引擎**，并附带 UTF-8 / GB18030 双编码脚本库 |
| **公文部件插入** | 依赖 Windows Word COM `BuildingBlockEntries` 读取 `公文部件.dotx` | **原生 JSAPI 与 OpenXML 直接生成**单行/联合红头、2.25磅红反线、信函红线、竖排装订线、单/双行版记 |
| **语音朗读校对** | Windows `System.Speech.Synthesis` | **Web Speech API + 银河麒麟 `speech-dispatcher` (`spd-say`) / `espeak-ng` 双通道语音合成** |
| **副本导出与扫描件** | Windows GDI+ `EnhMetaFileBits` 及外部 `make-look-scanned.exe` | **原生 `pdftoppm` + `Pillow`/`NumPy` 渲染 PNG/JPG 与仿真扫描版 PDF**（无需任何 `.exe`） |
| **文件管理器联动** | 硬编码调用 `explorer.exe /select` 与 `notepad.exe` | 原生适配银河麒麟 **Peony (牡丹文件管理器)** / `xdg-open` 与 `xclip`/`wl-copy` |

---

## 二、目录结构说明

```text
xkl-gw-kylin-hygon3350/
├── wps-js-addin/                  # 【模式一】WPS Office Linux 信创版/专业版原生 JS 加载项
│   ├── ribbon.xml                 # 完整还原原版 7 大功能组“公文(麒麟版)”功能区配置
│   ├── index.html                 # 加载项主入口
│   ├── jsplugins.xml              # WPS Linux 加载项注册文件
│   ├── publish.xml                # WPS Linux 加载项发布清单
│   ├── js/
│   │   ├── wps_enum.js            # WPS Writer 枚举常量与单位换算兼容层
│   │   ├── config.js              # GB/T 9704—2012 默认参数存储与本地麒麟桥接通信
│   │   ├── ribbon.js              # Ribbon 按钮事件主分发器与图标映射
│   │   ├── layout.js              # 页面设置(37/35/28/26mm, 22行×28字)、缩进、顶格、页眉清理
│   │   ├── format.js              # 断行重排、空行清理、中英文空格、标点全半角互转、超链接清理
│   │   ├── fonts.js               # 公文主标题、正文、一~四级标题、着重号、隐藏文字、字体替换
│   │   ├── styles.js              # 公文专用样式集生成与强制重建
│   │   ├── insert.js              # 红头版头、红色反线、信函线、装订线、单/双行版记、符号、电子公章
│   │   ├── pagenumber.js          # 四号宋体 — 1 — 页码（居中 / 双面奇右偶左外侧 / 居右）
│   │   ├── shapes.js              # 图形/表格选择、环绕方式批量转换、分节拆分、打开 Peony 目录
│   │   ├── export.js              # 生成同名 PDF / DOC / PNG / JPG / 仿真扫描版 PDF 副本
│   │   ├── security_speech.js     # 文档只读保护、密码管理、麒麟系统双引擎中文语音朗读
│   │   └── vba_ported_scripts.js  # 原版 42 个 VBA 脚本的 WPS Linux 原生 JS 移植引擎
│   ├── ui/                        # 5 个可视化配置与脚本中心对话框
│   │   ├── paper_setup.html       # 公文页面设置面板 (对应原 frmPaperSetup)
│   │   ├── font_change.html       # 字体批量替换与数字/西文字体设置 (对应原 frmFontChangeDialog)
│   │   ├── speak_option.html      # 银河麒麟语音朗读设置 (对应原 frmSpeakOption)
│   │   ├── about_settings.html    # 关于与 10 类公文字体自定义高级设置 (对应原 AboutBox1)
│   │   └── vba_center.html        # 内置公文脚本中心可视化面板
│   ├── images/                    # 完整保留原项目 60+ 个功能区图标资源
│   └── resources/公文部件.dotx    # 原始公文部件模板备份
├── kylin-desktop-app/             # 【模式二】银河麒麟独立桌面端工作台、CLI 批处理与桥接守护进程
│   └── xkl_gw/
│       ├── __init__.py
│       ├── docx_engine.py         # 基于 python-docx + lxml 的 GB/T 9704 OpenXML 排版与批处理引擎
│       ├── bridge_server.py       # 监听 127.0.0.1:38890 的 WPS 加载项托管与麒麟系统级桥接服务
│       └── gui_app.py             # 银河麒麟桌面可视化批处理工作台 (GUI) 与命令行工具 (CLI)
├── vba-scripts/                   # 原版 42 个 VBA 脚本双编码归档
│   ├── utf8/                      # 转换为 UTF-8 编码，便于在银河麒麟下直接查阅与复用
│   └── gb18030/                   # 保持 GB18030/ANSI 编码，兼容传统宏环境
├── scripts/                       # 银河麒麟 + 海光 3350 专属安装、字体配置与 DEB 打包脚本
│   ├── install_kylin_hygon3350.sh # 一键安装脚本（自动注册 WPS 插件 + 安装桌面端 + 配置字体）
│   ├── setup_gb9704_fonts.sh      # GB/T 9704 公文字体 fontconfig 别名映射脚本
│   └── build_deb_hygon3350.sh     # 构建标准 amd64 (.deb) 软件包脚本
└── xkl-gw-assistant_2.1.0-kylin1_amd64.deb  # 预构建好的银河麒麟海光 3350 (amd64) 安装包
```

---

## 三、在银河麒麟 V10 (海光 3350) 上的安装与使用方法

### 方式 A：使用一键安装脚本（推荐，免 root 权限即可安装）

```bash
chmod +x scripts/install_kylin_hygon3350.sh
./scripts/install_kylin_hygon3350.sh
```

### 方式 B：使用预构建的 `.deb` 软件包安装（系统级安装，面向全用户）

```bash
sudo dpkg -i xkl-gw-assistant_2.1.0-kylin1_amd64.deb
sudo apt-get install -f
```

### 启动与使用

1. **在 WPS Office Linux 中使用**：
   - 确保后台桥接服务已运行（安装后登录桌面会自动启动，也可在终端执行 `xkl-gw-bridge &`）。
   - 打开 **WPS 文字 (Writer)**，顶部功能区会自动加载 **【公文(麒麟版)】** 选项卡，即可交互式使用页面设置、一键自动排版、红头版头、版记、页码、公章及脚本库。
2. **使用银河麒麟独立桌面批处理工作台**：
   - 双击桌面上的 **【小恐龙公文排版助手(麒麟版)】** 图标，或在终端执行：
     ```bash
     xkl-gw-gui
     ```
3. **使用命令行批量排版公文**：
   ```bash
   # 对单个或多个 .docx 文件执行 2026 国标公文自动排版，并同时插入红头、双行版记且导出 PDF 和仿真扫描件
   xkl-gw-cli 通知草案.docx --action auto2026 \
     --red-header "○○市人民政府办公室文件" \
     --doc-number "○政办发〔2026〕15号" \
     --banji double \
     --export-pdf --scanned-pdf

   # 制作 A4 横向双面对折会议桌签标牌
   xkl-gw-cli --place-cards "张　三,李四五,王　五" -o 会议桌签.docx

   # 数字转人民币大写
   xkl-gw-cli --rmb 128560.50
   ```
